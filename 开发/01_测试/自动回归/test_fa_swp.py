# -*- coding: utf-8 -*-
"""固定资产 SWP（K1）判定器。

判定的是「数据层闭不闭合」，不开 Excel —— 重算那一步在交付前用 recalc.py 单独跑。

覆盖：TB 去重 / 期初反推 / 类别归集 / K.01 八个变动行 / 三张清单合计 / 上年回读 / 校验器。

**没有 sys.exit 就是假通过**（项目红线）。
"""

import os
import sys
import datetime as dt
from decimal import Decimal as D

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "程序"))

DB = os.path.join(ROOT, "tmp", "_v1280", "ledger.duckdb")

fails = []


def ck(name, cond, extra=""):
    print(("  ✓ " if cond else "  ✗ ") + name + (("  " + str(extra)) if extra and not cond else ""))
    if not cond:
        fails.append(name)


def main():
    if not os.path.isfile(DB):
        print("SKIP: 找不到瑞云测试账套 %s" % DB)
        return 0

    import duckdb
    from modules.fa_swp import prep, moves, synth, prior, checks

    conn = duckdb.connect(DB, read_only=True)

    # ---- 1) 数据准备
    p = prep.load(conn)
    ck("TB 去重删除 20 行", len(p["trace"]["dropped"]) == 20, len(p["trace"]["dropped"]))
    ck("瑞影期初反推 2 项", len(p["trace"]["adjusts"]) == 2, len(p["trace"]["adjusts"]))
    t = p["totals"]
    ck("原值期初 = 181,867,302.08", t["fa_beg"] == D("181867302.08"), t["fa_beg"])
    ck("原值期末 = 268,623,942.63", t["fa_end"] == D("268623942.63"), t["fa_end"])
    ck("折旧期初 = 83,670,982.81", t["dep_beg"] == D("83670982.81"), t["dep_beg"])
    ck("折旧期末 = 105,268,905.42", t["dep_end"] == D("105268905.42"), t["dep_end"])
    ck("减值期初 = 7,904,865.46", t["imp_beg"] == D("7904865.46"), t["imp_beg"])
    ck("减值期末 = 7,165,860.24", t["imp_end"] == D("7165860.24"), t["imp_end"])

    # ---- 2) 序时账分类
    m = moves.build(conn, p)

    def amt(sec, way):
        return m["totals"].get((sec, way), D("0"))

    # V1.2.94：原值段的增加方式从 2 个拆成 4 个（改造转入、转固冲回各成一行）
    ck("购置 2,989,060.54", amt("fa", "购置") == D("2989060.54"), amt("fa", "购置"))
    ck("在建工程转入 130,105,828.86", amt("fa", "在建工程转入") == D("130105828.86"),
       amt("fa", "在建工程转入"))
    ck("改造转入 8,715,059.16", amt("fa", "改造转入") == D("8715059.16"), amt("fa", "改造转入"))
    ck("转固冲回 -36,459,750.74", amt("fa", "转固冲回") == D("-36459750.74"),
       amt("fa", "转固冲回"))
    ck("处置或报废 -18,544,784.13", amt("fa", "处置或报废") == D("-18544784.13"),
       amt("fa", "处置或报废"))
    ck("内部重分类 -48,773.14", amt("fa", "内部重分类") == D("-48773.14"), amt("fa", "内部重分类"))
    ck("折旧计提 29,674,929.57", amt("dep", "计提") == D("29674929.57"), amt("dep", "计提"))
    ck("折旧转出 -8,077,006.96", amt("dep", "处置或报废") == D("-8077006.96"),
       amt("dep", "处置或报废"))
    ck("减值转销 -739,005.22", amt("imp", "处置或报废") == D("-739005.22"),
       amt("imp", "处置或报废"))
    ck("表4 = 计提", m["table4"]["diff"] == 0, m["table4"]["diff"])
    ck("表4 五个桶都有值或为 0", len(m["table4"]["buckets"]) == 5)
    ck("原值段计划 6 个方式", len(m["plan"]["fa"]) == 6, sorted(m["plan"]["fa"]))
    ck("折旧段计划 2 个方式（无改造/冲回）",
       m["plan"]["dep"] == {"计提", "处置或报废"}, sorted(m["plan"]["dep"]))

    # ---- 3) 三张清单
    syn = synth.build(p, m, dt.date(2026, 9, 30))
    fa = syn["fa"]
    ck("FA 期初原值 = K.01 年初",
       sum((r["fa_beg"] for r in fa), D(0)) == t["fa_beg"])
    ck("FA 期末原值 = K.01 年末",
       sum((r["fa_end"] for r in fa), D(0)) == t["fa_end"])
    ck("FA 期末折旧 = K.01 年末",
       sum((r["dep_end"] for r in fa), D(0)) == t["dep_end"])
    ck("变动清单 = 各增加方式",
       sum((r["orig"] for r in syn["moves"]), D(0))
       == sum((amt("fa", w) for w in moves.WAY_ORDER["fa"]
               if moves.WAY_DEFS[("fa", w)].kind == "add"), D("0")))
    ck("处置清单 = 各减少方式",
       sum((r["orig"] for r in syn["disposals"]), D(0))
       == -sum((amt("fa", w) for w in moves.WAY_ORDER["fa"]
                if moves.WAY_DEFS[("fa", w)].kind == "reduce"), D("0")))
    ck("清单方式含 4 个增加方式",
       {r["way"] for r in syn["moves"]} == {"购置", "在建工程转入", "改造转入", "转固冲回"},
       {r["way"] for r in syn["moves"]})
    ck("红字冲回行保留（原值为负）", any(r["add"] < 0 for r in fa))
    ck("清单行数 > 0", len(fa) > 0 and len(syn["moves"]) > 0 and len(syn["disposals"]) > 0,
       (len(fa), len(syn["moves"]), len(syn["disposals"])))

    # ---- 4) 校验器
    chk = checks.check(p, m, syn)
    ck("校验器全过（%d 项）" % len(chk["items"]), chk["ok"], chk["fails"])

    # ---- 5) 上年 K.01 往返
    out = os.path.join(ROOT, "tmp", "_regdata", "上年K.01_test.xlsx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    prior.build(p["cats"], "测试客户", "2025-12-31", out)
    back = prior.read(out)
    ok = all(abs(back[c]["fa"] - float(p["cats"][c]["fa_beg"])) < 0.005
             for c in p["cat_order"])
    ck("上年 K.01 回读 = 本年期初", ok)
    ck("上年 K.01 校验", checks.check_prior(p, back)["ok"])

    # ---- 6) 分类器边界
    ck("类别解析：竖线", prep.cat_of("机器设备|ZCLB02_SYS") == "机器设备")
    ck("类别解析：斜杠", prep.cat_of("ZCLB02_SYS/机器设备") == "机器设备")
    ck("类别解析：未知返回空", prep.cat_of("固定资产|1601") == "")

    # ---- 7) 嵌入层（V1.2.93）：上年页与 TB 页
    import openpyxl
    from modules.fa_swp import embed

    ck("sheet 名引用带单引号", embed.q("上年 K.01") == "'上年 K.01'")

    tpl = os.path.join(ROOT, "data", "bkd_tpl", "K1 SWP 固定资产 202YMMDD XYZ公司.xlsx")
    if os.path.isfile(tpl):
        wb = openpyxl.load_workbook(tpl)
        src_k = wb["K.01 Agree SL to GL"]
        src_l = wb["K.00 Lead Sheet"]
        want_cf_k = sum(len(r.rules) for r in src_k.conditional_formatting)
        want_cf_l = sum(len(r.rules) for r in src_l.conditional_formatting)

        r1 = embed.clone_sheet(wb, "K.01 Agree SL to GL", embed.SHEET_PRIOR_K01,
                               after="K.01 Agree SL to GL", expect_empty_cell="E12")
        r2 = embed.clone_sheet(wb, "K.00 Lead Sheet", embed.SHEET_PRIOR_LEAD,
                               after="K.00 Lead Sheet", expect_empty_cell="J47")
        ck("复制保留条件格式（K.01 %d / Lead %d）" % (want_cf_k, want_cf_l),
           r1["cf"] == want_cf_k and r2["cf"] == want_cf_l, (r1["cf"], r2["cf"]))
        ck("复制保留数据验证（Lead 3 条）", r2["dv"] == 3, r2["dv"])
        ck("复制保留合并区（5 / 20）", r1["merged"] == 5 and r2["merged"] == 20)
        sn = wb.sheetnames
        ck("上年页紧跟原页", sn[sn.index("K.01 Agree SL to GL") + 1] == embed.SHEET_PRIOR_K01)

        # 闸门：源页已有值时拒绝复制（防「先填后复制」把本年数抄成上年数）
        wb["K.01 Agree SL to GL"]["E12"] = 123.45
        try:
            embed.clone_sheet(wb, "K.01 Agree SL to GL", embed.SHEET_PRIOR_K01 + "x",
                              expect_empty_cell="E12")
            ck("复制闸门拦住「先填后复制」", False)
        except ValueError:
            ck("复制闸门拦住「先填后复制」", True)
        wb["K.01 Agree SL to GL"]["E12"] = None

        ws = wb[embed.SHEET_PRIOR_K01]
        saved = embed.clear_formulas(ws)
        ck("清公式：K.01 506 个全清", len(saved) == 506 and embed.count_formulas(ws) == 0,
           (len(saved), embed.count_formulas(ws)))
        ck("清公式前已快照（表2 =F10 还在）", saved.get((55, 2)) == "=F10", saved.get((55, 2)))
        wl = wb[embed.SHEET_PRIOR_LEAD]
        sl = embed.clear_formulas(wl)
        ck("清公式：Lead 40 个全清", len(sl) == 40 and embed.count_formulas(wl) == 0,
           (len(sl), embed.count_formulas(wl)))
        ck("assert_no_formula 对已清空页通过", embed.assert_no_formula(ws, "test") == 0)
    else:
        print("  · 跳过模板相关断言（找不到 %s）" % tpl)

    # ---- 8) 行结构工具（xlrows）：插行/公式平移/条件格式
    from modules.fa_swp import xlrows

    ck("平移：阈值以下的引用不动",
       xlrows.shift_formula("=E12+E13", 13, 2) == "=E12+E15", xlrows.shift_formula("=E12+E13", 13, 2))
    ck("平移：区间两端独立判定",
       xlrows.shift_formula("=SUM(E12:E17)", 15, 2) == "=SUM(E12:E19)",
       xlrows.shift_formula("=SUM(E12:E17)", 15, 2))
    # AB18 是 K.01 自己的格子，跟着走；'K.00 Lead Sheet'!I47 是别的表，一个字不改
    ck("平移：本表引用走、别的表引用不走",
       xlrows.shift_formula("=AB18-'K.00 Lead Sheet'!I47", 15, 2)
       == "=AB20-'K.00 Lead Sheet'!I47",
       xlrows.shift_formula("=AB18-'K.00 Lead Sheet'!I47", 15, 2))
    ck("平移：点名某表才动它的引用",
       xlrows.shift_formula("='K.01'!Z42", 40, 2, {"K.01"}) == "='K.01'!Z44",
       xlrows.shift_formula("='K.01'!Z42", 40, 2, {"K.01"}))
    ck("平移：跨表模式不碰本表引用",
       xlrows.shift_formula("=IF(C15=\"Minimal\",$C$4*100%)", 15, 2, {"K.01"},
                            same_sheet=False) == "=IF(C15=\"Minimal\",$C$4*100%)")
    ck("平移：数字常量不当引用",
       xlrows.shift_formula("=IFERROR(E33/E12,1)", 15, 2) == "=IFERROR(E35/E12,1)",
       xlrows.shift_formula("=IFERROR(E33/E12,1)", 15, 2))

    # 端到端：在模板上真插两行，验年末行、跨表引用、条件格式
    import openpyxl as _ox
    wb2 = _ox.load_workbook(tpl)
    ws2 = wb2["K.01 Agree SL to GL"]
    cf_before = sum(len(r.rules) for r in ws2.conditional_formatting)
    xlrows.apply_plan(wb2, ws2, [{"at": 15, "labels": ["改造转入", "转固冲回"]}])
    ck("插行后 C15/C16 是方式名", ws2["C15"].value == "改造转入" and ws2["C16"].value == "转固冲回",
       (ws2["C15"].value, ws2["C16"].value))
    ck("插行后原处置行下移到 17", ws2["C17"].value == "处置或报废", ws2["C17"].value)
    ck("新行骨架公式完整",
       ws2["G15"].value == "=SUM(E15:F15)" and str(ws2["Z16"].value).startswith("=SUM("),
       (ws2["G15"].value, ws2["Z16"].value))
    ck("条件格式条数不变（38）",
       sum(len(r.rules) for r in ws2.conditional_formatting) == cf_before, cf_before)
    ck("Lead E47 跟着指到 Z44",
       wb2["K.00 Lead Sheet"]["E47"].value == "='K.01 Agree SL to GL'!Z44",
       wb2["K.00 Lead Sheet"]["E47"].value)
    ck("Lead 自己的 D15 没被误改",
       "C15=" in str(wb2["K.00 Lead Sheet"]["D15"].value), wb2["K.00 Lead Sheet"]["D15"].value)
    xlrows.rewrite_plus_sum_row(ws2, 20, 12, 19)
    ck("年末行重写为连续区间",
       ws2["E20"].value == "=E12+E13+E14+E15+E16+E17+E18+E19", ws2["E20"].value)

    conn.close()
    print()
    if fails:
        print("失败 %d 项：%s" % (len(fails), "；".join(fails)))
        return 1
    print("全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
