# -*- coding: utf-8 -*-
"""固定资产 SWP · 类别动态化回归（插列 / 列块克隆 / 横向合计 / 条件格式）。

**只做内存级验证**：加载模板 → 假装有 8 个类别 → 走一遍完整的插列手术 →
断言结构。不落盘、不连数据库，几秒返回。

为什么要单独立一个文件：插列这套东西的失效模式是「合计仍平但少一块」，
端到端跑出来的底稿看不出来，必须逐格断言中间态才能防住。

跑法：
    程序/runtime/python/python.exe 01_测试/自动回归/test_fa_swp_cols.py
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "程序"))

import openpyxl                                          # noqa: E402

from modules.fa_swp import k01, xlcols                  # noqa: E402

TPL = os.path.join(ROOT, "data", "bkd_tpl", "K1 SWP 固定资产 202YMMDD XYZ公司.xlsx")

FAILS = []


def ck(name, cond, extra=""):
    if cond:
        print("  \u2713 %s" % name)
    else:
        print("  \u2717 %s  %s" % (name, extra))
        FAILS.append(name)


def main():
    if not os.path.isfile(TPL):
        print("找不到模板：%s" % TPL)
        return 2

    # ---------- 1) 公式按列平移 ----------
    print("公式按列平移：")
    ck("列号 >= at 的引用 +n",
       xlcols.shift_formula_cols("=E12+F12", 6, 3) == "=E12+I12")
    ck("区间两端独立判定",
       xlcols.shift_formula_cols("=SUM(W12:X12)", 26, 3) == "=SUM(W12:X12)")
    ck("点名某表才动它的引用",
       xlcols.shift_formula_cols("='K.01'!Z42", 26, 3, {"K.01"}) == "='K.01'!AC42")
    ck("跨表模式不碰本表引用",
       xlcols.shift_formula_cols("=SUM(E12,H12)", 26, 3, {"K.01"}, same_sheet=False)
       == "=SUM(E12,H12)")
    ck("数字常量不当引用",
       xlcols.shift_formula_cols("=IFERROR(AC12/AD12,1)", 20, 3) == "=IFERROR(AF12/AG12,1)")

    # ---------- 2) 插列：合并区 / 列宽 / 尾段 ----------
    print()
    print("插列手术：")
    wb = openpyxl.load_workbook(TPL)
    ws = wb["K.01 Agree SL to GL"]
    lay0 = k01.layout(ws)
    ck("模板有 7 个类别列块", len(lay0["cat_cols"]) == 7, lay0["cat_cols"])

    blocks = k01.plan_blocks(lay0, ["甲", "乙", "丙", "丁"])
    ck("4 类时不插列", blocks["n_new"] == 0 and blocks["insert_n"] == 0)
    ck("4 类时插入点在合计块首列（第 26 列）", blocks["insert_at"] == 26, blocks["insert_at"])

    blocks8 = k01.plan_blocks(lay0, list("ABCDEFGH"))
    ck("8 类时要插 3 列", blocks8["n_new"] == 1 and blocks8["insert_n"] == 3)
    ck("8 类时新块坐标 26/27/28",
       [blocks8["blocks"][7][k] for k in ("money", "name", "aud")] == [26, 27, 28])
    try:
        k01.plan_blocks(lay0, list("ABCDEFGHIJKLM"))
        ck("超过 12 类要报错", False)
    except ValueError:
        ck("超过 12 类要报错", True)

    before_merges = sorted(str(x) for x in ws.merged_cells.ranges)
    before_cf = sum(len(r.rules) for r in ws.conditional_formatting)
    ins = xlcols.insert_cols_keep(ws, 26, 3)
    after_merges = sorted(str(x) for x in ws.merged_cells.ranges)
    ck("横幅合并区搬到 AC9:AE9", "AC9:AE9" in after_merges and "Z9:AB9" not in after_merges,
       after_merges)
    ck("横幅文字跟着搬过去", str(ws["AC9"].value or "").startswith("|"), ws["AC9"].value)
    ck("其它合并区不动",
       [x for x in before_merges if x != "Z9:AB9"] == [x for x in after_merges if x != "AC9:AE9"])
    ck("条件格式条数不变（38）",
       sum(len(r.rules) for r in ws.conditional_formatting) == before_cf)
    ck("W 列宽保留（跨列维度没被拆坏）",
       ws.column_dimensions["W"].width == 11.84375 and ws.column_dimensions["W"].min == 23)
    ck("尾段维度重建为 35..16384",
       ws.column_dimensions["AI"].min == 35 and ws.column_dimensions["AI"].max == 16384)
    ck("新列继承原合计块的宽度",
       ws.column_dimensions["AC"].width == 11.3828125, ws.column_dimensions["AC"].width)
    ck("insert 报告合并区 1 处、尾段 1 个", ins["merges"] == 1 and ins["tail"] == 1, ins)

    # ---------- 3) 平移 + 克隆 ----------
    print()
    print("平移与克隆：")
    n_self = xlcols.shift_sheet_formulas_cols(ws, 26, 3)
    hits = xlcols.shift_workbook_col_refs(wb, ws.title, 26, 3)
    ck("本表平移 56 格", n_self == 56, n_self)
    ck("跨表引用改了 8 处", len(hits) == 8, len(hits))
    ck("Lead E47 指到 AC42",
       wb["K.00 Lead Sheet"]["E47"].value == "='K.01 Agree SL to GL'!AC42",
       wb["K.00 Lead Sheet"]["E47"].value)

    wmax = tuple(max((ws.column_dimensions[openpyxl.utils.get_column_letter(
        blocks8["blocks"][i][r])].width or 0) for i in range(8))
        for r in ("money", "name", "aud"))
    clone = xlcols.clone_block_cols(ws, 23, 26, range(1, ws.max_row + 1),
                                    value_rows_max=11, widths=wmax)
    ck("克隆 3 列 × 110 行", clone["cells"] > 300 and clone["formulas"] > 40, clone)
    ck("第 8 块类别名占位在中间列", ws["AA10"].value == "【…】", ws["AA10"].value)
    ck("第 8 块子表头齐",
       [ws[c].value for c in ("Z11", "AA11", "AB11")] == ["账面数", "账表调整/审计调整", "审定数"])
    ck("第 8 块审定 = 账面 + 调整", ws["AB12"].value == "=SUM(Z12:AA12)", ws["AB12"].value)
    ck("第 8 块年末是纵向加号", ws["AB18"].value == "=AB12+AB13+AB14+AB15+AB16+AB17")
    ck("第 8 块净值倒挤", ws["AB31"].value == "=AB12-AB19-AB25")
    ck("第 8 块 TB 差异行", ws["Z43"].value == "=Z18-Z42", ws["Z43"].value)
    ck("母本块没被改动", ws["Y12"].value == "=SUM(W12:X12)", ws["Y12"].value)

    # ---------- 4) 横向合计 ----------
    print()
    print("横向合计：")
    lay = k01.layout(ws)
    ck("插列后 layout 扫出 8 块", len(lay["cat_cols"]) == 8, lay["cat_cols"])
    tc = k01.total_cols(lay)
    ck("合计块在 AC/AD/AE", (tc["book"], tc["adj"], tc["aud"]) == (29, 30, 31), tc)

    roles = {"book": [b["money"] for b in blocks8["blocks"]],
             "adj": [b["name"] for b in blocks8["blocks"]],
             "aud": [b["aud"] for b in blocks8["blocks"]]}
    n_ac = xlcols.rewrite_across_sums(ws, roles)
    ck("重写 52 格横向合计", n_ac == 52, n_ac)
    ck("账面合计带上第 8 块",
       ws["AC12"].value == "=SUM(E12,H12,K12,N12,Q12,T12,W12,Z12)", ws["AC12"].value)
    ck("调整合计带上第 8 块",
       ws["AD12"].value == "=SUM(F12,I12,L12,O12,R12,U12,X12,AA12)", ws["AD12"].value)
    ck("审定合计带上第 8 块",
       ws["AE12"].value == "=SUM(G12,J12,M12,P12,S12,V12,Y12,AB12)", ws["AE12"].value)
    ck("TB 合计行也带上第 8 块",
       ws["AC44"].value == "=SUM(E44,H44,K44,N44,Q44,T44,W44,Z44)", ws["AC44"].value)
    ck("纵向合计不动", ws["AC18"].value == "=AC12+AC13+AC14+AC15+AC16+AC17")
    ck("CHECK 列跟着合计块走", ws["AF12"].value == "=AE12-'K.00 Lead Sheet'!J47",
       ws["AF12"].value)

    # ---------- 5) 条件格式重建 ----------
    print()
    print("条件格式：")
    segs = [(12, 17), (19, 23), (25, 29)]
    before = sum(len(r.rules) for r in ws.conditional_formatting)
    cf = xlcols.rebuild_block_cf(ws, 10, lay["cat_cols"], segs, first_cell="E12")
    ck("空值规则覆盖 8 块 × 3 段 = 24 个区间", cf["blank_ranges"] == 24, cf["blank_ranges"])
    ck("类别名规则覆盖 8 个格", cf["name_cells"] == 8, cf["name_cells"])
    ck("重写不新增规则（条数不变）",
       sum(len(r.rules) for r in ws.conditional_formatting) == before)

    found_blank = None
    for rng in ws.conditional_formatting:
        crs = xlcols._sqref_ranges(str(rng.sqref))
        if len(crs) == 24:
            found_blank = (str(rng.sqref), list(rng.rules))
    ck("找得到 24 区间的规则", found_blank is not None)
    if found_blank:
        cols = sorted({x[0] for x in xlcols._sqref_ranges(found_blank[0])})
        ck("覆盖到新块（第 26 列）", 26 in cols, cols)
        ck("公式仍指向块 1 的首格", found_blank[1][0].formula == ["LEN(TRIM(E12))=0"],
           found_blank[1][0].formula)

    n_tb_cf = xlcols.add_tb_cell_cf(ws, [26], [42, 44, 46, 48], ref_money=5)
    ck("新块补 4 条 TB 行空值规则", n_tb_cf == 4, n_tb_cf)

    # ---------- 6) 表2/表3 扩行 ----------
    print()
    print("表2 / 表3：")
    exp = k01.expand_table23(ws, wb, k01.layout(ws), blocks8["blocks"])
    ck("插 1 行到表2/表3", exp["extra"] == 1, exp)
    lay2 = k01.layout(ws)
    first = lay2["t2"] + 1
    sr = None
    for r in range(first, first + 40):
        if str(ws.cell(r, 2).value or "").strip() == "合计":
            sr = r
            break
    # 这个测试只插了列、没插行，所以表2 的合计从 62 下移到 63
    ck("合计行下移到第 63 行", sr == 63, sr)
    ck("表4 的合计跟随平移",
       ws.cell(sr + 18, 6).value == "=SUM(F%d:F%d)" % (sr + 13, sr + 17),
       ws.cell(sr + 18, 6).value)

    print()
    if FAILS:
        print("失败 %d 项：%s" % (len(FAILS), FAILS))
        return 1
    print("全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
