# -*- coding: utf-8 -*-
"""固定资产 SWP · 删空列/空行回归（三态平移 / 删列 / 删行 / 横向合计）。

**只做内存级验证**：加载模板 → 假装只有 4 个类别 → 删掉模板预留的 3 个空块 →
断言结构。不落盘、不连数据库。

为什么必须单独一个文件：删列的失效模式比插列更隐蔽 —— `delete_cols` 只搬格子和样式，
公式一个字符都不改，结果是**文件能打开、`#REF!` 数量为 0、所有差异恒 0**，
只有「合计块自己算成 0」这一条肉眼可见，而那通常要等到有人拿着底稿去对账。

跑法：
    程序/runtime/python/python.exe 01_测试/自动回归/test_fa_swp_shrink.py
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "程序"))

import openpyxl                                          # noqa: E402
from openpyxl.utils import get_column_letter as gl     # noqa: E402

from modules.fa_swp import k01, xlshrink as xs          # noqa: E402

TPL = os.path.join(ROOT, "data", "bkd_tpl", "K1 SWP 固定资产 202YMMDD XYZ公司.xlsx")

#: 4 个类别时的列块（模板前 4 块）与合计块
BLOCKS = [{"i": i, "money": 5 + 3 * i, "name": 6 + 3 * i, "aud": 7 + 3 * i,
           "cat": c} for i, c in enumerate(("机器设备", "电子设备", "其他设备", "运输设备"))]
ROLES = {"book": [b["money"] for b in BLOCKS],
         "adj": [b["name"] for b in BLOCKS],
         "aud": [b["aud"] for b in BLOCKS]}

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

    # ---------- 1) 三态标量 ----------
    print("三态平移（单格）：")
    ck("删除点左侧不动", xs._map_scalar(5, 17, 9) == 5)
    ck("删除区右侧左移 d", xs._map_scalar(26, 17, 9) == 17)
    ck("落在删除区 → None", xs._map_scalar(20, 17, 9) is None)
    ck("删除区上边界 → None", xs._map_scalar(17, 17, 9) is None)
    ck("删除区下边界外 → −d", xs._map_scalar(26, 17, 9) == 17)

    print("三态平移（区间）：")
    ck("整段在左侧不动", xs._map_interval(5, 7, 17, 9) == [(5, 7)])
    ck("整段在右侧左移", xs._map_interval(26, 28, 17, 9) == [(17, 19)])
    ck("整段被删 → 空", xs._map_interval(18, 22, 17, 9) == [])
    ck("跨界收缩", xs._map_interval(14, 20, 17, 9) == [(14, 16)],
       xs._map_interval(14, 20, 17, 9))
    ck("合计区间跨删除点收缩", xs._map_interval(55, 61, 59, 3) == [(55, 58)],
       xs._map_interval(55, 61, 59, 3))

    # ---------- 2) 三态公式 ----------
    print("三态公式：")
    # 旧的 7 项横向合计里带着 Q/T/W（马上要被删），必须先按新块列表重生成，
    # 才轮得到通用平移；直接平移就该抛 —— 这正是流程里「先 rewrite 再 shift」的原因。
    ck("横向合计直接平移 → 抛错（必须先重生成）",
       _raises(lambda: xs.shift_formula_del(
           "=SUM(E12,H12,K12,N12,Q12,T12,W12)", "col", 17, 9)))
    ck("删除区内的引用直接抛错",
       _raises(lambda: xs.shift_formula_del("=R10", "col", 17, 9)))
    ck("区间跨删除点收缩",
       xs.shift_formula_del("=SUM(C55:C61)", "row", 59, 3) == "=SUM(C55:C58)",
       xs.shift_formula_del("=SUM(C55:C61)", "row", 59, 3))
    ck("跨表引用指向被删列 → 抛错",
       _raises(lambda: xs.shift_formula_del(
           "='K.01 Agree SL to GL'!R10", "col", 17, 9,
           {"K.01 Agree SL to GL"}, same_sheet=False)))
    ck("同表模式不碰别的表的引用",
       xs.shift_formula_del("=AB12-'K.00 Lead Sheet'!I47", "col", 17, 9)
       == "=S12-'K.00 Lead Sheet'!I47",
       xs.shift_formula_del("=AB12-'K.00 Lead Sheet'!I47", "col", 17, 9))
    ck("#REF! 原样放行（模板自带 6 处）",
       xs.shift_formula_del("=SUM('FA list'!#REF!)", "col", 17, 9)
       == "=SUM('FA list'!#REF!)")

    # ---------- 3) 端到端：模板删 9 列 ----------
    print("删列手术：")
    wb = openpyxl.load_workbook(TPL)
    ws = wb["K.01 Agree SL to GL"]
    lay0 = k01.layout(ws)
    ck("模板有 7 个列块", len(lay0["cat_cols"]) == 7, len(lay0["cat_cols"]))

    plan = k01.plan_blocks(lay0, [b["cat"] for b in BLOCKS])
    ck("计划：删 9 列、插 0 列",
       plan["delete_n"] == 9 and plan["n_new"] == 0 and plan["delete_at"] == 17,
       (plan["delete_at"], plan["delete_n"], plan["n_new"]))
    ck("计划里的 blocks 只含实际类别（4 项）",
       len(plan["blocks"]) == 4 and all(b["cat"] for b in plan["blocks"]))

    r = xs.delete_cols_keep(ws, plan["delete_at"], plan["delete_n"])
    ck("列宽区间重建（删掉 9 条）", r["widths"] == 23, r["widths"])
    ck("合计横幅合并区搬到 Q9:S9",
       "Q9:S9" in sorted(str(x) for x in ws.merged_cells.ranges),
       sorted(str(x) for x in ws.merged_cells.ranges))

    tc = k01.total_cols(k01.layout(ws))
    ck("合计块落在 Q/R/S/T", (tc["book"], tc["adj"], tc["aud"], tc["check"]) == (17, 18, 19, 20),
       tc)

    # 横向合计先重生成，再通用平移 —— 顺序反了会在「引用落在删除区」上抛
    n = xs.rewrite_total_sums(ws, tc, ROLES)
    ck("横向合计重写 52 格", n == 52, n)
    ck("合计 = 4 块逐项", ws["Q12"].value == "=SUM(E12,H12,K12,N12)", ws["Q12"].value)

    # ---------- 4) 端到端：表2/表3 删 3 行 ----------
    print("表2 / 表3 删行：")
    lay1 = k01.layout(ws)
    exp = k01.expand_table23(ws, wb, lay1, plan["blocks"])
    ck("删 3 行、表2 合计落到第 59 行", exp["deleted"] == 3 and exp["sum_row"] == 59, exp)

    h_col = xs.shift_sheet_formulas_del(ws, "col", 17, 9)
    ck("列方向剩余平移 56 格", len(h_col) == 56, len(h_col))
    cf = xs.shift_cf_del(ws, "col", 17, 9)
    ck("CF 丢掉被删列的 24 段", cf["dropped"] == 24, cf)
    ck("CF 规则数 38 → 26", cf["rules"] == 26, cf["rules"])

    ck("合计行 SUM 区间收缩到 4 行",
       ws["C59"].value == "=SUM(C55:C58)", ws["C59"].value)
    ck("表3 合计行对新的合计块",
       ws["H59"].value == "=C59-Q18", ws["H59"].value)
    ck("表2 只剩 4 个类别行",
       [ws.cell(r, 2).value for r in range(55, 59)] ==
       ["=F10", "=I10", "=L10", "=O10"],
       [ws.cell(r, 2).value for r in range(55, 59)])
    ck("表4 上移到第 69 行", ws["B69"].value == "表4", ws["B69"].value)
    ck("表4 合计区间跟着缩",
       ws["F77"].value == "=SUM(F72:F76)", ws["F77"].value)
    ck("表4 计提引用新的合计块", ws["F78"].value == "=Q20", ws["F78"].value)

    # ---------- 5) 列宽与尾段 ----------
    print("列宽：")
    ck("合计列沿用原合计块宽度（11.38/18.15/16.84）",
       [round(ws.column_dimensions[gl(c)].width, 2) for c in (17, 18, 19)]
       == [11.38, 18.15, 16.84],
       [round(ws.column_dimensions[gl(c)].width, 2) for c in (17, 18, 19)])
    tail = [(k, d.min, d.max) for k, d in ws.column_dimensions.items()
            if d.max and d.max > 1000]
    ck("尾段维度跟着左移且 max 仍是 16384",
       tail == [("W", 23, 16384)], tail)
    ck("列块内宽度没被挪到左侧块上（H 仍 11.84）",
       round(ws.column_dimensions["H"].width, 2) == 11.84,
       ws.column_dimensions["H"].width)

    # ---------- 6) 结构断言 ----------
    print("结构：")
    lay2 = k01.layout(ws)
    ck("layout 扫出 4 个列块", len(lay2["cat_cols"]) == 4, lay2["cat_cols"])
    # 这一步还没填类别名，4 个真实块的表头仍写着 `【…】`（fill_table1 才改）。
    # 所以判据是「占位块只剩 4 个」，不是「一个都没有」——
    # 若删列没做，这里会是 7 个。
    ph = [c for c in range(5, 21) if "【" in str(ws.cell(lay2["cat_row"], c).value or "")]
    ck("占位块从 7 个减到 4 个（= 实际类别数）", len(ph) == 4, ph)
    ck("表1 列区间无空列",
       set(range(5, 21)) ==
       {b["money"] for b in BLOCKS} | {b["name"] for b in BLOCKS}
       | {b["aud"] for b in BLOCKS} | {17, 18, 19, 20})
    ck("年末行覆盖年初+全部变动行",
       ws.cell(lay2["sec"]["fa"]["end"], 5).value
       == "=" + "+".join("E%d" % r for r in range(lay2["sec"]["fa"]["beg"],
                                                  lay2["sec"]["fa"]["blank"] + 1)),
       ws.cell(lay2["sec"]["fa"]["end"], 5).value)

    print()
    if FAILS:
        print("失败 %d 项：%s" % (len(FAILS), "、".join(FAILS)))
        return 1
    print("全部通过")
    return 0


def _raises(fn):
    try:
        fn()
        return False
    except ValueError:
        return True


if __name__ == "__main__":
    sys.exit(main())
