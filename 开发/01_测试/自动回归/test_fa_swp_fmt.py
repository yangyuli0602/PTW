# -*- coding: utf-8 -*-
"""固定资产 SWP：数字格式归一 + Lead 调整链接 + 无批注 的判定器。

三组断言，都不依赖账套（只读模板，必要时临时 copy 一份）：

    A 规则   —— 哪些 number_format 算「金额系」要改、哪些不能碰
    B Lead   —— 调整汇总表五列表头、H 列是 SUMIFS、F 列仍是 0、幂等、缺标题要报错
    C 归一   —— 金额格统一、值不变、字体/填充/边框/列宽/合并/条件格式不变
    D 批注   —— 源码里不再出现 Comment(，生成件批注数为 0

**没有 sys.exit 就是假通过**（项目红线）。
"""

import inspect
import os
import shutil
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PROJ = os.path.abspath(os.path.join(ROOT, ".."))   # 平台根：data / tmp 在这层
sys.path.insert(0, os.path.join(ROOT, "程序"))

TEMPLATE = os.path.join(PROJ, "data", "bkd_tpl", "K1 SWP 固定资产 202YMMDD XYZ公司.xlsx")
LATEST = os.path.join(PROJ, "tmp", "_v1280", "bkd_out", "固定资产SWP_最新.xlsx")

fails = []


def ck(name, cond, extra=""):
    print(("  ✓ " if cond else "  ✗ ") + name + (("  " + str(extra)) if extra and not cond else ""))
    if not cond:
        fails.append(name)


def main():
    import openpyxl
    from core.xlsx_fmt import ACCT_FMT, is_money_format
    from modules.fa_swp import fmt, k01, lead

    # ---------------- A) 格式判定规则
    print("A 规则：")
    yes = [
        ACCT_FMT,
        "#,##0.00;[Red]-#,##0.00",                                   # 代码旧写法
        "#,##0.00",
        r'_(* #,##0.00_);_(* \(#,##0.00);_(* "-"??_);_(@_)',          # 模板「会计专用」
        r'#,##0.00_ ;[Red]\-#,##0.00\ ',
        r'#,##0;\(#,##0\);\-#,##0',                                   # 整数千分位也算金额
        r'_-* #,##0.00_-;\-* #,##0.00_-;_-* "-"??_-;_-@_-',
        '"¥"#,##0',
        '_ [$€-2]* #,##0.00_ ;_ [$€-2]* \\-#,##0.00_ ',
    ]
    no = [
        "General", "", "0", "#", "@",
        "yyyy-mm-dd", "hh:mm:ss", "mmm-yy", "yyyy/m/d",
        "0.00%", "0%", "0.0000", '0.00"万元"',
    ]
    bad = [f for f in yes if not is_money_format(f)] + [f for f in no if is_money_format(f)]
    ck("金额系 %d 串全部判为改" % len(yes), not [f for f in yes if not is_money_format(f)])
    ck("非金额系 %d 串全部判为不改" % len(no), not [f for f in no if is_money_format(f)], bad[:4])
    ck("ACCT_FMT 零段是短横线", '"-"' in ACCT_FMT)
    ck("ACCT_FMT 负数用括号", "(#,##0.00)" in ACCT_FMT)
    ck("ACCT_FMT 不带币种符号", not any(x in ACCT_FMT for x in ("¥", "$", "€")))

    if not os.path.isfile(TEMPLATE):
        print("SKIP: 找不到模板 %s" % TEMPLATE)
        print("\n失败 %d 项" % len(fails))
        return 1

    # ---------------- B) Lead 调整汇总表
    print("B Lead 调整汇总表：")
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb[lead.SHEET]
    prior_cats = {"机器设备": {"fa": 100.0, "dep": 10.0, "imp": 0.0},
                  "电子设备": {"fa": 200.0, "dep": 20.0, "imp": 0.0}}
    tab = lead.fill_table1(ws, prior_cats, "x.xlsx", prior_sheet=None)
    adj = lead.fill_adj_summary(ws, tab)

    lay = lead.locate_table1(ws)
    hdr = adj["header"]
    heads = [ws.cell(hdr, c).value for c in range(2, 7)]
    ck("五列表头 = 调整类型/主体/调整说明/科目名称/调整金额",
       heads == ["调整类型", "主体", "调整说明", "科目名称", "调整金额"], heads)

    hcol = lay["cols"]["adj_aud"]
    rows = list(lay["rows"].values())
    forms = [str(ws.cell(r, hcol).value or "") for r in rows]
    ck("H 列三行都是 SUMIFS 引用汇总表",
       all(f.startswith("=SUMIFS($F$%d:$F$" % adj["d0"]) for f in forms), forms[:2])
    ck("SUMIFS 条件值取 C 列科目名称（不是 D 列）",
       all(("$C%d)" % r) in f for r, f in zip(rows, forms)), forms[:2])
    ck("SUMIFS 金额列是 F、条件列是 E",
       all(("$E$%d:$E$%d" % (adj["d0"], adj["d1"])) in f for f in forms))

    fcol = lay["cols"]["adj_bs"]
    ck("F 列（账表调整）仍是常数 0",
       all(ws.cell(r, fcol).value in (0, 0.0) for r in rows))
    ic = lay["cols"]["aud"]
    ck("I 列审定数仍是模板公式 G+H",
       all(str(ws.cell(r, ic).value or "").startswith("=G%d+H%d" % (r, r)) for r in rows),
       [ws.cell(r, ic).value for r in rows])

    # 幂等：再调一次不应改变任何东西
    snap = [(r, ws.cell(r, hcol).value, ws.cell(r, fcol).value) for r in rows]
    lead.fill_adj_summary(ws, lead.fill_table1(ws, prior_cats, "x.xlsx"))
    same = [(r, ws.cell(r, hcol).value, ws.cell(r, fcol).value) for r in rows] == snap
    ck("连调两次结果一致（幂等）", same)

    # 数据区必须空：有残留数字会让 SUMIFS 算出非零、审定数悄悄变
    left = [(r, ws.cell(r, 6).value) for r in range(adj["d0"], adj["d1"] + 1)
            if ws.cell(r, 6).value not in (None, "")]
    ck("调整数据区为空", not left, left[:3])

    # 找不到标题行要报错，不能静默退回常数 0
    ws2 = wb[lead.SHEET]
    for r in range(40, 120):
        if "调整汇总表" in str(ws2.cell(r, 2).value or ""):
            ws2.cell(r, 2).value = "（被删）"
            break
    try:
        lead.fill_adj_summary(ws2, lay)
        ck("缺「调整汇总表」标题要报错", False, "没有报错")
    except ValueError as e:
        ck("缺「调整汇总表」标题要报错", "调整汇总表" in str(e), e)

    # ---------------- C) 全簿格式归一
    print("C 格式归一：")
    wb2 = openpyxl.load_workbook(TEMPLATE)
    before = {}
    for name in wb2.sheetnames:
        s = wb2[name]
        before[name] = {c.coordinate: (c.value, c.font.name, c.font.size,
                                       str(c.fill.fgColor.rgb), str(c.border.left.style),
                                       c.number_format)
                        for c in s._cells.values()}
    widths = {n: dict((k, v.width) for k, v in wb2[n].column_dimensions.items())
              for n in wb2.sheetnames}
    merges = {n: sorted(map(str, wb2[n].merged_cells.ranges)) for n in wb2.sheetnames}
    cf = {n: len(wb2[n].conditional_formatting) for n in wb2.sheetnames}

    tr = fmt.normalize_money(wb2)
    # 模板里没有「上年 K.01 / 上年 Lead / TB 页」——这三张是程序建的，属正常缺失。
    dynamic = {"上年 K.00 Lead Sheet", "上年 K.01 Agree SL to GL", "TB 固定资产"}
    ck("模板上缺页只有程序生成的这三张", set(tr["missing"]) <= dynamic, tr["missing"])
    ck("模板上应扫的页一张不漏", tr["sheets"] >= 12, tr["sheets"])
    ck("确实改了金额格（>500）", tr["cells"] > 500, tr["cells"])

    rest = [(n, c.number_format) for n in wb2.sheetnames for c in wb2[n]._cells.values()
            if c.number_format != ACCT_FMT and is_money_format(c.number_format)]
    ck("全簿无 money 系残留格式", not rest, rest[:3])

    changed = []
    for name in wb2.sheetnames:
        for c in wb2[name]._cells.values():
            old = before[name][c.coordinate]
            if (c.value, c.font.name, c.font.size, str(c.fill.fgColor.rgb),
                    str(c.border.left.style)) != (old[0], old[1], old[2], old[3], old[4]):
                changed.append((name, c.coordinate))
    ck("归一不动值 / 字体 / 填充 / 边框", not changed, changed[:5])

    ck("列宽未变", {n: dict((k, v.width) for k, v in wb2[n].column_dimensions.items())
                  for n in wb2.sheetnames} == widths)
    ck("合并区未变", {n: sorted(map(str, wb2[n].merged_cells.ranges))
                    for n in wb2.sheetnames} == merges)
    ck("条件格式条数未变", {n: len(wb2[n].conditional_formatting)
                      for n in wb2.sheetnames} == cf)

    # 不该被改的几类
    k = wb2[k01.SHEET_K01]
    kept = []
    for c in k._cells.values():
        f = c.number_format
        if f in ("yyyy-mm-dd", "0.00%", "@", "General", "0") and f != ACCT_FMT:
            kept.append(f)
    ck("日期/百分比/文本/序号格式保持原样", bool(kept), sorted(set(kept)))
    ck("归一后可重复执行且不再变动",
       fmt.normalize_money(wb2)["cells"] == 0, fmt.normalize_money(wb2)["cells"])

    # ---------------- D) 批量批注
    print("D 批注：")
    src = inspect.getsource(k01) + inspect.getsource(lead)
    ck("源码里不再出现 Comment(", "Comment(" not in src)
    ck("源码里不再 import openpyxl.comments", "openpyxl.comments" not in src)
    if os.path.isfile(LATEST):
        wb3 = openpyxl.load_workbook(LATEST)
        n = sum(1 for s in wb3.worksheets for c in s._cells.values() if c.comment)
        ck("生成件批注数为 0", n == 0, n)
        l3 = wb3[lead.SHEET]
        ck("生成件 Lead 五列表头齐全",
           [l3.cell(63, c).value for c in range(2, 7)] ==
           ["调整类型", "主体", "调整说明", "科目名称", "调整金额"],
           [l3.cell(63, c).value for c in range(2, 7)])
    else:
        print("  · 跳过生成件检查（%s 不存在）" % LATEST)

    print()
    if fails:
        print("失败 %d 项：%s" % (len(fails), "、".join(fails)))
        return 1
    print("全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
