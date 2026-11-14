# -*- coding: utf-8 -*-
"""Excel 行结构手术：插行、公式平移、条件格式平移。

**为什么必须自己做，而不是用 `ws.insert_rows()` 一把梭**：

`insert_rows` 只搬单元格的值与样式，**公式一个字符都不改**。K.01 的年末行是
`E18=E12+E13+E14+E15+E16+E17`，在 15 行前插两行后，这一格被推到 20 行，
内容还是 `=E12+…+E17` —— 新插的两个方式行**不进年末合计**。
而合计仍然平、报表仍然能打开，肉眼极难发现。这是本次改造最危险的失效模式，
所以「平移公式」和「重写年末行」是这套工具的必选项，不是可选优化。

同理还丢三样：合并区不重挂、行高不搬、条件格式的 sqref 不动（K.01 有 38 条，
其中一条挂了 21 个区间，是「空值涂黄」的告警，丢了新行就不会被标黄）。

**这套工具不认识 K.01 的业务语义**（段落、方式、类别都不认），
只吃 `ws / at / n / 公式字符串`，将来别的底稿要用可以整体搬走。
"""

from __future__ import annotations

import re
from copy import copy

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils import get_column_letter

#: A1 或 A1:B2，可带表名前缀（含引号与 '' 转义）
_REF = re.compile(
    r"^(?:(?:'(?P<q>(?:[^']|'')+)'|(?P<u>[^'!]+))!)?"
    r"(?P<a1>\$?[A-Z]{1,3}\$?\d+)(?::(?P<a2>\$?[A-Z]{1,3}\$?\d+))?$"
)
_A1 = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")


# ------------------------------------------------------------------ 公式平移

def _shift_a1(ref: str, at: int, n: int) -> str:
    m = _A1.match(ref)
    if not m:
        raise ValueError("认不出的单元格引用：%r" % ref)
    c_abs, col, r_abs, row = m.groups()
    r = int(row)
    if r >= at:
        r += n
    return "%s%s%s%d" % (c_abs, col, r_abs, r)


def shift_formula(formula: str, at: int, n: int,
                  target_sheets: set[str] | None = None,
                  same_sheet: bool = True) -> str:
    """把公式里「行号 >= at」的引用 +n。

    ``same_sheet``：True 时改「无表名前缀」的引用（本表自己）；False 时**只改**
    带 ``target_sheets`` 前缀的引用。
    这两个开关必须分开：改别的表里指向 K.01 的引用时，那张表自己的
    ``=IF(C15=…)`` 是它本表的格子，跟着 +2 就把它自己的表弄乱了。

    只改 ``OPERAND`` + ``subtype == "RANGE"``。数字（``IFERROR(x,1)`` 里的 1）、
    字符串里的 "A1" 都不能碰，这也是必须上 Tokenizer 而不能字符串 replace 的原因。
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    out = []
    for tok in Tokenizer(formula).items:
        if tok.type == "OPERAND" and tok.subtype == "RANGE":
            m = _REF.match(tok.value)
            if not m:
                # 模板里本来就有失效引用（如 `'FA list'!#REF!`），原样留着，
                # 那是模板自身的错误（74 个错误格基线的一部分），不归本工具管。
                if "#REF" in tok.value:
                    out.append(tok.value)
                    continue
                # 整列 A:A、3D 引用、定义名、结构化引用 —— 直接炸，不静默跳过
                raise ValueError("公式里有不支持的引用：%r（%s）" % (tok.value, formula))
            sheet = None
            if m["q"] is not None:
                sheet = m["q"].replace("''", "'")
            elif m["u"] is not None:
                sheet = m["u"]
            hit = (same_sheet and sheet is None) or \
                  (target_sheets is not None and sheet in target_sheets)
            if hit:
                a1 = _shift_a1(m["a1"], at, n)
                a2 = _shift_a1(m["a2"], at, n) if m["a2"] else ""
                pref = tok.value[:tok.value.index(m["a1"])]
                # Token 的签名是 (value, type, subtype)，别按直觉写成 type 在前
                tok = Token(pref + a1 + (":" + a2 if a2 else ""), tok.type, tok.subtype)
            out.append(tok.value)
        else:
            out.append(tok.value)
    return "=" + "".join(out)


def shift_sheet_formulas(ws, at: int, n: int, target: str | None = None) -> int:
    """本表所有公式过一遍平移。``target`` 给了就同时改指向该表的引用。"""
    sheets = {target} if target else None
    hit = 0
    for row in ws.iter_rows():
        for c in row:
            if c.data_type != "f":
                continue
            new = shift_formula(c.value, at, n, sheets)
            if new != c.value:
                c.value = new
                hit += 1
    return hit


def shift_workbook_refs(wb, target: str, at: int, n: int) -> list[tuple]:
    """其它 sheet 里指向 ``target`` 且行号 >= at 的引用全部 +n。

    返回 ``[(sheet, coord, before, after)]``，写进 trace 并可断言。
    """
    hits = []
    for ws in wb.worksheets:
        if ws.title == target:
            continue
        for row in ws.iter_rows():
            for c in row:
                if c.data_type != "f":
                    continue
                new = shift_formula(c.value, at, n, {target}, same_sheet=False)
                if new != c.value:
                    hits.append((ws.title, c.coordinate, c.value, new))
                    c.value = new
    return hits


# ------------------------------------------------------------------ 条件格式

def _shift_ranges(sqref, at: int, n: int) -> str:
    """一个 sqref（可能含多个空格分隔区间）→ 平移后的字符串。"""
    out = []
    for cr in sqref.ranges:
        if cr.max_row < at:                 # 整块在上方
            r1, r2 = cr.min_row, cr.max_row
        elif cr.min_row >= at:              # 整块在下方
            r1, r2 = cr.min_row + n, cr.max_row + n
        else:                               # 插入点落在区间内部：撑开
            r1, r2 = cr.min_row, cr.max_row + n
        out.append("%s%d:%s%d" % (get_column_letter(cr.min_col), r1,
                                  get_column_letter(cr.max_col), r2))
    return " ".join(out)


def shift_cf(ws, at: int, n: int) -> int:
    """条件格式区间平移（落点三态，见 _shift_ranges）。

    多个区间必须写回**同一条规则**（空格分隔）—— K.01 那条「空值涂黄」
    挂了 21 个区间，拆成 21 条会让规则重复 21 份、文件体积翻倍。
    """
    from openpyxl.formatting.formatting import ConditionalFormattingList

    keep = []
    for rng in list(ws.conditional_formatting):
        keep.append((_shift_ranges(rng.sqref, at, n), list(rng.rules)))
    ws.conditional_formatting = ConditionalFormattingList()
    for sq, rules in keep:
        for rule in rules:
            ws.conditional_formatting.add(sq, rule)
    return sum(len(rules) for _sq, rules in keep)


def shift_dv(ws, at: int, n: int) -> int:
    """数据验证区间平移（K.01 没有，但模板换版可能有）。"""
    hit = 0
    for dv in ws.data_validations.dataValidation:
        dv.sqref = _shift_ranges(dv.sqref, at, n)
        hit += 1
    return hit


# ------------------------------------------------------------------ 插行

def insert_rows_keep(ws, at: int, n: int) -> dict:
    """在 ``at`` 前插 n 行，并把合并区/行高/CF/DV 都摆正。

    顺序不能反：**先 unmerge 再 insert**。插入点若落在合并区内部，
    新行的格子会变成 MergedCell，写值直接抛异常或静默丢。
    """
    # 1) 备份行高（insert_rows 完全不搬 row_dimensions）
    dims = {r: {"height": d.height, "custom": d.customHeight, "hidden": d.hidden}
            for r, d in list(ws.row_dimensions.items()) if r >= at}

    # 2) 拆掉跨过插入点 / 在插入点下方的合并区
    moved = []
    for mc in list(ws.merged_cells.ranges):
        if mc.max_row >= at:
            moved.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
            ws.unmerge_cells(str(mc))

    # 3) 插行
    ws.insert_rows(at, n)

    # 4) 重挂合并区
    for r1, r2, c1, c2 in moved:
        ws.merge_cells(start_row=r1 + n if r1 >= at else r1, start_column=c1,
                       end_row=r2 + n if r2 >= at else r2, end_column=c2)

    # 5) 行高回填（自下而上，避免被自动创建的维度干扰）
    for r in sorted(dims, reverse=True):
        d = ws.row_dimensions[r + n]
        # customHeight 是只读 property，设 height 时 openpyxl 会自己置位
        d.height = dims[r]["height"]
        d.hidden = dims[r]["hidden"]

    # 6) 条件格式 / 数据验证
    shift_cf(ws, at, n)
    shift_dv(ws, at, n)

    return {"at": at, "n": n, "merges": len(moved), "dims": len(dims)}


def copy_row_style(ws, src_row: int, dst_row: int, c1: int = 1, c2: int = 0) -> None:
    """把 src_row 的样式（含行高）刷到 dst_row，不搬值与公式。"""
    c2 = c2 or ws.max_column
    for c in range(c1, c2 + 1):
        s, d = ws.cell(src_row, c), ws.cell(dst_row, c)
        if s.has_style:
            d._style = copy(s._style)
    if src_row in ws.row_dimensions and ws.row_dimensions[src_row].height:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


# ------------------------------------------------------------------ 行内容

def rewrite_plus_sum_row(ws, row: int, r1: int, r2: int, cols=range(5, 29)) -> list[str]:
    """把 row 行写成 ``=X{r1}+X{r1+1}+…+X{r2}``（模板原句式，不用 SUM 区间）。

    保持加号句式是为了让审阅人能一眼看出「覆盖了哪几行」——
    这次改造最危险的失效是「合计仍平但少一行」，加号句式肉眼可查。

    ``r1`` 要传**年初行**（模板原文就是 ``E18=E12+E13+…+E17``，含 E12）。
    传成 ``beg+1`` 会把年初漏掉：年末 = 各变动行之和，少了一整个期初，
    而报表看上去只是「年初那一行没参与」，极容易误判成格式问题。
    """
    made = []
    for c in cols:
        L = get_column_letter(c)
        f = "=" + "+".join("%s%d" % (L, r) for r in range(r1, r2 + 1))
        ws.cell(row, c, f)
        made.append("%s%d" % (L, row))
    return made


def write_skeleton_row(ws, row: int, money_cols=(5, 8, 11, 14, 17, 20, 23)) -> None:
    """给新变动行补骨架公式：7 个审定列 + Z/AA/AB 三个合计列。

    新行**不是空白行**。模板的 ``[…]`` 行自带 10 个骨架公式
    （``G17=SUM(E17:F17)``、``Z17=SUM(E17,…,W17)``），少了它们 Z 合计不认这一行，
    年末行的 Z 列就会少算。
    """
    G = get_column_letter
    for m in money_cols:
        ws.cell(row, m + 2, "=SUM(%s%d:%s%d)" % (G(m), row, G(m + 1), row))
    ws.cell(row, 26, "=SUM(%s)" % ",".join("%s%d" % (G(m), row) for m in money_cols))
    ws.cell(row, 27, "=SUM(%s)" % ",".join("%s%d" % (G(m + 1), row) for m in money_cols))
    ws.cell(row, 28, "=SUM(%s)" % ",".join("%s%d" % (G(m + 2), row) for m in money_cols))


# ------------------------------------------------------------------ 编排

def apply_plan(wb, ws, plan: list[dict], label_col: int = 3) -> dict:
    """按计划自下而上插行。

    ``plan`` = ``[{"at": 15, "labels": ["改造转入","转固冲回"]}]``，
    ``at`` 是**模板当前坐标**。自下而上执行，保证先插的下方行不被上方插入挪位。
    每插一次立刻平移本表公式与全簿跨表引用。
    """
    trace = {"inserts": [], "xrefs": [], "order": []}
    for item in sorted(plan, key=lambda x: -x["at"]):
        at, labels = item["at"], list(item["labels"])
        n = len(labels)
        if not n:
            continue
        insert_rows_keep(ws, at, n)

        # ★ 顺序：先平移已有公式，再写新行。
        # 反过来的话，新行骨架里的 `=SUM(E15:F15)` 会被平移成 `E17`（因为 15>=at），
        # 新行的审定列就指到别的行去了 —— 而且不报错，Z 合计看着还有数。
        self_hits = shift_sheet_formulas(ws, at, n)
        other_hits = shift_workbook_refs(wb, ws.title, at, n)

        # 原 at 行被推到 at+n，用它的样式做母本
        for i, lab in enumerate(labels):
            r = at + i
            copy_row_style(ws, at + n, r)
            ws.cell(r, label_col, lab)
            write_skeleton_row(ws, r)

        trace["inserts"].append({"at": at, "n": n, "labels": labels})
        trace["order"].append(at)
        trace["xrefs"].append({"at": at, "n": n, "self": self_hits, "other": other_hits})
    return trace
