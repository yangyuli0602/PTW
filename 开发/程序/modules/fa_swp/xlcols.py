# -*- coding: utf-8 -*-
"""Excel 列结构手术：插列、按列公式平移、列块克隆、横向合计重写。

与 `xlrows.py` 对称，但**改的是列不是行**。两者必须分开：行手术里塞列逻辑会让
「平移」这个动作带 axis 参数，而 `shift_formula` 的签名已经被 58 项回归钉住了。

**为什么插列比插行更危险**（实测，不是推测）：

`ws.insert_cols()` 与 `insert_rows()` 一样只搬格子，**公式一个字符都不改**。
在 Z 列前插 3 列后，`AC12` 仍是 `=SUM(E12,H12,K12,N12,Q12,T12,W12)` —— 7 项，
新插的第 8 个类别块**不进合计**。而 `AB`（审定合计）同样少一项，K.01 的 CHECK
是 `AB − Lead`，于是 CHECK 恒 0、表3 恒 0、TB 差异恒 0，文件能打开、没有任何
错误值，只有金额是错的。

所以这里有三件事是必选项，不是优化：
1. `rewrite_across_sums` —— 横向合计必须重写成覆盖全部块的形态；
2. `rebuild_block_cf` —— 那条挂了 21 个区间的「空值涂黄」规则必须整条重写，
   平移只会把 7 段挪到新列号，新块永远进不去（新类别的空格就不会被标黄）；
3. `insert_cols_keep` 里的列宽按**区间**搬 —— 按列名搬会拆坏
   `W(min=23,max=24)` 和 `AF(min=32,max=16384)`，金额列被压到 11 列宽显示 `####`。

这套工具不认识 K.01 的业务语义（类别、段落、方式都不认），只吃
`ws / at / n / 列号列表`，别的底稿可以整体搬走。
"""

from __future__ import annotations

import re
from copy import copy

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter, column_index_from_string

#: 与 xlrows 同一套（那边是行版），这里只用来拆 A1 引用
_REF = re.compile(
    r"^(?:(?:'(?P<q>(?:[^']|'')+)'|(?P<u>[^'!]+))!)?"
    r"(?P<a1>\$?[A-Z]{1,3}\$?\d+)(?::(?P<a2>\$?[A-Z]{1,3}\$?\d+))?$"
)
_A1 = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")

#: 展开列宽的列数上限。超过这个就当成「尾段维度」处理，只重建一个维度，
#: 绝不逐列展开 —— 16384 列展开会把内存撑爆、文件撑到几十 MB。
_WIDTH_CAP = 200


# ------------------------------------------------------------------ 公式平移

def _shift_a1_col(ref: str, at: int, n: int) -> str:
    """只改列号，行号原样。"""
    m = _A1.match(ref)
    if not m:
        raise ValueError("认不出的单元格引用：%r" % ref)
    c_abs, col, r_abs, row = m.groups()
    c = column_index_from_string(col)
    if c >= at:
        c += n
    return "%s%s%s%s" % (c_abs, get_column_letter(c), r_abs, row)


def shift_formula_cols(formula: str, at: int, n: int,
                       target_sheets: set[str] | None = None,
                       same_sheet: bool = True) -> str:
    """把公式里「列号 >= at」的引用列 +n。行号一律不动。

    ``same_sheet`` / ``target_sheets`` 的语义与 `xlrows.shift_formula` 完全一致：
    改本表公式时用默认值；改别的页里指向本表的引用时传 `same_sheet=False`。
    这两个开关必须分开 —— 改 Lead 里指向 K.01 的 `Z42` 时，顺手把 Lead 自己的
    `=IF(C15=…)` 也 +3，整张封面就乱了。
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    out = []
    for tok in Tokenizer(formula).items:
        if tok.type == "OPERAND" and tok.subtype == "RANGE":
            m = _REF.match(tok.value)
            if not m:
                # 整列 A:A、整行 5:5、3D 引用、定义名。模板里还留着
                # `'FA list'!#REF!`（6 处，属于 74 个错误格基线），原样放行。
                if "#REF" in tok.value:
                    out.append(tok.value)
                    continue
                raise ValueError("公式里有不支持的引用：%r（%s）" % (tok.value, formula))
            sheet = None
            if m["q"] is not None:
                sheet = m["q"].replace("''", "'")
            elif m["u"] is not None:
                sheet = m["u"]
            hit = (same_sheet and sheet is None) or \
                  (target_sheets is not None and sheet in target_sheets)
            if hit:
                a1 = _shift_a1_col(m["a1"], at, n)
                a2 = _shift_a1_col(m["a2"], at, n) if m["a2"] else ""
                pref = tok.value[:tok.value.index(m["a1"])]
                # Token 签名是 (value, type, subtype)，别按直觉写反
                tok = Token(pref + a1 + (":" + a2 if a2 else ""), tok.type, tok.subtype)
            out.append(tok.value)
        else:
            out.append(tok.value)
    return "=" + "".join(out)


def shift_sheet_formulas_cols(ws, at: int, n: int, target: str | None = None) -> int:
    """本表所有公式过一遍按列平移。"""
    sheets = {target} if target else None
    hit = 0
    for row in ws.iter_rows():
        for c in row:
            if c.data_type != "f":
                continue
            new = shift_formula_cols(c.value, at, n, sheets)
            if new != c.value:
                c.value = new
                hit += 1
    return hit


def shift_workbook_col_refs(wb, target: str, at: int, n: int) -> list[tuple]:
    """其它 sheet 里指向 ``target`` 且列号 >= at 的引用全部 +n。

    K.01 的 Z 列被 8 处引用（Lead 三格、K.02.1 一格、K.02.2 三格、K.03.2 一格），
    Z 右移后漏改一格，Lead 的互校就非 0。
    """
    hits = []
    for ws in wb.worksheets:
        if ws.title == target:
            continue
        for row in ws.iter_rows():
            for c in row:
                if c.data_type != "f":
                    continue
                new = shift_formula_cols(c.value, at, n, {target}, same_sheet=False)
                if new != c.value:
                    hits.append((ws.title, c.coordinate, c.value, new))
                    c.value = new
    return hits


# ------------------------------------------------------------------ 条件格式

def _shift_sqref_cols(sqref, at: int, n: int) -> str:
    """一个 sqref → 按列三态平移后的字符串。"""
    out = []
    for cr in sqref.ranges:
        if cr.max_col < at:                 # 整块在左
            c1, c2 = cr.min_col, cr.max_col
        elif cr.min_col >= at:              # 整块在右
            c1, c2 = cr.min_col + n, cr.max_col + n
        else:                               # 插入点落在区间内部：撑开
            c1, c2 = cr.min_col, cr.max_col + n
        out.append("%s%d:%s%d" % (get_column_letter(c1), cr.min_row,
                                  get_column_letter(c2), cr.max_row))
    return " ".join(out)


def shift_cf_cols(ws, at: int, n: int) -> int:
    """条件格式区间按列平移（三态，见 `_shift_sqref_cols`）。

    多区间必须写回同一条规则 —— 拆成 N 条会让规则重复 N 份。
    """
    keep = [( _shift_sqref_cols(rng.sqref, at, n), list(rng.rules))
            for rng in list(ws.conditional_formatting)]
    ws.conditional_formatting = ConditionalFormattingList()
    for sq, rules in keep:
        for rule in rules:
            ws.conditional_formatting.add(sq, rule)
    return sum(len(rules) for _sq, rules in keep)


def shift_dv_cols(ws, at: int, n: int) -> int:
    """数据验证区间按列平移（K.01 是 0 条，模板换版可能有）。"""
    hit = 0
    for dv in ws.data_validations.dataValidation:
        dv.sqref = _shift_sqref_cols(dv.sqref, at, n)
        hit += 1
    return hit


# ------------------------------------------------------------------ 列宽

def _snapshot_widths(ws, cap: int = _WIDTH_CAP) -> tuple[dict, list]:
    """把 ColumnDimension 当**区间**记下来，而不是「字母 → 宽度」。

    `W(min=23,max=24)` 与 `AF(min=32,max=16384)` 是模板里的两个跨列维度。
    按字母搬只会搬第一个（23 搬走、24 留在原地变成默认宽）；而 `del AF`
    会把 32 列之后整片默认宽度一起删掉。

    返回 ``(per, tail)``：
    - ``per``：``{列号: (width, hidden, bestFit)}``，区间展开成逐列（≤ cap 列）
    - ``tail``：``[(lo, width, hidden)]``，hi 超过 cap 的尾段，只记起点
    """
    per, tail = {}, []
    for key, dim in list(ws.column_dimensions.items()):
        lo = dim.min or (column_index_from_string(key) if key.isalpha() else None)
        if not lo:
            continue
        hi = dim.max or lo
        w = dim.width
        if w is None:
            continue
        if hi > cap:
            tail.append((lo, w, bool(dim.hidden)))
            continue
        for c in range(lo, hi + 1):
            per[c] = (w, bool(dim.hidden), bool(dim.bestFit))
    return per, tail


def _restore_widths(ws, per: dict, tail: list, at: int, n: int) -> int:
    """先全删再按新位置写回。

    必须「先全删」：`insert_cols` 把宽度留在**原列号**上（实测 Z 仍是 11.38，
    而内容已经跑到 AC），逐 key 改会留下一堆重复维度。
    另外 `ws.column_dimensions[k]` 读一个不存在的 key 会**自动创建**
    width=13 的默认维度，所以读一律用 `.get()`。
    """
    for key in list(ws.column_dimensions.keys()):
        try:
            del ws.column_dimensions[key]
        except (KeyError, TypeError):
            pass

    written = 0
    for c, (w, hidden, best) in per.items():
        nc = c + n if c >= at else c
        d = ws.column_dimensions[get_column_letter(nc)]
        d.width = w
        # min/max 必须显式写。只设 width 时 openpyxl 建出来的维度 min/max 都是
        # None，序列化出去 <col> 元素缺 min/max 属性，Excel 直接忽略这条宽度 ——
        # 表现为「列宽设了又变回去」，而且文件里看着有维度，排查极费时间。
        d.min = nc
        d.max = nc
        if hidden:
            d.hidden = True
        if best:
            d.bestFit = True
        written += 1
    for lo, w, hidden in tail:
        nlo = lo + n if lo >= at else lo
        d = ws.column_dimensions[get_column_letter(nlo)]
        d.width = w
        d.min = nlo
        d.max = 16384
        if hidden:
            d.hidden = True
        written += 1
    return written


# ------------------------------------------------------------------ 插列主体

def insert_cols_keep(ws, at: int, n: int) -> dict:
    """在 ``at`` 列前插 ``n`` 列，把合并区与列宽摆正。

    六步，顺序是硬约束：

    1. 快照列宽（按区间）
    2. **旧坐标** unmerge 所有跨过/位于插入点右侧的合并区，记下左上格的值
    3. `insert_cols`
    4. 重建列宽
    5. 按**新坐标**重挂合并区（跨点的撑开、右侧的右移），写回备份值
    6. 不碰公式、不碰条件格式 —— 交给调用方

    第 2 步必须在第 3 步之前。反过来做（先插后解）时 openpyxl 的
    `_clean_merge_range` 会把值连同样式一起清成 Normal —— 这条在
    `bkd/expense.py` 的 V1.2.35 真实翻车过，注释还在。
    """
    per, tail = _snapshot_widths(ws)

    merges = []
    for mc in list(ws.merged_cells.ranges):
        if mc.max_col >= at:
            merges.append((mc.min_row, mc.min_col, mc.max_row, mc.max_col,
                           ws.cell(mc.min_row, mc.min_col).value))
            ws.unmerge_cells(str(mc))

    ws.insert_cols(at, n)

    widths = _restore_widths(ws, per, tail, at, n)

    for (r1, c1, r2, c2, val) in merges:
        if c1 < at <= c2:                   # 插入点落在合并区内部：撑开
            nc1, nc2 = c1, c2 + n
        else:                               # 整块在右侧：右移
            nc1, nc2 = c1 + n, c2 + n
        ws.merge_cells(start_row=r1, start_column=nc1, end_row=r2, end_column=nc2)
        if val is not None:
            ws.cell(r1, nc1).value = val

    return {"at": at, "n": n, "merges": len(merges),
            "widths": len(per), "tail": len(tail)}


# ------------------------------------------------------------------ 列块克隆

def clone_block_cols(ws, src: int, dst: int, rows, n: int = 3,
                     *, value_rows_max: int, widths: tuple | None = None) -> dict:
    """把 ``src`` 起的 ``n`` 列（一个类别块）克隆到 ``dst``。

    ``src`` / ``dst`` 都是**块首列**（账面列）。

    值复制规则是分行的，不是一律复制：
    - ``row <= value_rows_max``（表头三行：类别名行、子表头行、横幅行）：
      **复制值与样式** —— `账面数` / `账表调整·审计调整` / `审定数` / `【…】`
      必须带过去，否则新块没有表头。
    - ``row > value_rows_max``：**只复制样式与公式，不复制值** —— 否则会把
      第 7 块（或其他母本块）的数字抄进新块，看起来像「新类别已经填好了」。

    块内公式用 ``shift_formula_cols(f, at=src, n=dst-src)`` 平移即可，
    因为已验证块内公式**只引用本块三列**（``=SUM(W12:X12)``、``=W12+W13+…``、
    ``=W12-W19-W25``），整块右移后统一 +3 恰好正确。
    """
    step = dst - src
    made = {"cells": 0, "formulas": 0, "values": 0}

    for i in range(n):
        sc, dc = src + i, dst + i
        if widths and i < len(widths):
            ws.column_dimensions[get_column_letter(dc)].width = widths[i]
        for r in rows:
            s = ws.cell(r, sc)
            d = ws.cell(r, dc)
            if s.has_style:
                d._style = copy(s._style)
            if s.data_type == "f":
                d.value = shift_formula_cols(s.value, at=src, n=step)
                made["formulas"] += 1
            elif r <= value_rows_max and s.value not in (None, ""):
                d.value = s.value
                made["values"] += 1
            made["cells"] += 1
    return made


# ------------------------------------------------------------------ 横向合计

def _across_sum_cols(formula) -> dict | None:
    """识别「同行、逐项枚举的横向 SUM」，返回 ``{"row": r, "cols": [...]}``。

    要排除三种形近但语义不同的公式：
    - ``=SUM(E12:F12)``（区间）—— 汇总是「账面+调整」，是块内关系，不是跨块
    - ``=SUM(E12:E17)``（行区间）—— 纵向合计
    - ``=SUM(F75:F79)``（表4 合计）—— 区间

    判定靠三条：参数全是**单格**、所有引用的**行号相同**、**至少两个不同列**
    （同列两项的 `=SUM(E12,E13)` 不是横向枚举）。
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return None
    try:
        toks = Tokenizer(formula).items
    except Exception:                                   # noqa: BLE001
        return None
    if len(toks) < 4:
        return None
    if toks[0].type != "FUNC" or toks[0].value.upper() != "SUM(":
        return None
    if toks[-1].type != "FUNC" or toks[-1].value != ")":
        return None

    refs = []
    for t in toks[1:-1]:
        if t.type == "SEP":
            continue
        if t.type == "OPERAND" and t.subtype == "RANGE":
            m = _REF.match(t.value)
            if not m or m["a2"] or m["q"] is not None or m["u"] is not None:
                return None                             # 区间或跨表，不是本类
            refs.append(m["a1"])
        else:
            return None
    if len(refs) < 2:
        return None

    rows, cols = set(), []
    for r in refs:
        m = _A1.match(r)
        if not m:
            return None
        cols.append(column_index_from_string(m.group(2)))
        rows.add(int(m.group(4)))
    if len(rows) != 1 or len(set(cols)) < 2:
        return None
    return {"row": rows.pop(), "cols": cols}


def rewrite_across_sums(ws, block_role_cols: dict[str, list[int]]) -> int:
    """把横向枚举 SUM 重写成**覆盖全部块**的形态。

    ``block_role_cols`` = ``{"book": [...], "adj": [...], "aud": [...]}``，
    三个列表长度必须相同（每块三个角色各一列）。

    为什么不按行号枚举（``Z12/AA12/AB12`` 及 42/44/46/48 行）：
    插列后行号是稳定的，但**哪些行该有横向合计**是由模板决定的，写死行号等于
    把模板知识复制一份到代码里；而模板换版时新增一行横向合计，写死版本会静默漏掉。
    识别式一遍过全表，覆盖不到才说明模板变了，比漏改安全。
    """
    roles = {k: list(v) for k, v in block_role_cols.items() if v}
    hits = 0
    for row in ws.iter_rows():
        for c in row:
            if c.data_type != "f":
                continue
            info = _across_sum_cols(c.value)
            if not info:
                continue
            for cols in roles.values():
                if len(info["cols"]) == len(cols) and set(info["cols"]) == set(cols):
                    break                                # 已经是全覆盖，跳过
                if set(info["cols"]) <= set(cols):
                    new = "=SUM(" + ",".join(
                        "%s%d" % (get_column_letter(x), info["row"]) for x in cols) + ")"
                    if new != c.value:
                        c.value = new
                        hits += 1
                    break
    return hits


# ------------------------------------------------------------------ 空值告警重建

def _cf_rules(ws) -> list:
    """把 ``[(sqref 字符串, [rule, ...])]`` 摊平出来，便于按结构找规则。"""
    return [(str(rng.sqref), list(rng.rules)) for rng in ws.conditional_formatting]


def _find_blank_rule(ws, money_cols: list[int],
                     seg_rows: list[tuple[int, int]]) -> tuple[str, list] | None:
    """找「空值涂黄」那条：单列区间、行段集合与三个变动段完全一致。

    不能要求它「已经覆盖全部块列」—— 插列之后老规则只覆盖插列前的那些列，
    新列本来就不在里面，那正是要修的东西。按**行段特征**认领才不会找错。
    """
    want = set(money_cols)
    segset = set(seg_rows)
    for sq, rules in _cf_rules(ws):
        rng = _sqref_ranges(sq)
        if not rng:
            continue
        if any(c1 != c2 for c1, _r1, c2, _r2 in rng):
            continue                                # 跨列区间，不是本类
        if not {c1 for c1, _r1, _c2, _r2 in rng} <= want:
            continue                                # 覆盖了类别块以外的列
        if {(r1, r2) for _c1, r1, _c2, r2 in rng} != segset:
            continue                                # 行段不是那三段
        return sq, rules
    return None


def _find_name_rule(ws, cat_row: int) -> tuple[str, list] | None:
    """找「类别名格为空就涂黄」那条：全是单格且都在类别名行。"""
    for sq, rules in _cf_rules(ws):
        rng = _sqref_ranges(sq)
        if not rng:
            continue
        if all(c1 == c2 and r1 == r2 == cat_row for c1, r1, c2, r2 in rng):
            return sq, rules
    return None


def _sqref_ranges(sq: str) -> list[tuple]:
    """``"E12:E17 H12:H17"`` → ``[(5,12,5,17), (8,12,8,17)]``。"""
    from openpyxl.worksheet.cell_range import CellRange

    out = []
    for part in sq.split():
        try:
            cr = CellRange(part)
        except Exception:                               # noqa: BLE001
            return []
        out.append((cr.min_col, cr.min_row, cr.max_col, cr.max_row))
    return out


def _replace_rule(ws, old_sq: str, new_sq: str, rules: list) -> None:
    """把 ``old_sq`` 上的规则整体搬到 ``new_sq``，其余规则原样保留。"""
    keep = [(sq, rs) for sq, rs in _cf_rules(ws) if sq != old_sq]
    keep.append((new_sq, rules))
    ws.conditional_formatting = ConditionalFormattingList()
    for sq, rs in keep:
        for rule in rs:
            ws.conditional_formatting.add(sq, rule)


def rebuild_block_cf(ws, cat_row: int, name_cols: list[int],
                     seg_rows: list[tuple[int, int]], *, first_cell: str = "E12") -> dict:
    """整条重写两条「块相关」的条件格式。

    - 空值涂黄：``N 块 × 3 段`` 的区间，覆盖每块的**账面列**，跳过年末行
    - 类别名涂黄：N 个单格，都在类别名行

    **必须是重写，不能是平移。** 平移只会把原来的 7 段挪到新列号，新插的第 8 块
    永远进不去 —— 新类别的空单元格不会被标黄，而「黄底 = 该填未填」是这套底稿
    的人工待办入口，丢了就没人补。更麻烦的是规则**条数仍是 38 条**，看文件属性
    完全发现不了。
    """
    money = [c - 1 for c in name_cols]

    sq_blank = " ".join(
        "%s%d:%s%d" % (get_column_letter(c), a, get_column_letter(c), b)
        for c in money for a, b in seg_rows
    )
    sq_name = " ".join("%s%d" % (get_column_letter(c), cat_row) for c in name_cols)

    done = {"blank": False, "name": False}

    hit = _find_blank_rule(ws, money, seg_rows)
    if hit:
        old_sq, rules = hit
        # 公式里的首格跟着新 sqref 的第一段走，防止将来块 1 左移
        for rule in rules:
            if getattr(rule, "formula", None):
                rule.formula = ["LEN(TRIM(%s))=0" % first_cell]
        _replace_rule(ws, old_sq, sq_blank, rules)
        done["blank"] = True

    hit = _find_name_rule(ws, cat_row)
    if hit:
        old_sq, rules = hit
        first_name = "%s%d" % (get_column_letter(name_cols[0]), cat_row)
        for rule in rules:
            if getattr(rule, "formula", None):
                rule.formula = ["LEN(TRIM(%s))=0" % first_name]
        _replace_rule(ws, old_sq, sq_name, rules)
        done["name"] = True

    if not all(done.values()):
        raise ValueError("K.01 找不到要重建的条件格式（空值=%s / 类别名=%s）。"
                         "模板结构变了，不能静默跳过 —— 丢了黄底告警，"
                         "人工待办就没人补。" % (done["blank"], done["name"]))
    return {"blank_ranges": len(money) * len(seg_rows),
            "name_cells": len(money), "sqref": sq_blank[:120]}


def add_tb_cell_cf(ws, money_cols: list[int], tb_rows: list[int],
                   ref_money: int) -> int:
    """给新插的类别块补上 TB 核对行的「空值涂黄」单格规则。

    模板给 7 个块 × 4 个 TB 行各挂了一条单格规则（共 32 条）。新块不加，
    它的 TB 核对格空白时不会标黄，而检查的人会以为「这个格本来就不填」。
    条数刻意**不合并**成 4 条 —— 合并会改动规则总数，回归基准要跟着动，
    收益为零。
    """
    src = None
    for sq, rules in _cf_rules(ws):
        rng = _sqref_ranges(sq)
        if len(rng) == 1 and rng[0][0] == rng[0][2] == ref_money \
                and rng[0][1] in tb_rows and rng[0][1] == rng[0][3]:
            src = rules
            break
    if not src:
        return 0
    hit = 0
    for c in money_cols:
        for r in tb_rows:
            sq = "%s%d" % (get_column_letter(c), r)
            for rule in src:
                ws.conditional_formatting.add(sq, copy_rule(rule))
                hit += 1
    return hit


def copy_rule(rule):
    """复制一条条件格式规则（同一 rule 对象挂到多个 sqref 上时 openpyxl 会共享
    dxfId，虽然能写出来，但改一处会连带改全部，所以复制一份更安全）。"""
    from copy import deepcopy

    try:
        return deepcopy(rule)
    except Exception:                                   # noqa: BLE001
        return rule
