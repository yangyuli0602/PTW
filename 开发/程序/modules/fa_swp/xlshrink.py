# -*- coding: utf-8 -*-
"""Excel 列/行收缩手术：删列、删行、按删除点三态平移、横向合计重写。

与 `xlcols.py`（插列）、`xlrows.py`（插行）并列。**为什么不塞进那两个文件**：
它们的 `shift_formula_cols(at, n)` 是「双向单调平移」，签名被 100 多项回归钉死；
而删除需要第三种状态（引用落在被删区间里就没有了），参数含义也不同（`d` 是删掉的
数量，不是偏移量）。混在一起会让「平移」这个动作带 axis + 模式两个隐式开关。

## 为什么删比插危险

`ws.delete_cols()` / `ws.delete_rows()` 只搬 `_cells` 里的 Cell 对象，**公式一个
字符都不改**。实测在本模板上删掉第 17~25 列（4 个类别、模板留了 7 块）：

    Q12  =SUM(E12,H12,K12,N12,Q12,T12,W12)   ← 引用自己，循环引用
    Q18  =Z12+Z13+…+Z19                      ← 去加右侧空区，年末合计恒 0
    表3  H62=C62-Z18                          ← 指向空，表3 恒 0、CHECK 恒 0

文件能正常打开、`#REF!` 数量为 0、没有任何错误值。这是最难查的一类失效：
合计仍然「平」，因为 CHECK 比的 `AB` 也一起少了。

## 三态规则

对一个标量坐标 v，删除区间是 [at, at+d)：

    v < at          → 不变
    at <= v < at+d  → **不存在**（单格引用到这里就是 bug，抛错）
    v >= at+d       → v - d

对一个区间 [v1, v2] 则允许**收缩**：`SUM(C55:C61)` 在 at=59,d=3 下变成
`SUM(C55:C58)` —— 这是合计行「把要删的数据行括进去了」的正常修法，不该抛错。
区间整体落在删除区里才是 bug。

引用落在删除区一律**抛错**，不写 `#REF!`。写 `#REF!` 会让 Excel 打开时弹
「需要修复」，而且会和「这格本来就该是错误值」混在一起，事后分不清是没改还是
真的没有。条件格式是唯一例外：整段区间被删就直接丢弃，因为那条规则本来就在标记
要删的列。
"""

from __future__ import annotations

import re

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter, column_index_from_string

#: 与 xlcols 同一套正则，只用来拆 A1 引用
_REF = re.compile(
    r"^(?:(?:'(?P<q>(?:[^']|'')+)'|(?P<u>[^'!]+))!)?"
    r"(?P<a1>\$?[A-Z]{1,3}\$?\d+)(?::(?P<a2>\$?[A-Z]{1,3}\$?\d+))?$"
)
_A1 = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")

#: 列宽展开上限，超过当尾段（与 xlcols 一致）
_WIDTH_CAP = 200


# ------------------------------------------------------------------ 三态坐标

def _map_scalar(v: int, at: int, d: int) -> int | None:
    """单点三态。返回 None 表示「这个坐标已经不存在了」。"""
    if v < at:
        return v
    if v < at + d:
        return None
    return v - d


def _map_interval(lo: int, hi: int, at: int, d: int) -> list[tuple[int, int]]:
    """区间三态，返回 0 或 1 段。

        [lo, hi] 与删除区间 [at, at+d) 的关系：
        全在左边        → 原样
        全在右边        → 整体左移 d
        完全落在删除区  → 丢弃（返回空列表）
        跨界            → 裁掉被删部分
    """
    if lo > hi:
        lo, hi = hi, lo
    end = at + d
    if hi < at:
        return [(lo, hi)]
    if lo >= end:
        return [(lo - d, hi - d)]
    if lo >= at and hi < end:
        return []
    a = lo if lo < at else at
    b = hi - d if hi >= end else at - 1
    return [(a, b)] if b >= a else []


# ------------------------------------------------------------------ 公式平移

def _split(ref: str) -> tuple[str, int, str, int]:
    m = _A1.match(ref)
    if not m:
        raise ValueError("认不出的单元格引用：%r" % ref)
    c_abs, col, r_abs, row = m.groups()
    return c_abs, column_index_from_string(col), r_abs, int(row)


def _join(c_abs: str, c: int, r_abs: str, r: int) -> str:
    return "%s%s%s%d" % (c_abs, get_column_letter(c), r_abs, r)


def _shift_a1_del(ref: str, axis: str, at: int, d: int) -> str | None:
    """只改一个轴；返回 None 表示这个引用落在了删除区。"""
    c_abs, c, r_abs, r = _split(ref)
    if axis == "col":
        nc = _map_scalar(c, at, d)
        if nc is None:
            return None
        return _join(c_abs, nc, r_abs, r)
    nr = _map_scalar(r, at, d)
    if nr is None:
        return None
    return _join(c_abs, c, r_abs, nr)


def _shift_range_del(a1: str, a2: str, axis: str,
                     at: int, d: int) -> tuple[str, str] | None:
    """区间三态；返回 None 表示整段都在删除区。"""
    c1a, c1, r1a, r1 = _split(a1)
    c2a, c2, r2a, r2 = _split(a2)
    lo, hi = (c1, c2) if axis == "col" else (r1, r2)
    segs = _map_interval(lo, hi, at, d)
    if not segs:
        return None
    a, b = segs[0]
    if axis == "col":
        return _join(c1a, a, r1a, r1), _join(c2a, b, r2a, r2)
    return _join(c1a, c1, r1a, a), _join(c2a, c2, r2a, b)


def shift_formula_del(formula: str, axis: str, at: int, d: int,
                      target_sheets: set[str] | None = None,
                      same_sheet: bool = True) -> str:
    """按删除点重写公式里的引用（三态）。

    ``axis`` = ``"col"`` / ``"row"``，只改那一个轴。
    ``same_sheet`` / ``target_sheets`` 语义与 `xlcols.shift_formula_cols` 一致：
    改本表用默认值；改别的页里指向本表的引用时传 ``same_sheet=False``，
    否则会把别页**自己**的格子也一起挪。
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    out = []
    for tok in Tokenizer(formula).items:
        if tok.type == "OPERAND" and tok.subtype == "RANGE":
            m = _REF.match(tok.value)
            if not m:
                # 整列 A:A / 3D / 定义名 / `'FA list'!#REF!`（模板自带 6 处）
                if "#REF" in tok.value:
                    out.append(tok.value)
                    continue
                raise ValueError("公式里有不支持的引用：%r" % tok.value)
            sheet = None
            if m["q"] is not None:
                sheet = m["q"].replace("''", "'")
            elif m["u"] is not None:
                sheet = m["u"]
            hit = (same_sheet and sheet is None) or \
                  (target_sheets is not None and sheet in target_sheets)
            if hit:
                pref = tok.value[:tok.value.index(m["a1"])]
                if m["a2"]:
                    seg = _shift_range_del(m["a1"], m["a2"], axis, at, d)
                    if seg is None:
                        raise ValueError("区间 %s:%s 整段落在删除区" % (m["a1"], m["a2"]))
                    body = seg[0] + ":" + seg[1]
                else:
                    one = _shift_a1_del(m["a1"], axis, at, d)
                    if one is None:
                        raise ValueError("引用 %s 落在删除区" % m["a1"])
                    body = one
                tok = Token(pref + body, tok.type, tok.subtype)
            out.append(tok.value)
        else:
            out.append(tok.value)
    return "=" + "".join(out)


def shift_sheet_formulas_del(ws, axis: str, at: int, d: int,
                             target: str | None = None) -> list[tuple]:
    """本表所有公式过一遍。引用落在删除区 → 立刻抛，并指出是哪一格。"""
    sheets = {target} if target else None
    hits = []
    for row in ws.iter_rows():
        for c in row:
            if c.data_type != "f":
                continue
            try:
                new = shift_formula_del(c.value, axis, at, d, sheets)
            except ValueError as exc:                   # noqa: BLE001
                raise ValueError(
                    "%s!%s 的公式 %r 引用了被删除的%s（%s）：%s"
                    % (ws.title, c.coordinate, c.value,
                       "列" if axis == "col" else "行",
                       "第 %d~%d %s" % (at, at + d - 1, "列" if axis == "col" else "行"),
                       exc)) from None
            if new != c.value:
                hits.append((c.coordinate, c.value, new))
                c.value = new
    return hits


def shift_workbook_refs_del(wb, target: str, axis: str, at: int, d: int) -> list[tuple]:
    """其它 sheet 里指向 ``target`` 的引用按删除点重写。

    指向被删区间的**抛错**（说明别处还在引用一个马上要消失的块），
    不做静默平移 —— 平移到左边的列上是错的，且不会报错。
    """
    hits = []
    for ws in wb.worksheets:
        if ws.title == target:
            continue
        for row in ws.iter_rows():
            for c in row:
                if c.data_type != "f":
                    continue
                try:
                    new = shift_formula_del(c.value, axis, at, d,
                                            {target}, same_sheet=False)
                except ValueError as exc:               # noqa: BLE001
                    raise ValueError(
                        "%s!%s 引用了 %s 里即将被删除的%s：%r（%s）"
                        % (ws.title, c.coordinate, target,
                           "列" if axis == "col" else "行", c.value, exc)) from None
                if new != c.value:
                    hits.append((ws.title, c.coordinate, c.value, new))
                    c.value = new
    return hits


# ------------------------------------------------------------------ 条件格式

def shift_cf_del(ws, axis: str, at: int, d: int) -> dict:
    """条件格式 sqref 三态重写；整段落在删除区就丢弃该段。

    必须**整体重建列表**，不能逐条 del：`ConditionalFormattingList` 按 sqref
    字符串索引，改了一半再删会留下旧 key，文件里出现两份指向不同列的同一规则。
    """
    keep, dropped = [], 0
    for rng in list(ws.conditional_formatting):
        parts = []
        for cr in rng.sqref.ranges:
            if axis == "col":
                segs = _map_interval(cr.min_col, cr.max_col, at, d)
                if not segs:
                    dropped += 1
                    continue
                c1, c2 = segs[0]
                r1, r2 = cr.min_row, cr.max_row
            else:
                segs = _map_interval(cr.min_row, cr.max_row, at, d)
                if not segs:
                    dropped += 1
                    continue
                r1, r2 = segs[0]
                c1, c2 = cr.min_col, cr.max_col
            parts.append("%s%d:%s%d" % (get_column_letter(c1), r1,
                                        get_column_letter(c2), r2))
        if parts:
            keep.append((" ".join(parts), list(rng.rules)))

    ws.conditional_formatting = ConditionalFormattingList()
    for sq, rules in keep:
        for rule in rules:
            ws.conditional_formatting.add(sq, rule)
    return {"ranges": len(keep), "dropped": dropped,
            "rules": sum(len(rs) for _sq, rs in keep)}


def shift_dv_del(ws, axis: str, at: int, d: int) -> int:
    """数据验证 sqref 三态（K.01 是 0 条，模板换版可能有）。"""
    hit = 0
    for dv in ws.data_validations.dataValidation:
        parts = []
        for cr in dv.sqref.ranges:
            if axis == "col":
                segs = _map_interval(cr.min_col, cr.max_col, at, d)
                if not segs:
                    continue
                c1, c2 = segs[0]
                r1, r2 = cr.min_row, cr.max_row
            else:
                segs = _map_interval(cr.min_row, cr.max_row, at, d)
                if not segs:
                    continue
                r1, r2 = segs[0]
                c1, c2 = cr.min_col, cr.max_col
            parts.append("%s%d:%s%d" % (get_column_letter(c1), r1,
                                        get_column_letter(c2), r2))
        dv.sqref = " ".join(parts) if parts else "A1"
        hit += 1
    return hit


# ------------------------------------------------------------------ 尺寸（宽/高）

def _snapshot_dims(ws, axis: str) -> tuple[list, list]:
    """列宽/行高按**区间**快照。

    列宽里 `W(min=23,max=24)` 与 `AF(min=32,max=16384)` 是跨列维度；按「字母 →
    宽度」搬会拆坏它们（实测 H..P 九列全拿到被删列的宽度，金额显示 `####`）。
    行高没有跨行维度，每行一条。
    """
    units, tail = [], []
    if axis == "col":
        for key, dim in list(ws.column_dimensions.items()):
            lo = dim.min or (column_index_from_string(key) if key.isalpha() else None)
            if not lo or dim.width is None:
                continue
            hi = dim.max or lo
            if hi > _WIDTH_CAP:                 # 尾段只记起点，绝不展开 16384 列
                tail.append((lo, dim.width, bool(dim.hidden)))
                continue
            units.append((lo, hi, dim.width, bool(dim.hidden), bool(dim.bestFit)))
    else:
        for key, dim in list(ws.row_dimensions.items()):
            idx = dim.index or (int(key) if str(key).isdigit() else None)
            if not idx or dim.height is None:
                continue
            units.append((idx, idx, dim.height, bool(dim.hidden), False))
    return units, tail


def _write_dim(ws, axis: str, a: int, b: int, w, hidden: bool, best: bool) -> None:
    if axis == "col":
        d = ws.column_dimensions[get_column_letter(a)]
        d.width = w
        # min/max 必须显式写：只设 width 时 openpyxl 建出来的维度 min/max 是 None，
        # 序列化出去 <col> 缺 min/max，Excel 直接忽略这条宽度。
        d.min = a
        d.max = b
        if hidden:
            d.hidden = True
        if best:
            d.bestFit = True
    else:
        d = ws.row_dimensions[a]
        d.height = w
        if hidden:
            d.hidden = True


def _restore_dims_del(ws, units: list, tail: list, axis: str,
                      at: int, d: int) -> int:
    """先全删再按新坐标写回。`delete_cols` 会把宽度留在**原列号**上。"""
    store = ws.column_dimensions if axis == "col" else ws.row_dimensions
    for key in list(store.keys()):
        try:
            del store[key]
        except (KeyError, TypeError):
            pass

    n = 0
    for lo, hi, w, hidden, best in units:
        for a, b in _map_interval(lo, hi, at, d):
            _write_dim(ws, axis, a, b, w, hidden, best)
            n += 1
    for lo, w, hidden in tail:
        if lo < at:
            nlo = lo
        elif lo >= at + d:
            nlo = lo - d
        else:                                   # 尾段起点被删：塌到删除点，绝不丢
            nlo = at
        dd = ws.column_dimensions[get_column_letter(nlo)]
        dd.width = w
        dd.min = nlo
        dd.max = 16384
        if hidden:
            dd.hidden = True
        n += 1
    return n


# ------------------------------------------------------------------ 删除主体

def _collect_merges(ws, axis: str, at: int) -> list[tuple]:
    """记下要处理的合并区（含左上格的值）并在**旧坐标**拆掉。

    必须先拆后删：`delete_cols` 不动合并区，删完它要么悬在表外、要么把左移进来的
    别的数据吞进去（实测 `B66:F69` 上移后盖住「表4」标题）。
    """
    got = []
    for mc in list(ws.merged_cells.ranges):
        edge = mc.max_col if axis == "col" else mc.max_row
        if edge < at:
            continue
        got.append((mc.min_row, mc.min_col, mc.max_row, mc.max_col,
                    ws.cell(mc.min_row, mc.min_col).value))
        ws.unmerge_cells(str(mc))
    return got


def _rehang_merges(ws, merges: list[tuple], axis: str, at: int, d: int) -> int:
    kept = 0
    for (r1, c1, r2, c2, val) in merges:
        if axis == "col":
            segs = _map_interval(c1, c2, at, d)
            if not segs:
                continue
            nc1, nc2 = segs[0]
            nr1, nr2 = r1, r2
        else:
            segs = _map_interval(r1, r2, at, d)
            if not segs:
                continue
            nr1, nr2 = segs[0]
            nc1, nc2 = c1, c2
        ws.merge_cells(start_row=nr1, start_column=nc1,
                       end_row=nr2, end_column=nc2)
        if val is not None:
            ws.cell(nr1, nc1).value = val
        kept += 1
    return kept


def delete_cols_keep(ws, at: int, d: int) -> dict:
    """删掉 ``at`` 起的 ``d`` 列，把列宽与合并区摆正。

    四步，顺序是硬约束（与 `insert_cols_keep` 对称）：

    1. 快照列宽（按区间）
    2. **旧坐标** unmerge 所有伸进删除区或在其右侧的合并区，记下左上值
    3. `delete_cols`
    4. 三态重建列宽 → 按新坐标重挂合并区并写回值

    不碰公式、不碰条件格式 —— 交给调用方，因为调用方知道该先重写哪些格
    （比如横向合计必须先按新块列表重生成，否则通用平移会因为「引用落在删除区」而抛）。
    """
    units, tail = _snapshot_dims(ws, "col")
    merges = _collect_merges(ws, "col", at)
    ws.delete_cols(at, d)
    widths = _restore_dims_del(ws, units, tail, "col", at, d)
    kept = _rehang_merges(ws, merges, "col", at, d)
    return {"at": at, "d": d, "widths": widths,
            "merges": kept, "merges_dropped": len(merges) - kept}


def delete_rows_keep(ws, at: int, d: int) -> dict:
    """删掉 ``at`` 起的 ``d`` 行，把行高与合并区摆正（对称实现）。"""
    units, tail = _snapshot_dims(ws, "row")
    merges = _collect_merges(ws, "row", at)
    ws.delete_rows(at, d)
    heights = _restore_dims_del(ws, units, tail, "row", at, d)
    kept = _rehang_merges(ws, merges, "row", at, d)
    return {"at": at, "d": d, "heights": heights,
            "merges": kept, "merges_dropped": len(merges) - kept}


# ------------------------------------------------------------------ 横向合计

def _across_sum_cols(formula) -> dict | None:
    """识别「同行、逐项枚举的横向 SUM」，返回 ``{"row": r, "cols": [...]}``。

    排除 `=SUM(E12:F12)`（块内区间）、`=SUM(E12:E17)`（纵向）、`=SUM(F75:F79)`
    （表4 合计）—— 三条判定：参数全是单格、行号相同、至少两个不同列。
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
                return None
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


def rewrite_total_sums(ws, total_cols: dict, block_role_cols: dict) -> int:
    """把合计三列里的横向枚举 SUM **无条件**按新块列表重写。

    现有的 `xlcols.rewrite_across_sums` 判据是「旧列 ⊆ 新列」，只扩不缩；
    删列时旧公式**多出**被删的三项，判据不成立 → 命中 0、静默不改，接着通用平移
    就会因为引用落在删除区而抛。所以这里不看旧列，只看「这是不是一格横向合计」：

    - 只扫合计三列（``total_cols`` 的 book/adj/aud）
    - 公式必须是 `SUM(单格, 单格, …)` 且所有引用**同一行**且就是本行
    - 重写成 `SUM(块1对应角色, 块2…, …)`

    `G12=SUM(E12:F12)`（区间）、`AA18=AA12+…`（加号）、`C62=SUM(C55:C61)`
    （同一列纵向）三条都不会误命中。
    """
    roles = {k: list(v) for k, v in block_role_cols.items() if v}
    hits = 0
    for role, tcol in total_cols.items():
        cols = roles.get(role)
        if not cols:
            continue
        for r in range(1, ws.max_row + 1):
            c = ws.cell(r, tcol)
            if c.data_type != "f":
                continue
            info = _across_sum_cols(c.value)
            if not info or info["row"] != r:
                continue
            new = "=SUM(" + ",".join(
                "%s%d" % (get_column_letter(x), r) for x in cols) + ")"
            if new != c.value:
                c.value = new
                hits += 1
    return hits
