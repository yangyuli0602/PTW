# -*- coding: utf-8 -*-
"""固定资产 SWP · 把「上年 K.01」「上年 Lead」「原始 TB」嵌进本年度稿。

三条硬约束（用户 2026-11-10 定）：

  1. 上年页**保留全部格式/样貌/颜色**，单元格内容**粘贴为数值**（一个公式都不留）；
  2. TB 页只放 1601/1602/1603 的 44 行**原始**记录（含整块重复 20 行，标出来）；
  3. 隐藏级别 ``hidden``（与货币资金一致），不用 veryHidden。

**为什么用同簿 ``copy_worksheet`` 而不是 core.py 那套跨簿方案**：
core 的 ``copy_prior_lead_sheet`` 是给「从另一个文件搬上年 Lead」写的，它硬上限
400 行 × **40 列**，而 K.01 是 A1:DF110 —— 41 列以后还有 722 个带格式的格子，
走它会把右半边悄悄切掉。本模块的源和目标是同一个工作簿，``_style`` 索引同簿有效，
不存在 core 注释里说的「跨簿串字体」问题（实测 E12 的 ``_style`` 两边完全一致）。

**``copy_worksheet`` 会丢四样东西**，必须补回来，而且补完要自检条数：

    K.01：合并区 5 ✅ ｜ 条件格式 38 ❌ 全丢 ｜ 数据验证 0 ｜ 110 列 ✅
    Lead：合并区 20 ✅ ｜ 条件格式 9 ❌ 全丢 ｜ 数据验证 3 ❌ 全丢
    两张：冻结窗格 / zoomScale / showGridLines ❌ 全为 None

少了那 38 条条件格式，K.01 的空白告警（浅黄底）就没了，别人打开会以为格式没做。
"""

from __future__ import annotations

import re
from copy import copy

from openpyxl.formatting.rule import Rule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter

from .fmt import ACCT_FMT, DATE_FMT, INT_FMT

#: sheet 名只在这里定义一次。公式里引用一律走 q()，不手写字符串 ——
#: 少一对单引号就是 #NAME? 或者「文件已损坏」，而且看公式本身看不出来。
SHEET_TB = "TB 固定资产"
SHEET_PRIOR_K01 = "上年 K.01 Agree SL to GL"
SHEET_PRIOR_LEAD = "上年 K.00 Lead Sheet"

MONEY_FMT = ACCT_FMT
WARN_FONT = Font(name="微软雅黑", size=9, color="C00000", bold=True)


def q(name: str) -> str:
    """sheet 名 → 公式里的引用前缀。中文名和空格都必须带单引号。"""
    return "'%s'" % str(name).replace("'", "''")


# ------------------------------------------------------------------ sheet 复制

def clone_sheet(wb, src_name: str, new_name: str, after: str | None = None,
                expect_empty_cell: str | None = None) -> dict:
    """在同一工作簿里复制一张 sheet，补齐 copy_worksheet 丢掉的东西。

    ``expect_empty_cell``：复制前断言源页的这个格还是空的（如 ``"Z42"``），
    用来保证「复制发生在填报之前」—— 顺序反了会把本年数复制成上年数，
    而且一分钱错误都不报，是这一整套改造里最难发现的错。
    """
    if src_name not in wb.sheetnames:
        raise KeyError("工作簿里没有 %r，现有 %s" % (src_name, wb.sheetnames))

    src = wb[src_name]
    if expect_empty_cell and src[expect_empty_cell].value not in (None, ""):
        raise ValueError(
            "拒绝复制：%s 的 %s 已经有值（%r），说明填报已经发生。"
            "复制必须放在任何 fill 之前，否则上年页会被填成本年数。"
            % (src_name, expect_empty_cell, src[expect_empty_cell].value))

    if new_name in wb.sheetnames:
        del wb[new_name]

    dst = wb.copy_worksheet(src)
    dst.title = new_name

    # 条件格式：同簿可以直接复用 rule.dxf（wb._differential_styles.add 会按相等去重）
    for rng in src.conditional_formatting:
        for rule in rng.rules:
            dst.conditional_formatting.add(str(rng.sqref), Rule(
                type=rule.type, dxf=rule.dxf, stopIfTrue=rule.stopIfTrue,
                formula=rule.formula, operator=rule.operator,
                text=rule.text, timePeriod=rule.timePeriod))

    # 数据验证
    for dv in src.data_validations.dataValidation:
        ndv = copy(dv)
        ndv.sqref = copy(dv.sqref)
        dst.add_data_validation(ndv)

    # 视图（copy_worksheet 不搬这三样，缺了看起来「格子线全冒出来、缩放变 100%」）
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines
    dst.sheet_view.zoomScale = src.sheet_view.zoomScale
    dst.freeze_panes = src.freeze_panes
    if src.sheet_properties.tabColor is not None:
        dst.sheet_properties.tabColor = copy(src.sheet_properties.tabColor)

    # 图片（运行时无 Pillow，模板读入时本来就是空列表；将来装了会自动跟上）
    warnings = []
    try:
        for im in getattr(src, "_images", []):
            dst.add_image(copy(im))
    except Exception as e:                              # noqa: BLE001
        warnings.append("图片复制跳过：%s" % e)

    # 位置
    if after and after in wb.sheetnames:
        wb.move_sheet(dst, offset=wb.sheetnames.index(after) - wb.sheetnames.index(new_name) + 1)

    # 自检：合并区 / 条件格式 / 数据验证 条数必须和源一致，不一致直接抛，不警告后继续
    got_m, want_m = len(dst.merged_cells.ranges), len(src.merged_cells.ranges)
    got_c = sum(len(r.rules) for r in dst.conditional_formatting)
    want_c = sum(len(r.rules) for r in src.conditional_formatting)
    got_d, want_d = len(dst.data_validations.dataValidation), len(src.data_validations.dataValidation)
    bad = []
    if got_m != want_m:
        bad.append("合并区 %d≠%d" % (got_m, want_m))
    if got_c != want_c:
        bad.append("条件格式 %d≠%d" % (got_c, want_c))
    if got_d != want_d:
        bad.append("数据验证 %d≠%d" % (got_d, want_d))
    if bad:
        raise RuntimeError("%s → %s 复制后格式不全：%s" % (src_name, new_name, "；".join(bad)))

    return {"sheet": dst, "name": new_name, "merged": got_m, "cf": got_c,
            "dv": got_d, "images": len(getattr(dst, "_images", [])), "warnings": warnings}


# ------------------------------------------------------------------ 公式 → 数值

def clear_formulas(ws) -> dict[tuple[int, int], str]:
    """把所有公式**先快照、再清空**，返回 ``{(row, col): "=..."}``。

    顺序反了就再也拿不回原文（清空后无从回解），K.01 的表2 会变成一排空行。
    """
    saved = {}
    for (r, c), cell in list(ws._cells.items()):
        if cell.data_type == "f":
            saved[(r, c)] = cell.value
            cell.value = None
    return saved


_SIMPLE_SAME = re.compile(r"^=\$?([A-Z]{1,3})\$?(\d+)$")


def resolve_simple_refs(ws, saved: dict, rounds: int = 3) -> int:
    """把 ``=F10`` 这类**同表**单格引用解成值（表2 的 B55:B61 就是 ``=F10``）。

    只解同表单格引用。跨表引用（``='K.01 Agree SL to GL'!Z42``）一律不解 ——
    那指向的是**本年**的 K.01，搬进上年页就成了「用本年的数冒充上年」，
    页面上有数、看着像，一分钱都不报错。
    """
    n = 0
    for _ in range(max(1, rounds)):
        changed = 0
        for (r, c), f in list(saved.items()):
            if not isinstance(f, str):
                continue
            m = _SIMPLE_SAME.match(f.strip())
            if not m:
                continue
            v = ws.cell(row=int(m.group(2)), column=column_index_from_string(m.group(1))).value
            if v is not None and not (isinstance(v, str) and v.startswith("=")):
                ws.cell(row=r, column=c, value=v)
                del saved[(r, c)]
                changed += 1
        n += changed
        if not changed:
            break
    return n


def count_formulas(ws) -> int:
    return sum(1 for cell in ws._cells.values() if cell.data_type == "f")


def assert_no_formula(ws, label: str = "") -> int:
    """保存前最后一关：残留公式必须为 0（用户口径：粘贴为数值）。"""
    n = count_formulas(ws)
    if n:
        raise ValueError("%s 里还残留 %d 个公式，没有做到「粘贴为数值」" % (label or ws.title, n))
    return n


# ------------------------------------------------------------------ TB 页

TB_COLS = (
    ("company", "主体", 30),
    ("acct_code", "科目编码", 10),
    ("acct_name", "科目名称", 16),
    ("aux1", "辅助核算1", 12),
    ("aux2", "辅助核算2", 14),
    ("aux_text", "辅助核算全文", 20),
    ("open_bal", "期初余额", 16),
    ("end_bal", "期末余额", 16),
    ("cur_debit", "本期借方", 15),
    ("cur_credit", "本期贷方", 15),
    ("src_row", "TB行号", 8),
    ("cat", "类别归集", 12),
    ("counted", "是否计入汇总", 12),
    ("dup_of", "重复于行号", 10),
)


def write_tb_sheet(wb, rows_raw: list[dict], dropped: list[dict],
                   cats_end: dict | None = None) -> dict:
    """画 TB 页：44 行原始记录 + 类别归集 + 是否重复标记。

    ``rows_raw`` 必须是 ``prep.load_tb(conn)`` 重查出来的原始行 ——
    ``prep.load()["rows"]`` 已经被 ``reconcile_opening`` 改过 ``open_bal``，
    不是 TB 原样，放进来就与「原始数据」四个字不符。

    ``cats_end`` 传本年矩阵，用于页末写「与 K.01 差异」核对。
    """
    from . import prep     # 避免循环 import

    if SHEET_TB in wb.sheetnames:
        del wb[SHEET_TB]
    ws = wb.create_sheet(SHEET_TB)

    dup_keys = {(d["company"], d["acct"], d["src_row"]): d["dup_of"] for d in dropped}

    ws["A1"] = "TB 科目余额表原始记录（1601 / 1602 / 1603，含整块重复 %d 行）" % len(dropped)
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("本页为原始 TB 直出，**未去重**。M 列「否」＝整块重复导入行（同一段数据被导了两遍，"
                "不在汇总里）；K.01 的 TB 核对块只取 M 列「是」的行。"
                "1602/1603 在 TB 是贷方负数，K.01 按正数口径取负号。")
    ws["A2"].font = WARN_FONT
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)
    ws.row_dimensions[2].height = 34

    hdr = 4
    thin_fill = PatternFill("solid", fgColor="4472C4")
    for i, (_k, head, width) in enumerate(TB_COLS, 1):
        c = ws.cell(hdr, i, head)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = thin_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width

    first = hdr + 1
    for r, row in enumerate(rows_raw, first):
        key = (row["company_name"], row["acct_code"], row["src_row"])
        is_dup = key in dup_keys
        vals = {
            "company": row["company_name"], "acct_code": row["acct_code"],
            "acct_name": row["acct_name"], "aux1": row.get("aux1", ""),
            "aux2": row.get("aux2", ""), "aux_text": row.get("aux_text", ""),
            "open_bal": float(row["open_bal"]), "end_bal": float(row["end_bal"]),
            "cur_debit": float(row["cur_debit"]), "cur_credit": float(row["cur_credit"]),
            "src_row": row["src_row"], "cat": prep.category(row),
            "counted": "否" if is_dup else "是",
            "dup_of": dup_keys.get(key, ""),
        }
        for i, (k, _h, _w) in enumerate(TB_COLS, 1):
            c = ws.cell(r, i, vals[k])
            c.font = Font(name="微软雅黑", size=9)
            if k in ("open_bal", "end_bal", "cur_debit", "cur_credit"):
                c.number_format = MONEY_FMT
            elif k == "src_row":
                c.number_format = INT_FMT
            if is_dup:
                c.fill = PatternFill("solid", fgColor="F2F2F2")
    last = first + len(rows_raw) - 1

    # 页末三行核对 —— 绝不进 SUMIFS 范围（否则 TB→K.01→TB 成环）
    foot = last + 2
    ws.cell(foot, 1, "原始合计（含重复）").font = Font(bold=True, size=9)
    ws.cell(foot + 1, 1, "去重后合计（K.01 采用）").font = Font(bold=True, size=9)
    for i, k in enumerate(("open_bal", "end_bal"), 7):
        col = get_column_letter(i)
        ws.cell(foot, i, "=SUM($%s$%d:$%s$%d)" % (col, first, col, last)).number_format = MONEY_FMT
        ws.cell(foot + 1, i,
                '=SUMIFS($%s$%d:$%s$%d,$M$%d:$M$%d,"是")' % (col, first, col, last, first, last)
                ).number_format = MONEY_FMT

    if cats_end:
        r = foot + 2
        ws.cell(r, 1, "与 K.01 表1 年末核对：K.01 的 TB 块用 SUMIFS 直接取本页，差异为 0 才正常"
                ).font = Font(size=8, color="595959")

    ws.freeze_panes = ws.cell(first, 1)
    ws.sheet_view.showGridLines = False

    return {"sheet": ws, "hdr_row": hdr, "first": first, "last": last, "foot": foot,
            "range": "$%d:$%d" % (first, last),
            "cols": {k: get_column_letter(i) for i, (k, _h, _w) in enumerate(TB_COLS, 1)},
            "rows": len(rows_raw), "dup": len(dropped)}


# ------------------------------------------------------------------ 上年 K.01

def fill_prior_k01(ws, lay: dict, cats_used: list[str], prior_cats: dict, note: str) -> dict:
    """上年 K.01 只写年末余额行；年初与变动行留空（模板的 containsBlanks 会自动涂黄）。"""
    r = lay["rows"]
    zero = 0.0
    cells = {}

    def put(row, col, val):
        c = ws.cell(row, col, val)
        c.number_format = MONEY_FMT
        cells["%s%d" % (get_column_letter(col), row)] = val
        return c

    for i, cat in enumerate(cats_used):
        name_col = lay["cat_cols"][i]
        money_col = name_col - 1
        head = ws.cell(lay["cat_row"], name_col, cat)
        head.font = Font(bold=True, size=10)
        head.alignment = Alignment(horizontal="center", vertical="center")

        p = prior_cats[cat]
        put(r["fa"]["end"], money_col, float(p["fa"]))
        put(r["dep"]["end"], money_col, float(p["dep"]))
        put(r["imp"]["end"], money_col, float(p["imp"]))
        if "nbv" in r:
            if "beg" in r["nbv"]:
                put(r["nbv"]["beg"], money_col, float(p["fa"] - p["dep"] - p["imp"]))
            if "end" in r["nbv"]:
                put(r["nbv"]["end"], money_col, float(p["fa"] - p["dep"] - p["imp"]))

    # 合计列在**最后一个类别块之后的第 2 列**（7 块时是 Z，8 块时是 AC）。
    # 写死 26 在 8 类时会算错位：合计被写进第 8 块的账面列，
    # 表1 的「工具器具」那一列显示的就是集团合计。
    tot = {k: sum(prior_cats[c][k] for c in cats_used) for k in ("fa", "dep", "imp")}
    sum_col = (list(lay.get("cat_cols") or [26])[-1] + 2) if lay.get("cat_cols") else 26
    for key in ("fa", "dep", "imp"):
        row = r.get(key, {}).get("end")
        if row:
            ws.cell(row, sum_col, float(tot[key])).number_format = MONEY_FMT
            cells["%s%d" % (get_column_letter(sum_col), row)] = tot[key]
    nbv = tot["fa"] - tot["dep"] - tot["imp"]
    for row in r.get("nbv", {}).values():
        ws.cell(row, sum_col, float(nbv)).number_format = MONEY_FMT
        cells["%s%d" % (get_column_letter(sum_col), row)] = nbv

    # 说明放 B5（模板 B5 是空的说明区）
    cell = ws.cell(5, 2, note)
    cell.font = WARN_FONT
    cell.fill = PatternFill("solid", fgColor="FFF2CC")
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[5].height = 30
    return {"cells": cells, "cats": list(cats_used)}


# ------------------------------------------------------------------ 上年 Lead

def fill_prior_lead(ws, rows: dict, prior_cats: dict, info: dict, prior_date: str) -> dict:
    """上年 Lead：封面 + 表1 三行（E/G/I 同值，F/H 为 0），J 列留空。

    ``I`` 列必须等于 ``E+F+H`` —— 本年 Lead 的 J47 引的就是上年页的 I 列，
    差一分钱，本年 K.01 的 AC12 就立刻非 0。
    """
    tot_fa = float(sum(prior_cats[c]["fa"] for c in prior_cats))
    tot_dep = float(sum(prior_cats[c]["dep"] for c in prior_cats))
    tot_imp = float(sum(prior_cats[c]["imp"] for c in prior_cats))

    ws["C2"] = info.get("customer", "")
    ws["C3"] = prior_date
    ws["C3"].number_format = DATE_FMT
    if info.get("standard"):
        ws["C6"] = info["standard"]
    if info.get("currency"):
        ws["C7"] = info["currency"]

    out = {}
    for key, val in (("fa", tot_fa), ("dep", tot_dep), ("imp", tot_imp)):
        r = rows[key]
        for col in (5, 7, 9):                 # E / G / I
            c = ws.cell(r, col, val)
            c.number_format = MONEY_FMT
        for col in (6, 8):                    # F / H
            ws.cell(r, col, 0).number_format = MONEY_FMT
        out[key] = {"row": r, "value": val, "I": "I%d" % r}

    return out


def hide(wb, names: list[str]) -> list[str]:
    done = []
    for n in names:
        if n in wb.sheetnames:
            wb[n].sheet_state = "hidden"
            done.append(n)
    return done
