# -*- coding: utf-8 -*-
"""固定资产 SWP · 三张清单页的写入。

模板里的 ``FA list`` / ``新增清单`` / ``处置清单`` 三张表**只有说明文字、没有表头行**
（说明写在 B2~B5：「项目组可以自定义固定资产清单表格」），所以表头由这里建。

表头行固定从第 7 行开始（上面留给模板原有说明），数据从第 8 行起。
每个写入函数返回 ``{字段: (起始列字母, 结束列字母)}``，供 K.01 拼 SUMIFS 用 ——
**列字母是算出来的，不是写死的**，以后增删列不会把公式指错。
"""

from __future__ import annotations

import datetime as dt

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .fmt import ACCT_FMT, DATE_FMT, INT_FMT, MONEY_FMT, PCT_FMT, TEXT_FMT
HDR_FILL = PatternFill("solid", fgColor="4472C4")
HDR_FONT = Font(bold=True, color="FFFFFF", size=9)
WARN_FONT = Font(name="微软雅黑", size=9, color="C00000", bold=True)
BODY_FONT = Font(name="微软雅黑", size=9)

HDR_ROW = 7
FIRST_ROW = 8

#: FA list 的列：(字段, 表头, 宽度, 类型)。类型 money/date/int/text/rate
FA_COLS = (
    ("no", "序号", 6, "int"),
    ("code", "资产编码", 14, "text"),
    ("name", "资产名称", 46, "text"),
    ("cat", "资产类别", 12, "text"),
    ("dept", "使用部门", 14, "text"),
    ("loc", "存放地点", 18, "text"),
    ("start", "启用日期", 12, "date"),
    ("life", "折旧年限(月)", 12, "int"),
    ("resid", "残值率", 9, "rate"),
    ("fa_beg", "期初原值", 16, "money"),
    ("dep_beg", "期初累计折旧", 16, "money"),
    ("imp_beg", "期初减值准备", 15, "money"),
    ("nbv_beg", "期初净值", 16, "money"),
    ("add", "本期增加", 16, "money"),
    ("reduce", "本期减少", 16, "money"),
    ("prov_dep", "本期计提折旧", 15, "money"),
    ("reduce_dep", "本期减少折旧", 15, "money"),
    ("fa_end", "期末原值", 16, "money"),
    ("dep_end", "期末累计折旧", 16, "money"),
    ("imp_end", "期末减值准备", 15, "money"),
    ("nbv_end", "期末净值", 16, "money"),
)

MV_COLS = (
    ("no", "序号", 6, "int"),
    ("cat", "资产类别", 12, "text"),
    ("date", "凭证日期", 12, "date"),
    ("vno", "凭证号", 18, "text"),
    ("name", "资产名称", 50, "text"),
    ("dept", "使用部门", 14, "text"),
    ("way", "新增方式", 16, "text"),
    ("orig", "原值", 16, "money"),
    ("note", "备注", 26, "text"),
)

DP_COLS = (
    ("no", "序号", 6, "int"),
    ("cat", "资产类别", 12, "text"),
    ("date", "处置日期", 12, "date"),
    ("vno", "凭证号", 18, "text"),
    ("name", "资产名称", 50, "text"),
    ("way", "处置方式", 14, "text"),
    ("orig", "处置原值", 16, "money"),
    ("dep", "累计折旧", 16, "money"),
    ("imp", "减值准备", 15, "money"),
    ("nbv", "账面净值", 16, "money"),
    ("gain", "处置损益", 14, "money"),
    ("note", "备注", 22, "text"),
)


def _thin():
    s = Side(style="thin", color="D9D9D9")
    return Border(left=s, right=s, top=s, bottom=s)


def _write_head(ws, cols, title, warn) -> dict:
    """写标题 + 红字声明 + 表头，返回字段→列区间。"""
    ws.cell(1, 1, title).font = Font(bold=True, size=13)
    ws.cell(2, 1, warn).font = WARN_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=min(len(cols), 12))
    ws.row_dimensions[2].height = 30
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="center")

    bd = _thin()
    for i, (key, head, width, _kind) in enumerate(cols, 1):
        cell = ws.cell(HDR_ROW, i, head)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bd
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[HDR_ROW].height = 32

    return {key: (get_column_letter(i), get_column_letter(i)) for i, (key, *_rest) in enumerate(cols, 1)}


def _put_row(ws, r, cols, row, warn_blank=True) -> None:
    bd = _thin()
    for i, (key, _head, _w, kind) in enumerate(cols, 1):
        v = row.get(key)
        if isinstance(v, float):
            v = round(v, 2)
        cell = ws.cell(r, i, v)
        cell.font = BODY_FONT
        cell.border = bd
        if kind == "money":
            cell.number_format = MONEY_FMT
        elif kind == "date":
            cell.number_format = DATE_FMT
            cell.alignment = Alignment(horizontal="center")
        elif kind == "int":
            cell.number_format = INT_FMT
            cell.alignment = Alignment(horizontal="center")
        elif kind == "rate":
            cell.number_format = PCT_FMT
            cell.alignment = Alignment(horizontal="center")
        elif key == "code":
            cell.number_format = TEXT_FMT     # 防前导零被吃掉
    if warn_blank:
        for key in ("dept", "loc"):
            i = [c[0] for c in cols].index(key) + 1 if key in [c[0] for c in cols] else None
            if i and ws.cell(r, i).value in (None, ""):
                ws.cell(r, i).fill = PatternFill("solid", fgColor="FFF2CC")


WARN = ("本清单为按客户科目辅助核算与序时账重建的模拟明细：金额、日期、凭证号、摘要全部取自客户真实账套，"
        "资产编码、使用部门、存放地点、折旧年限为程序补全。仅用于验证底稿版式与勾稽，不可作为审计证据。")


def write_fa(ws, rows) -> dict:
    cols = _write_head(ws, FA_COLS, "FA list 固定资产明细清单（期末）", WARN)
    ws.cell(4, 1, "说明：期初余额合计 = K.01 表1 年初；期末余额合计 = K.01 表1 年末；"
                  "「本期计提折旧」「本期减少折旧」两列供 K.01 计提行取数。").font = Font(size=9, color="595959")
    for i, r in enumerate(rows, FIRST_ROW):
        _put_row(ws, i, FA_COLS, r)
    n = len(rows)
    tr = FIRST_ROW + n
    ws.cell(tr, 1, "合计").font = Font(bold=True, size=9)
    for key in ("fa_beg", "dep_beg", "imp_beg", "nbv_beg", "add", "reduce", "prov_dep",
                "reduce_dep", "fa_end", "dep_end", "imp_end", "nbv_end"):
        i = [c[0] for c in FA_COLS].index(key) + 1
        col = get_column_letter(i)
        c = ws.cell(tr, i, "=SUM(%s%d:%s%d)" % (col, FIRST_ROW, col, tr - 1))
        c.number_format = MONEY_FMT
        c.font = Font(bold=True, size=9)
    ws.freeze_panes = ws.cell(FIRST_ROW, 1)
    ws.auto_filter.ref = "A%d:%s%d" % (HDR_ROW, get_column_letter(len(FA_COLS)), tr - 1)
    return cols


def write_moves(ws, rows) -> dict:
    cols = _write_head(ws, MV_COLS, "固定资产变动清单（本期增加）",
                       "本清单的每一行对应一张真实凭证，金额与 K.01「购置」「在建工程转入」行合计一致。")
    ws.cell(4, 1, "说明：模板原页名为「新增清单」（K.01 的交叉索引指向 K.02.1），"
                  "按用户口径同时承载购置、在建工程转入两类变动，用「新增方式」列区分。"
                  ).font = Font(size=9, color="595959")
    for i, r in enumerate(rows, FIRST_ROW):
        _put_row(ws, i, MV_COLS, r)
    n = len(rows)
    tr = FIRST_ROW + n
    ws.cell(tr, 1, "合计").font = Font(bold=True, size=9)
    i = [c[0] for c in MV_COLS].index("orig") + 1
    col = get_column_letter(i)
    c = ws.cell(tr, i, "=SUM(%s%d:%s%d)" % (col, FIRST_ROW, col, tr - 1))
    c.number_format = MONEY_FMT
    c.font = Font(bold=True, size=9)
    ws.freeze_panes = ws.cell(FIRST_ROW, 1)
    ws.auto_filter.ref = "A%d:%s%d" % (HDR_ROW, get_column_letter(len(MV_COLS)), tr - 1)
    return cols


def write_disposals(ws, rows) -> dict:
    cols = _write_head(ws, DP_COLS, "固定资产处置清单（本期减少）",
                       "本清单的每一行对应一张真实凭证；「处置方式」列区分处置或报废、内部转出，"
                       "分别对应 K.01 的两个减少行。")
    for i, r in enumerate(rows, FIRST_ROW):
        r = dict(r)
        r["nbv"] = r["orig"] - r["dep"] - r["imp"]
        r["gain"] = 0.0          # 瑞云仅 3 张凭证带 6115 资产处置损益，本表按 0 列示
        r.setdefault("note", "报废/内部转移，无处置损益" if r.get("gain") == 0 else "")
        _put_row(ws, i, DP_COLS, r)
    n = len(rows)
    tr = FIRST_ROW + n
    ws.cell(tr, 1, "合计").font = Font(bold=True, size=9)
    for key in ("orig", "dep", "imp", "nbv"):
        i = [c[0] for c in DP_COLS].index(key) + 1
        col = get_column_letter(i)
        c = ws.cell(tr, i, "=SUM(%s%d:%s%d)" % (col, FIRST_ROW, col, tr - 1))
        c.number_format = MONEY_FMT
        c.font = Font(bold=True, size=9)
    ws.freeze_panes = ws.cell(FIRST_ROW, 1)
    ws.auto_filter.ref = "A%d:%s%d" % (HDR_ROW, get_column_letter(len(DP_COLS)), tr - 1)
    return cols
