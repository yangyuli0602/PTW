# -*- coding: utf-8 -*-
"""上年底稿（K.01）的生成与回读。

**为什么要有这个模块**：用户要求「期初取上年 K.01」。瑞云没有历史固定资产底稿，
所以这里生成一份**模拟的**上年 K.01（金额直接用本年 TB 的期初数），
让本期程序真的去读一个外部文件取期初，而不是把期初硬编码在代码里。
这样「链接上期」这条链路是真跑通的，换个客户只要换文件即可。

生成的表故意做得很朴素（类别 × 原值/折旧/减值，一行合计），
因为它的唯一用途是被本年 K.01 回读；版式不需要跟正式底稿一致。
页首用红字标明它是模拟文件，防止流出去被当成真底稿。
"""

from __future__ import annotations

import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import prep
from .fmt import ACCT_FMT

TITLE = "K.01 Agree SL to GL（模拟上年底稿）"
SHEET = "K.01 Agree SL to GL"

#: 表头文字 → 读取时的字段名。改表头就能被 read() 认到，不写死列号。
HEADS = (("cat", ("固定资产类别", "类别")),
         ("fa", ("原值",)),
         ("dep", ("累计折旧",)),
         ("imp", ("减值准备",)),
         ("nbv", ("净值",)))


def _style(ws, ncol, nrows):
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"].font = Font(name="微软雅黑", size=9, color="C00000", bold=True)
    for c in range(1, ncol + 1):
        cell = ws.cell(5, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for r in range(6, 6 + nrows + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            cell.border = border
            if c > 1:
                cell.number_format = ACCT_FMT
        ws.cell(r, 1).alignment = Alignment(horizontal="left")


def build(cats: dict, customer: str, balance_date: str, out_path: str,
          order: list[str] | None = None) -> str:
    """写一份上年 K.01。

    ``cats`` 是 ``prep`` 的类别矩阵。**上年末余额 = 本年期初**，所以取 ``*_beg``
    字段，不要取 ``*_end`` —— 取错了会把本年的期末当成上年的期末，
    表面上看 K.01 的年初等于年末（变动全 0），净值却差了一整个年度。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET

    ws["A1"] = TITLE
    ws["A2"] = ("本文件为按客户科目余额表重建的模拟上年底稿，仅用于验证「期初取上年 K.01」"
                "这条取数链路；金额取自本年 TB 期初余额，版式从简，不可作为审计证据。")
    ws["A3"] = "客户名称：%s" % customer
    ws["A4"] = "资产负债表日：%s" % balance_date
    ws["A5"] = "表1 固定资产后推表 · 年末余额（单位：元）"

    heads = ["固定资产类别", "原值", "累计折旧", "减值准备", "净值"]
    hdr = 6
    for i, h in enumerate(heads, 1):
        ws.cell(hdr, i, h)

    order = list(order or cats)
    for i, cat in enumerate(order, 1):
        r = hdr + i
        m = cats[cat]
        ws.cell(r, 1, cat)
        ws.cell(r, 2, float(m["fa_beg"]))
        ws.cell(r, 3, float(m["dep_beg"]))
        ws.cell(r, 4, float(m["imp_beg"]))
        ws.cell(r, 5, "=B%d-C%d-D%d" % (r, r, r))

    tr = hdr + len(order) + 1
    ws.cell(tr, 1, "合计")
    for c, col in ((2, "B"), (3, "C"), (4, "D")):
        ws.cell(tr, c, "=SUM(%s%d:%s%d)" % (col, hdr + 1, col, tr - 1))
    ws.cell(tr, 5, "=B%d-C%d-D%d" % (tr, tr, tr))

    ws.cell(tr + 2, 1, "注：原值、累计折旧、减值准备均为正数口径；净值 = 原值 − 累计折旧 − 减值准备。")

    _style(ws, 5, len(order))
    ws.column_dimensions["A"].width = 26
    for c in "BCDE":
        ws.column_dimensions[c].width = 18
    ws.freeze_panes = "A7"

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)
    return out_path


def read(path: str, cats: list[str] | None = None) -> dict:
    """回读上年 K.01，返回 {类别: {"fa":..,"dep":..,"imp":..}}。

    按表头文字定位列、按类别名定位行；``cats`` 是本年期初需要的类别清单，
    缺任何一个就报错，不返回半张空表 —— 期初错一分钱，整张 K.01 都是错的。
    ``cats`` 不给时按表里实际出现的类别返回（只读场景用）。
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb.worksheets[0]

    hdr, cols = None, {}
    for r in range(1, min(ws.max_row, 20) + 1):
        texts = {}
        for c in range(1, min(ws.max_column, 10) + 1):
            t = str(ws.cell(r, c).value or "").replace(" ", "").replace("\n", "")
            if t:
                texts[c] = t
        hit = {}
        for key, kws in HEADS:
            for c, t in texts.items():
                if any(k in t for k in kws) and c not in hit.values():
                    hit[key] = c
                    break
        if {"cat", "fa", "dep", "imp"} <= set(hit):
            hdr, cols = r, hit
            break
    if hdr is None:
        raise ValueError("上年 K.01 里找不到「固定资产类别/原值/累计折旧/减值准备」表头：%s" % path)

    out = {}
    for r in range(hdr + 1, ws.max_row + 1):
        name = str(ws.cell(r, cols["cat"]).value or "").strip()
        if not name or name == "合计":
            continue
        if cats is None or name in cats:
            out[name] = {
                "fa": float(ws.cell(r, cols["fa"]).value or 0),
                "dep": float(ws.cell(r, cols["dep"]).value or 0),
                "imp": float(ws.cell(r, cols["imp"]).value or 0),
            }
    miss = [c for c in (cats or ()) if c not in out]
    if miss:
        raise ValueError("上年 K.01 缺少类别 %s，不能拿它当期初" % miss)
    return out
