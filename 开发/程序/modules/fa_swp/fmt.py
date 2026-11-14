# -*- coding: utf-8 -*-
"""全簿数字格式归一。

底稿里同一页会出现两三种零、两种负号，根子在两处：

1. 代码自己写的格子用 ``#,##0.00;[Red]-#,##0.00``（零是 0.00、负数红色）
2. 模板自带的格子用 Excel「会计专用」或带 ``¥``/``$``/``€`` 的货币格式
   —— 代码没覆盖到，就原样留在成品里

所以光改常量不够，保存前还要把**整页**扫一遍，把「金额系」的格子统一成
``ACCT_FMT``（会计专用·无币种、零显示 -、负数括号）。

只改 ``number_format``：字体、填充、边框、列宽、合并区、条件格式一律不碰
——上年那两页客户要求保留原样貌，动一个就会走样。
"""

from __future__ import annotations

import os
import sys

try:
    from core.xlsx_fmt import (ACCT_FMT, DATE_FMT, INT_FMT, MONEY_FMT,  # noqa: F401
                               PCT_FMT, TEXT_FMT, is_money_format)
except ImportError:                                   # 独立运行时兜底
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from core.xlsx_fmt import (ACCT_FMT, DATE_FMT, INT_FMT, MONEY_FMT,  # noqa: F401
                               PCT_FMT, TEXT_FMT, is_money_format)

#: 要归一的页。页名逐字照抄 —— 「汇总 」和「K.02.1 新增测试 」都带尾空格。
FMT_SHEETS = (
    "汇总 ",
    "K.00 Lead Sheet", "上年 K.00 Lead Sheet",
    "K.01 Agree SL to GL", "上年 K.01 Agree SL to GL",
    "FA list", "新增清单", "处置清单", "TB 固定资产",
    "K.02.1 新增测试 ", "K.02.1a 新增选样输出",
    "K.02.2 处置测试", "K.02.2a 处置选样输出",
    "K.03.1 SAP", "K.03.2 折旧测试TOD", "K.03.3 折旧政策复核",
)


def normalize_money(wb, sheets=FMT_SHEETS, fmt: str = ACCT_FMT) -> dict:
    """把指定页的金额格统一成 ``fmt``，返回留痕。

    遍历用 ``ws._cells.values()``：只走**真实存在**的格，不新建。
    用 ``iter_rows()`` 会把整块矩形都实例化（FA list 近 2000 行 × 35 列），
    凭空多出几万个空样式格，文件白白变大。

    合并区非左上角的 MergedCell 不可写，跳过。
    """
    from openpyxl.cell.cell import MergedCell

    tr = {"sheets": 0, "cells": 0, "skipped": 0, "formats": {}, "missing": []}
    for name in sheets:
        ws = wb[name] if name in wb.sheetnames else None
        if ws is None:
            tr["missing"].append(name)
            continue
        tr["sheets"] += 1
        for cell in list(ws._cells.values()):
            if isinstance(cell, MergedCell):
                continue
            f = cell.number_format
            if f == fmt:
                continue
            if not is_money_format(f):
                tr["skipped"] += 1
                continue
            cell.number_format = fmt
            tr["cells"] += 1
            tr["formats"][f] = tr["formats"].get(f, 0) + 1
    return tr
