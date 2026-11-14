# -*- coding: utf-8 -*-
"""数字格式常量 + 「这串格式算不算金额」的判据。

**全平台只在这里写格式串**，各模块一律 import，不再各自复制一份。

为什么要有 ``is_money_format``：
客户模板里的金额格式有十几种写法（``#,##0.00``、带 ``[Red]`` 的、带 ``¥``/``$``/``€``
的、Excel 内置「会计专用」的），靠精确匹配字符串认不全，模板一换版就漏。
按结构判断（第一段里有没有 ``0``/``#`` 占位、小数位几位、有没有 ``%`` 或日期字母）
对「金额」这一类稳定成立。

为什么不按 numFmt id 白名单：175/178/183/184/185/200/201/43/165 是同一批样式在
不同模板里的编号，换一份客户模板就变。

用户 2026-11-13 口径：底稿里所有数字格统一成「会计专用·无币种」——
两位小数、千分位、**零显示为 -**、负数用括号。
"""

from __future__ import annotations

import re

#: 会计专用·无币种。零显示「-」，负数用括号，末尾带一位对齐空格。
#: 与 ``modules/bkd/core.py`` 的 ACCT_FMT 逐字符一致。
ACCT_FMT = '_ * #,##0.00_ ;_ * (#,##0.00)_ ;_ * "-"??_ ;_ @_ '

#: 兼容旧名字：fa_swp 里原来有 20 余处引用 MONEY_FMT，改名要动一大片。
MONEY_FMT = ACCT_FMT

PCT_FMT = "0.00%"
DATE_FMT = "yyyy-mm-dd"
INT_FMT = "0"
TEXT_FMT = "@"

_QUOTE = re.compile(r'"[^"]*"')
_BRACKET = re.compile(r"\[[^\]]*\]")
#: `_x`（留白一个字符宽）与 `*x`（重复填充）—— 不是数字占位符
_PAD = re.compile(r"[_*].")
_CJK = re.compile(r"[\u4e00-\u9fff]")


def _first_section(fmt: str) -> str:
    """取格式串的第一段（正数段）。引号里的分号不算分隔。"""
    out, inq = [], False
    for ch in fmt:
        if ch == '"':
            inq = not inq
        if ch == ";" and not inq:
            break
        out.append(ch)
    return "".join(out)


def is_money_format(fmt) -> bool:
    """这串 number_format 是不是「金额系」，应当统一成 ACCT_FMT。

    判定只看第一段（正数段），按计划里的顺序短路：

        空 / General          → 否
        引号里有中文          → 否（``0.00"万元"`` 这种带单位，硬转会丢单位）
        剥掉 [Red] [$-804] _x *x 引号字面量 \\ 之后：
          含 %               → 否
          含 y/d/h/s         → 否（日期时间。**必须先剥 [Red]** —— Red 里有 d）
          没有 0 或 # 占位    → 否（纯文本 @）
          小数位 > 2          → 否（汇率 0.0000，转两位会丢精度）
        其余                 → 是
    """
    if not fmt:
        return False
    f = str(fmt).strip()
    if not f or f.lower() == "general":
        return False

    for lit in _QUOTE.findall(f):
        if _CJK.search(lit):
            return False

    sec = _first_section(f)
    sec = _QUOTE.sub("", sec)
    sec = _BRACKET.sub("", sec)      # [Red] / [$€-2] / [$-804] 一起剥
    sec = _PAD.sub("", sec)          # _ * 这类填充符
    sec = sec.replace("\\", "")
    if not sec.strip():
        return False

    if "%" in sec:
        return False
    if re.search(r"[ydhs]", sec, re.I):
        return False
    if not re.search(r"[0#]", sec):
        return False
    # 纯 `0` / `#` 是序号、折旧年限(月) 这类计数列，不是金额。
    # 金额至少带千分位或小数位；模板里的金额格没有一个写成裸 `0`。
    if sec.strip() in ("0", "#"):
        return False

    m = re.search(r"\.(0+)", sec)
    if m and len(m.group(1)) > 2:
        return False
    return True
