# -*- coding: utf-8 -*-
"""固定资产 SWP · K.00 Lead Sheet 的填充与互链。

Lead 是整份底稿的封面页，四行数字（原值/累计折旧/减值准备/净值）与 K.01 双向锁死：

    Lead!E47 = K.01!Z42      K.01!AC12 = AB12 − Lead!J47
    Lead!E48 = K.01!Z44      K.01!AC18 = AB18 − Lead!I47
    Lead!E49 = K.01!Z46      K.01!AC19 = AB19 − Lead!J48

改任何一边，另一边的 CHECK 立刻变红。这就是用户说的「表1 生成调整和抵消的
lead 表内链接公式」—— 模板自带的这六个格子已经写好了，本模块**只填不改公式**，
补的是 F 列（账表调整，仍为 0）、J 列（上期末审定数）和 H 列。

H 列（审计调整）不再是常数 0：照货币资金的做法，引用下方「调整汇总表」按科目名称
汇总（见 ``fill_adj_summary``）。汇总表空着时 SUMIFS 算出来就是 0，审定数与原行为一致。
"""

from __future__ import annotations

from copy import copy as _copy_style

from openpyxl.styles import Alignment, Font

from . import embed, prep
from .fmt import ACCT_FMT, DATE_FMT, MONEY_FMT

SHEET = "K.00 Lead Sheet"

#: Lead 封面的标签 → 基础信息字段
LABELS = (("customer", ("客户名称", "被审计单位")),
          ("balance_date", ("资产负债表日",)),
          ("te", ("可容忍误差",)),
          ("sad", ("名义金额",)),
          ("standard", ("适用会计准则", "会计准则")),
          ("currency", ("记账本位币",)))

#: 表1 列（表头文字 → 字段）
COLS = (("code", ("总账科目编码",)), ("name", ("科目名称",)), ("idx", ("索引号",)),
        ("book", ("期末账面数",)), ("adj_bs", ("账表调整",)), ("raw", ("期末未审数",)),
        ("adj_aud", ("审计调整",)), ("aud", ("期末审定数",)), ("prior", ("上期末审定数",)),
        ("chg", ("变动金额",)), ("chgp", ("变动%", "变动率")))


def _txt(v) -> str:
    return str(v or "").replace(" ", "").replace("\n", "").replace("\u3000", "")


def fill_header(ws, info) -> dict:
    """写封面 B2:B7 那一块。值写在标签右侧第一个空格，按标签文字找，不写死格。"""
    done = {}
    for key, kws in LABELS:
        val = info.get(key)
        if val in (None, ""):
            continue
        hit = None
        for r in range(1, 13):
            for c in range(1, 5):
                t = _txt(ws.cell(r, c).value)
                if not t or not any(k in t for k in kws):
                    continue
                for cc in range(c + 1, min(ws.max_column, c + 4) + 1):
                    if ws.cell(r, cc).value in (None, ""):
                        hit = (r, cc)
                        break
                if hit:
                    break
            if hit:
                break
        if not hit:
            continue
        cell = ws.cell(hit[0], hit[1], val)
        if key in ("te", "sad"):
            cell.number_format = ACCT_FMT
        elif key == "balance_date":
            cell.number_format = DATE_FMT
        done[key] = "%s%d" % (chr(64 + hit[1]) if hit[1] <= 26 else "?", hit[0])
    return done


def locate_table1(ws) -> dict:
    """定位 Lead 表1：表头行、列映射、原值/折旧/减值三行的行号。

    抽成独立函数是因为「上年 Lead」页也要用同一套定位 —— 复制出来的页结构一致，
    但行号要现场算，不能写死 47/48/49（换模板就错位）。
    """
    hdr = None
    for r in range(30, min(ws.max_row, 60) + 1):
        if "总账科目编码" in _txt(ws.cell(r, 2).value):
            hdr = r
            break
    if hdr is None:
        raise ValueError("Lead 里找不到表1 表头（B 列「总账科目编码」）")

    cols = {}
    for c in range(1, min(ws.max_column, 20) + 1):
        t = _txt(ws.cell(hdr, c).value)
        for key, kws in COLS:
            if key not in cols and any(k in t for k in kws):
                cols[key] = c
    need = ("book", "adj_bs", "adj_aud", "prior")
    if any(k not in cols for k in need):
        raise ValueError("Lead 表1 列不全，缺 %s" % [k for k in need if k not in cols])

    rows = {}
    for r in range(hdr + 1, hdr + 12):
        name = _txt(ws.cell(r, cols.get("name", 3)).value)
        if not name:
            continue
        if "原值" in name:
            rows["fa"] = r
        elif "累计折旧" in name:
            rows["dep"] = r
        elif "减值" in name:
            rows["imp"] = r
    if len(rows) < 3:
        raise ValueError("Lead 表1 找不到原值/累计折旧/减值准备三行：%s" % rows)
    return {"header_row": hdr, "cols": cols, "rows": rows}


def fill_table1(ws, prior_cats, prior_file, prior_sheet=None) -> dict:
    """补 F 列调整数（0）、J 列上期末审定数。

    H 列（审计调整）**不在这里写**：它要引用下方「调整汇总表」，
    由 ``fill_adj_summary`` 统一写。这样即使汇总表整块被删掉，也不会留下一个
    看着像 0、实际该有数的常数格。

    ``prior_sheet`` 给了就写公式引用「上年 Lead」的 I 列（期末审定数），
    不给则退回写死值（旧行为，保证旧调用点与回归测试不变）。
    """
    lay = locate_table1(ws)
    cols, rows = lay["cols"], lay["rows"]

    # Lead 没有类别维度，直接对字典求和。原来遍历固定元组，换成数据驱动后
    # 元组里可能没有某个类别，合计就会静默少算。
    tot = {"fa": sum(v["fa"] for v in prior_cats.values()),
           "dep": sum(v["dep"] for v in prior_cats.values()),
           "imp": sum(v["imp"] for v in prior_cats.values())}

    info = {}
    for key, val in tot.items():
        r = rows[key]
        # E 列（期末账面数）模板已是 =K.01!Z42/Z44/Z46，不覆盖
        ws.cell(r, cols["adj_bs"], 0).number_format = MONEY_FMT
        cell = ws.cell(r, cols["prior"], 0)
        if prior_sheet:
            # 引上年 Lead 的 I 列（期末审定数）。不能引 J 列 —— 那是「前年」，差一整年。
            cell.value = "=%s!$I$%d" % (embed.q(prior_sheet), r)
        else:
            cell.value = float(val)
        cell.number_format = MONEY_FMT
        info[key] = {"row": r, "prior": float(val)}

    # 说明：O 列（定性判断）由人写，程序不代判
    for r in rows.values():
        ws.cell(r, cols.get("chg", 11)).alignment = Alignment(horizontal="right")
    return {"header_row": lay["header_row"], "cols": cols, "rows": info}


#: 调整汇总表的列（B 起连续五列）。模板只有 B 列「调整类型」，C..F 由程序补。
ADJ_HEADERS = ("调整类型", "主体", "调整说明", "科目名称", "调整金额")


def fill_adj_summary(ws, tab) -> dict:
    """把 Lead 底部的「调整汇总表」补成五列，并让表1 的 H 列引用它。

    模板原本只有 B 列一个「调整类型」表头，金额无处可填，H 列（审计调整）
    只能靠人手工抄。现在补成 调整类型/主体/调整说明/科目名称/调整金额，
    H 列改成 SUMIFS 按科目名称取数 —— 填一笔调整，审定数自动跟着动。

    与货币资金（``modules/bkd/core.py::_fill_lead_adj``）的两处差异必须照抄对：

    * 固定资产 Lead 表1 的**科目名称在 C 列**（货币资金在 D 列），
      所以 SUMIFS 的条件值写 ``$C{row}``，不能抄 ``$D{row}``；
    * 只接管 H 列（审计调整）。F 列（账表调整）由 ``fill_table1`` 保持常数 0。

    ``tab`` 用 ``fill_table1`` 的返回值，不重算 layout。
    """
    lay = locate_table1(ws)
    title = None
    for r in range(40, min(ws.max_row, 120) + 1):
        if "调整汇总表" in _txt(ws.cell(r, 2).value):
            title = r
            break
    if title is None:
        # 宁可不出，也不出一份没有调整入口的底稿：H 列会退回常数 0，
        # 复核人填了调整也进不来，差错是静默的。
        raise ValueError("Lead 里找不到「调整汇总表」标题行（B 列），不能建立审计调整链接")
    hdr = title + 1

    # 补 C..F 表头：文字补空位，样式整格抄 B{hdr}（与货币资金一致）
    for i, txt in enumerate(ADJ_HEADERS[1:], start=3):
        cell = ws.cell(hdr, i)
        if not _txt(cell.value):
            cell.value = txt
        try:
            cell._style = _copy_style(ws.cell(hdr, 2)._style)
        except Exception:
            pass
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 数据区：优先读 B 列数据验证（模板给了 B64:B66 三行），再到「表2」截断
    d0 = d1 = None
    for dv in ws.data_validations.dataValidation:
        ref = str(dv.sqref)
        if not ref.startswith("B"):
            continue
        if "管理层账表调整" not in str(dv.formula1 or ""):
            continue
        d0 = hdr + 1
        try:
            d1 = int("".join(ch for ch in ref.split(":")[-1] if ch.isdigit()))
        except ValueError:
            d1 = hdr + 3
        break
    if not d0:
        d0, d1 = hdr + 1, hdr + 3
    for r in range(d0, min(d1 + 8, ws.max_row + 1)):
        if "表2" in _txt(ws.cell(r, 2).value):
            d1 = max(d0, r - 1)
            break

    # 金额列统一样式：人工填调整时也是会计格式
    for r in range(d0, d1 + 1):
        ws.cell(r, 6).number_format = ACCT_FMT

    hcol = lay["cols"]["adj_aud"]
    rows = []
    for item in tab["rows"].values():
        row = item["row"]
        cell = ws.cell(row, hcol,
                       "=SUMIFS($F${a}:$F${b},$E${a}:$E${b},$C{r})".format(
                           a=d0, b=d1, r=row))
        cell.number_format = ACCT_FMT
        cell.alignment = Alignment(horizontal="right", vertical="center")
        rows.append(row)

    return {"title": title, "header": hdr, "d0": d0, "d1": d1,
            "hcol": hcol, "rows": rows}


def add_notes(ws, lines) -> None:
    """在 Lead 的 Notes 区写留痕。找不到 [A] 就追加在表1 下方，不硬塞。"""
    r0 = None
    for r in range(40, min(ws.max_row, 80) + 1):
        if "[A]" in _txt(ws.cell(r, 2).value) or "波动说明" in _txt(ws.cell(r, 2).value):
            r0 = r
            break
    if r0 is None:
        return
    for i, s in enumerate(lines):
        c = ws.cell(r0 + 1 + i, 2, s)
        c.font = Font(name="微软雅黑", size=8, color="595959")
