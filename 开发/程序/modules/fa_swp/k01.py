# -*- coding: utf-8 -*-
"""固定资产 SWP · K.01 Agree SL to GL 的填充。

模板四张表（行号是瑞云这版模板的；代码里全部按标签现场定位，不写死）：

    表1  R12-R49   原值 / 累计折旧 / 减值准备 / 净值 的后推 + TB 核对块
    表2  R55-R62   FA list 汇总（原值/折旧/减值/净值 × 类别）
    表3  R55-R62   表2 与表1 的差（模板公式，一个字不动，只要求全 0）
    表4  R75-R82   折旧费用与损益类科目勾稽

**符号口径（与旧 asset.py 相反，必须记牢）**：
    原值、累计折旧、减值准备**全部填正数**；净值 = 原值 − 折旧 − 减值；
    减少用负数（处置行、转出行）。

**哪些写公式、哪些写死值**：
    · 变动行写 SUMIFS 引用清单页 —— 改清单，K.01 跟着变；
    · 年初余额写公式引「上年 K.01」年末行 —— 同簿引用，不建外部工作簿链接；
    · TB 核对块写 SUMIFS 引「TB 固定资产」页 —— 只取未标记重复的行；
    · 来源说明一律写在 Notes 区。**不写单元格批注**：公式里已经带页名和单元格，
      再挂一遍批注只是把同一句话说两次，还挡视线（用户 2026-11-13 口径）。
"""

from __future__ import annotations

from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from . import embed, moves, prep, xlrows, xlshrink
from .fmt import ACCT_FMT, MONEY_FMT
SHEET_K01 = "K.01 Agree SL to GL"
SHEET_MV = "新增清单"          # 模板原页名，页内标题写「变动清单」
SHEET_DP = "处置清单"
SHEET_FA = "FA list"

SECTIONS = (("fa", ("原值",)), ("dep", ("累计折旧",)), ("imp", ("减值准备",)),
            ("nbv", ("净值",)), ("mv", ("变动",)))
ROWS = (("beg", ("年初余额",)), ("buy", ("购置",)), ("cip", ("在建工程转入",)),
        ("disp", ("处置或报废",)), ("reclass", ("内部重分类",)), ("blank", ("[…]",)),
        ("end", ("年末余额",)), ("prov", ("计提",)))


def _txt(v) -> str:
    return str(v or "").replace(" ", "").replace("\n", "").replace("\u3000", "")


def layout(ws) -> dict:
    """按标签找出表1 的类别列块与行、TB 核对块、表2 表头、表4 各行。"""
    cat_row = None
    for r in range(1, 30):
        if any("固定资产类别" in _txt(ws.cell(r, c).value) for c in range(1, 8)):
            cat_row = r
            break
    if cat_row is None:
        raise ValueError("K.01 里找不到「固定资产类别」表头行")

    # 扫到「合计」为止，不设死上限。原来写死 min(max_column, 40)，
    # 12 个类别时列块名落在第 39 列，正好贴边 —— 再多一块就被静默截断，
    # 表1 只认得前 11 块，而合计照样平。
    cat_cols = []
    limit = min(ws.max_column, 200)
    for c in range(1, limit + 1):
        t = _txt(ws.cell(cat_row, c).value)
        if t == "合计":
            break
        if not t or "固定资产类别" in t:
            continue
        cat_cols.append(c)
    if not cat_cols:
        raise ValueError("K.01 表头行 %d 上找不到类别列块（扫了 1~%d 列）" % (cat_row, limit))

    # 段落怎么判：模板把「原值」「累计折旧」「减值准备」写在**段中间**（B15/B21/B27），
    # 段首那几行 B 列是空的，靠「最近一个 B 列标签」会把购置、在建转入判丢。
    # 所以改用 C 列「年初余额」的出现次序切段 —— 原值/折旧/减值/净值四段
    # 各以一次「年初余额」开头，顺序固定。
    # 段落怎么判：模板把「原值」「累计折旧」「减值准备」写在**段中间**（B15/B21/B27），
    # 段首那几行 B 列是空的，靠「最近一个 B 列标签」会把购置、在建转入判丢。
    # 所以改用 C 列「年初余额」的出现次序切段 —— 原值/折旧/减值/净值四段
    # 各以一次「年初余额」开头，顺序固定。
    #
    # 段内的变动行用**有序列表** `mv`，不用「标签→行号」的扁平字典：
    # 三段各有一行叫「处置或报废」，扁平字典会互相覆盖，动态插行后更是没法区分。
    sec_order = ("fa", "dep", "imp", "nbv")
    sec, section, si, mode = {}, None, -1, None
    for r in range(cat_row + 1, cat_row + 45):
        c_txt = _txt(ws.cell(r, 3).value)
        if not c_txt:
            continue
        if "年初余额" in c_txt:
            si += 1
            section = sec_order[si] if si < len(sec_order) else None
            if section:
                sec[section] = {"beg": r, "mv": [], "blank": None, "end": None}
            mode = "beg"
            continue
        if section is None:
            continue
        cur = sec[section]
        if c_txt.startswith("[") and cur["blank"] is None:
            cur["blank"] = r
            mode = "blank"
            continue
        if "年末余额" in c_txt:
            cur["end"] = r
            mode = "end"
            continue
        if mode == "mv" or (mode == "beg" and cur["mv"] == []):
            cur["mv"].append((r, c_txt))
            mode = "mv"

    need = ("fa", "dep", "imp")
    for k in need:
        cur = sec.get(k) or {}
        if not cur.get("beg") or not cur.get("end") or cur.get("blank") is None:
            raise ValueError("K.01 的 %s 段没解析全（beg/blank/end）：%s" % (k, cur))

    # 兼容旧键：embed / checks 还在用 rows[段][key] 取行号
    rows = {}
    keymap = (("beg", "beg"), ("buy", "购置"), ("cip", "在建工程转入"),
              ("disp", "处置或报废"), ("reclass", "内部重分类"), ("end", "end"),
              ("prov", "计提"), ("blank", "blank"))
    for k in need:
        cur = sec[k]
        d = {"beg": cur["beg"], "end": cur["end"], "blank": cur["blank"]}
        for alias, label in keymap:
            if label in ("beg", "end", "blank"):
                continue
            for r0, txt0 in cur["mv"]:
                if txt0 == label:
                    d[alias] = r0
                    break
        rows[k] = d
    for r0, txt0 in sec["nbv"]["mv"]:
        rows.setdefault("nbv", {})[txt0] = r0
    for k in ("beg", "end"):
        if sec["nbv"].get(k):
            rows.setdefault("nbv", {})[k] = sec["nbv"][k]

    tb = {}
    for r in range(rows["fa"]["end"], min(ws.max_row, rows["fa"]["end"] + 40) + 1):
        b = _txt(ws.cell(r, 2).value)
        if not b.startswith("TB-"):
            continue
        if "原值" in b:
            tb["fa"] = r
        elif "累计折旧" in b:
            tb["dep"] = r
        elif "减值" in b:
            tb["imp"] = r
        elif "净值" in b:
            tb["nbv"] = r
    if set(tb) < {"fa", "dep", "imp"}:
        raise ValueError("K.01 的 TB 核对块找不到（找过 %d~%d）"
                         % (rows["fa"]["end"], rows["fa"]["end"] + 20))

    t2 = None
    for r in range(rows["fa"]["end"], ws.max_row + 1):
        if "固定资产类别" in _txt(ws.cell(r, 2).value):
            t2 = r
            break
    if t2 is None:
        raise ValueError("K.01 里找不到表2 表头")

    t4 = {}
    for r in range(t2, ws.max_row + 1):
        t = _txt(ws.cell(r, 5).value)
        if t and t not in t4:
            t4[t] = r
    if "合计" not in t4:
        raise ValueError("K.01 里找不到表4（E 列没有「合计」）")

    return {"cat_row": cat_row, "cat_cols": cat_cols, "rows": rows, "sec": sec,
            "tb": tb, "t2": t2, "t4": t4}


def _ref(sheet: str, pair) -> str:
    return "'%s'!$%s:$%s" % (sheet, pair[0], pair[1])


def _sumifs(sheet, sum_pair, cat_pair, cat_abs, way_pair=None, way=None) -> str:
    f = "SUMIFS(%s,%s,%s" % (_ref(sheet, sum_pair), _ref(sheet, cat_pair), cat_abs)
    if way_pair and way:
        f += ",%s,\"%s\"" % (_ref(sheet, way_pair), way)
    return f + ")"


def _way_formula(way, cat_abs: str, cat_map: dict):
    """一个方式 → 表1 某个格子的公式。

    「表1 用词」与「清单用词」的差异（内部重分类 / 内部转出）只在这里翻译一次，
    靠 ``Way.list_way``。写死字符串的话，某段会恒 0 而合计看着还是对的。
    """
    if way.src == "mv":
        f = _sumifs(SHEET_MV, cat_map["mv"][way.field], cat_map["mv"]["cat"], cat_abs,
                    cat_map["mv"]["way"], way.list_way or way.name)
    elif way.src == "dp":
        f = _sumifs(SHEET_DP, cat_map["dp"][way.field], cat_map["dp"]["cat"], cat_abs,
                    cat_map["dp"]["way"], way.list_way or way.name)
    elif way.src == "fa":
        f = _sumifs(SHEET_FA, cat_map["fa"][way.field], cat_map["fa"]["cat"], cat_abs)
    else:                                   # const：本年固定为 0（如减值计提）
        return 0
    # 必须带前导 "="。少了它 Excel 会把整串当文本，列显示 0、合计照平，
    # 看上去像「本期没有这笔业务」，实际是公式根本没生效。
    return "=" + ("-" if way.sign < 0 else "") + f


def plan_inserts(lay: dict, mvs: dict) -> list[dict]:
    """算出要在哪一行前插入哪些新方式。

    规则：模板原有方式**一行都不删**，数据里出现的新方式按 ``WAY_ORDER``
    的规范顺序插到「它前面最近一个已有行」的下一行 —— 得到
    ``购置 → 在建工程转入 → 改造转入 → 转固冲回 → 处置或报废 → 内部重分类``，
    先增后减。同一插入点的多个方式合成一次插入，保持规范顺序。
    """
    out = []
    for sec_key in ("fa", "dep", "imp"):
        s = lay["sec"][sec_key]
        existing = [w for _r, w in s["mv"]]
        order = moves.WAY_ORDER[sec_key]
        data = mvs["plan"].get(sec_key, set())
        target = [w for w in order if w in existing or w in data]
        target += [w for w in existing if w not in order]      # 模板自定义行
        new = [w for w in target if w not in existing]
        if not new:
            continue
        # mv 是 [(行号, 标签)]，这里要的是「标签 → 行号」；写反了会让每个新方式
        # 都往前找不到锚点、全部落到段首（插在「年初余额」下面一行）。
        got = {w: r0 for r0, w in s["mv"]}
        pos = {w: i for i, w in enumerate(target)}
        groups, order_at = {}, []
        for w in new:
            at = None
            for j in range(pos[w] - 1, -1, -1):
                if target[j] in got:
                    at = got[target[j]] + 1
                    break
            if at is None:
                at = s["beg"] + 1
            if at not in groups:
                groups[at] = []
                order_at.append(at)
            groups[at].append(w)
        for at in order_at:
            out.append({"sec": sec_key, "at": at, "labels": groups[at]})
    return out


def rewrite_year_ends(ws, lay) -> dict:
    """三段年末行 = 年初 + 全部变动行（含 […] 行），逐列重写。

    必须在「插行 + 公式平移」之后做：平移会把原来的 ``+E17`` 变成 ``+E19``，
    跳过新插的两行；重写直接覆盖成连续区间，把新行纳入。
    """
    made = {}
    for sec_key in ("fa", "dep", "imp"):
        s = lay["sec"][sec_key]
        made[sec_key] = xlrows.rewrite_plus_sum_row(ws, s["end"], s["beg"], s["blank"])
    return made


def plan_blocks(lay: dict, cats: list[str], *, hard_max: int = 12,
                block_size: int = 3) -> dict:
    """算出类别列块计划。**在任何列手术之前一次算完。**

    三种情形，互斥：

    - **类别数 < 模板块数**（瑞云 4 类、模板 7 块）→ ``delete_n > 0``，
      把「实际类别之后、合计块之前」的那些空块**真正删掉**。留着它们就是
      K.01 中间挂着 3 组 `【…】`，看着像还没填完。
    - **相等** → 两个数都是 0，只把 `【…】` 改成真名 + 写公式。
    - **类别数 > 模板块数** → ``n_new > 0``，在「最后一块的审定列 + 1」前插列
      （7 块时是第 26 列 Z）。为什么不插在 E 前面：那里是第一个块，右边 500 多个
      公式格全要跟着动，而插在末尾只影响 56 格，同一件事差一个数量级。

    ``blocks`` 一律**只含实际类别**（长度 == ``n_need``）。以前它会被补到
    ``max(n_have, n_need)`` 并带一串 ``cat=None`` 的空块，于是横向合计「恰好覆盖」
    校验里混进了不存在的列，删列时也分不清哪些该删。

    **绝不能「插一组 → 重跑 layout → 再插第二组」**：第一次插完，第二组的
    目标列号已经变了，按原计划再插就会错位，而且错位后合计仍然平。
    """
    tpl = list(lay["cat_cols"])
    n_have = len(tpl)
    n_need = len(cats)
    if n_need > hard_max:
        raise ValueError(
            "数据有 %d 个资产类别（%s），单张 K.01 最多支持 %d 个。"
            "请合并同类，或拆成两期底稿。" % (n_need, "、".join(cats), hard_max))
    n_new = max(0, n_need - n_have)
    n_del = max(0, n_have - n_need)
    first_money = tpl[0] - 1
    blocks = [{
        "i": i,
        "money": first_money + block_size * i,
        "name": first_money + block_size * i + 1,
        "aud": first_money + block_size * i + 2,
        "cat": cats[i],
    } for i in range(n_need)]
    return {
        "n_have": n_have, "n_need": n_need, "n_all": n_need, "n_new": n_new,
        "insert_at": tpl[-1] + 2,          # 最后一块审定列 + 1 = 合计块首列
        "insert_n": n_new * block_size,
        # 要删的起始列 = 第一个多余空块的账面列（4 类时第 17 列 Q），
        # 删的列数 = 空块数 × 3（9 列），正好把 Q..Y 整段拿掉，
        # 合计块从 Z/AA/AB/AC 左移到 Q/R/S/T。
        "delete_at": first_money + block_size * n_need if n_del else 0,
        "delete_n": n_del * block_size,
        "blocks": blocks,
    }


def total_cols(lay: dict) -> dict:
    """合计块的四列：账面 / 调整 / 审定 / CHECK。"""
    book = lay["cat_cols"][-1] + 2
    return {"book": book, "adj": book + 1, "aud": book + 2, "check": book + 3}


def expand_table23(ws, wb, lay, blocks) -> dict:
    """表2 / 表3 的行数对齐到类别数（多则插、少则删）。

    模板给 7 行数据位（55–61）+ 1 行合计（62）。两个方向都会出问题：

    - **类别多于 7 个**：第 8 个没地方放，而 `C62=SUM(C55:C61)` 不会自己扩 ——
      表3 的差异会**恰好等于第 8 块的金额**；如果同时把表1 的横向合计也写漏了，
      两边一起少，差异恒 0、CHECK 恒 0，看起来全绿。
    - **类别少于 7 个**（瑞云 4 类）：后 3 行是空类别，类别名 `=R10/=U10/=X10`
      指向的列已经被删了，总计 `SUM(C55:C61)` 也把空行括了进去。这几行必须删掉，
      不能只清空 —— 留 3 行空白在「合计」上方，看底稿的人一定会问。

    删行后必须补三件事，`delete_rows` 一件都不做：公式三态平移（含合计区间收缩）、
    跨表引用、条件格式 sqref。行高与合并区由 `delete_rows_keep` 负责。
    """
    first = lay["t2"] + 1
    sum_row = None
    for r in range(first, min(first + 40, ws.max_row) + 1):
        if _txt(ws.cell(r, 2).value) == "合计":
            sum_row = r
            break
    if sum_row is None:
        raise ValueError("K.01 表2 找不到「合计」行（在 B 列找过 %d~%d）" % (first, first + 40))

    have = sum_row - first
    extra = len(blocks) - have

    if extra > 0:
        xlrows.insert_rows_keep(ws, sum_row, extra)
        # 表4 在下面，它的 =SUM(F75:F79) / =F80-F81 会跟着平移；=Z20 因为 20 行在上方
        # 而不动，实测这个 pass 天然正确，不需要手工重写。
        xlrows.shift_sheet_formulas(ws, sum_row, extra)
        for i in range(extra):
            xlrows.copy_row_style(ws, sum_row - 1, sum_row + i)
        return {"extra": extra, "deleted": 0, "sum_row": sum_row + extra, "have": have}

    if extra < 0:
        d = -extra
        at = first + len(blocks)          # 第一个多余空行；合计行随后落到 at
        xlshrink.delete_rows_keep(ws, at, d)
        xlshrink.shift_sheet_formulas_del(ws, "row", at, d)
        xlshrink.shift_workbook_refs_del(wb, ws.title, "row", at, d)
        xlshrink.shift_cf_del(ws, "row", at, d)
        return {"extra": 0, "deleted": d, "sum_row": at, "have": have}

    return {"extra": 0, "deleted": 0, "sum_row": sum_row, "have": have}


def fill_table1(ws, lay, prep_res, mvs, prior_cats, cat_map, prior_file,
                prior_sheet=None, tb=None, blocks=None) -> dict:
    """表1：类别名、年初、各变动行、TB 核对块。

    行数与列数都由数据决定：变动行由 `plan_inserts` 插出来，类别块由
    `plan_blocks` 算出来。**不做「只填有数据的类别」的过滤** —— 那个过滤会
    把「期初期末都是 0、中间全处置掉了」的类别整列丢掉，而它的变动金额不为 0，
    丢了列、合计仍平，属于最难查的一类错。
    """
    if blocks:
        used = [b["cat"] for b in blocks if b["cat"]]
        name_cols = [b["name"] for b in blocks if b["cat"]]
    else:
        used = list(prep_res["cat_order"])
        name_cols = list(lay["cat_cols"])[:len(used)]

    if len(used) > len(name_cols):
        raise ValueError(
            "类别有 %d 个，可用列块只有 %d 个。插列计划没执行或执行失败，"
            "拒绝生成 —— 少一块时合计仍然平，看不出来。类别：%s"
            % (len(used), len(name_cols), used))

    info = {"cats": [], "cells": {}}
    r = lay["rows"]

    for i, cat in enumerate(used):
        name_col, money_col = name_cols[i], name_cols[i] - 1
        head = ws.cell(lay["cat_row"], name_col, cat)
        head.font = Font(bold=True, size=10)
        head.alignment = Alignment(horizontal="center", vertical="center")
        abs_cat = "$%s$%d" % (get_column_letter(name_col), lay["cat_row"])
        info["cats"].append(cat)
        info["cells"][cat] = money_col

        m = prep_res["cats"][cat]
        p = prior_cats[cat]
        L0 = get_column_letter(money_col)

        def put(row, value, note=None):
            """写一格金额。``note`` 是代码内的来源说明，**不再写单元格批注** ——
            公式本身已带页名与单元格，来源一眼可见（用户 2026-11-13 口径）。"""
            cell = ws.cell(row, money_col, value)
            cell.number_format = ACCT_FMT
            return cell

        # 年初：优先引用「上年 K.01」页的年末行（可追溯、可点开看来源）；
        # 没给 prior_sheet 才退回写死值（旧行为，保证旧调用点不变）。
        if prior_sheet:
            qk = embed.q(prior_sheet)
            put(r["fa"]["beg"], "=%s!$%s$%d" % (qk, L0, r["fa"]["end"]),
                "引用：%s 年末余额行" % prior_sheet)
            put(r["dep"]["beg"], "=%s!$%s$%d" % (qk, L0, r["dep"]["end"]),
                "引用：%s 年末余额行" % prior_sheet)
            put(r["imp"]["beg"], "=%s!$%s$%d" % (qk, L0, r["imp"]["end"]),
                "引用：%s 年末余额行" % prior_sheet)
        else:
            put(r["fa"]["beg"], float(p["fa"]), "来源：%s 表1 合计行原值" % prior_file)
            put(r["dep"]["beg"], float(p["dep"]), "来源：%s 表1 合计行累计折旧" % prior_file)
            put(r["imp"]["beg"], float(p["imp"]), "来源：%s 表1 合计行减值准备" % prior_file)

        # ---- 变动行：遍历三段的有序 mv，按「方式」生成 SUMIFS。
        # 表1 有哪几行由 plan_inserts 插出来的结构决定，这里**不写死任何行号**，
        # 也不写死「购置」「处置或报废」这些词 —— 全从 moves.WAY_DEFS 取。
        # 模板原有但本期金额为 0 的方式（如减值计提）照样会有一行，显示 0。
        for sec_key in ("fa", "dep", "imp"):
            for row, label in lay["sec"][sec_key]["mv"]:
                way = moves.WAY_DEFS.get((sec_key, label))
                if way is None:
                    continue          # 模板占位行或项目组自定义行，留空由人填
                put(row, _way_formula(way, abs_cat, cat_map),
                    "%s：取%s" % (way.name, {"mv": "变动清单", "dp": "处置清单",
                                            "fa": "FA list", "const": "固定 0"}[way.src]))

        # TB 核对块：优先 SUMIFS 引用「TB 固定资产」页（只取「是否计入汇总=是」的行，
        # 1602/1603 取反），否则退回写死值。
        if tb:
            qt = embed.q(tb["sheet"].title)
            cc = tb["cols"]

            def tb_col(key):
                # 必须两段都写全：$H$5:$H$48。写成 $H$5:$48 缺第二段的列字母，
                # Excel 会直接判为无效引用（看着像公式，打开却是 #NAME? 或报文件损坏）。
                col = cc[key]
                return "%s!$%s$%d:$%s$%d" % (qt, col, tb["first"], col, tb["last"])
            for acct, row, sign in (("1601", lay["tb"]["fa"], ""),
                                    ("1602", lay["tb"]["dep"], "-"),
                                    ("1603", lay["tb"]["imp"], "-")):
                # 条件写 "1601*" 而不是 "1601"：客户用明细科目时（1601.01 固定资产_电子设备）
                # 精确匹配一行都取不到，核对行恒为 0，而「差异」行会等于整个年末余额 ——
                # 看着像巨额差异，其实是漏了通配符。
                put(row, '=%sSUMIFS(%s,%s,"%s*",%s,%s,%s,"是")' % (
                    sign, tb_col("end_bal"), tb_col("acct_code"), acct,
                    tb_col("cat"), abs_cat, tb_col("counted")),
                    "引用 TB 页：科目 %s 及其明细科目、类别=本列表头、只取未标记重复的行" % acct)
        else:
            put(lay["tb"]["fa"], float(m["fa_end"]), "TB 1601 期末（已去重）")
            put(lay["tb"]["dep"], float(m["dep_end"]), "TB 1602 期末（已去重，取绝对值）")
            put(lay["tb"]["imp"], float(m["imp_end"]), "TB 1603 期末（已去重，取绝对值）")
        if "nbv" in lay["tb"]:
            L = get_column_letter(money_col)
            put(lay["tb"]["nbv"], "=%s%d-%s%d-%s%d" % (
                L, lay["tb"]["fa"], L, lay["tb"]["dep"], L, lay["tb"]["imp"]))

    return info


def fill_table23(ws, lay, cat_map, blocks) -> dict:
    """表2 取 FA list 汇总；表3 与表1 逐块比对。

    全量重写而不是「只补新增的几行」：4 个类别时生成出来的公式与模板**逐字相同**
    （回归基准因此不变），8 个类别时多出来的行也自动带上正确的块引用。
    增量补写省不了多少事，却容易漏掉合计行的区间。
    """
    fa = cat_map["fa"]
    got = [b for b in blocks if b["cat"]]
    first = lay["t2"] + 1
    total_book = total_cols(lay)["book"]

    # 合计行必须**扫出来**，不能 first + 类别数：模板给了 7 行数据位（57–63）
    # 但合计在 64。按类别数算的话，4 个类别会把合计写到第 61 行、盖掉一个
    # 数据位，而第 64 行的老合计还留着 —— 表2 出现两个合计，表3 差异恒非 0。
    sum_row = None
    for r in range(first, min(first + 40, ws.max_row) + 1):
        if _txt(ws.cell(r, 2).value) == "合计":
            sum_row = r
            break
    if sum_row is None:
        raise ValueError("K.01 表2 找不到「合计」行（B 列找过 %d~%d）" % (first, first + 40))

    for i, b in enumerate(got):
        r = first + i
        ws.cell(r, 2, "=%s%d" % (get_column_letter(b["name"]), lay["cat_row"]))
        for field, col in (("fa_end", 3), ("dep_end", 4), ("imp_end", 5)):
            ws.cell(r, col, "=" + _sumifs(
                SHEET_FA, fa[field], fa["cat"], "$B$%d" % r)).number_format = MONEY_FMT
        ws.cell(r, 6, "=C%d-D%d-E%d" % (r, r, r)).number_format = MONEY_FMT

    for col in (3, 4, 5, 6):
        L = get_column_letter(col)
        ws.cell(sum_row, col, "=SUM(%s%d:%s%d)" % (L, first, L, sum_row - 1)
                ).number_format = MONEY_FMT

    # 表3：逐块比对（列 8~11 = H~K），末行对合计块
    end_rows = {"fa": lay["rows"]["fa"]["end"], "dep": lay["rows"]["dep"]["end"],
                "imp": lay["rows"]["imp"]["end"], "nbv": lay["rows"]["nbv"]["end"]}
    srcs = {"fa": 3, "dep": 4, "imp": 5, "nbv": 6}
    for i, b in enumerate(got):
        r = first + i
        for j, key in enumerate(("fa", "dep", "imp", "nbv")):
            ws.cell(r, 8 + j, "=%s%d-%s%d" % (
                get_column_letter(srcs[key]), r,
                get_column_letter(b["money"]), end_rows[key])).number_format = MONEY_FMT
    for j, key in enumerate(("fa", "dep", "imp", "nbv")):
        ws.cell(sum_row, 8 + j, "=%s%d-%s%d" % (
            get_column_letter(srcs[key]), sum_row,
            get_column_letter(total_book), end_rows[key])).number_format = MONEY_FMT

    return {"rows": len(got), "sum_row": sum_row}


def fill_table4(ws, lay, t4res, customer="") -> None:
    """表4：折旧费用与损益类科目勾稽。

    **不插行** —— 模板 F81 是 ``=Z20``，插行会让它指错格；这 5 行是按费用性质
    设计的，不是按科目，多出来的科目合并进对应行、编码列全列出。
    """
    order = ("制造费用", "销售费用", "管理费用", "研发费用", "其他")
    # 模板第 5 行 E 列写的是占位符「[…]」，不是「其他」。找不到就用它，
    # 否则 6402 其他业务支出那 11 万永远进不了 F80，表4 差异恒为 11 万。
    other_row = lay["t4"].get("其他") or lay["t4"].get("[…]")
    for name in order:
        r = other_row if name == "其他" else lay["t4"].get(name)
        if not r:
            continue
        rows = t4res["details"].get(name, [])
        if rows:
            # 客户名从项目信息取。原来写死「瑞云集团（5 主体合并）」，换客户会带着
            # 上一家的名字出现在交付件里 —— 这种错一眼就会被客户看见。
            ws.cell(r, 2, customer or "")
            ws.cell(r, 3, "、".join(x["code"] for x in rows))
            ws.cell(r, 5, "、".join(x["name"] for x in rows))
        if name == "其他" and ws.cell(r, 5).value is None:
            ws.cell(r, 5, "其他业务支出-折旧费")
        ws.cell(r, 6, float(t4res["buckets"][name])).number_format = MONEY_FMT
        for c in (2, 3, 5):
            ws.cell(r, c).alignment = Alignment(horizontal="left", vertical="center")
