# -*- coding: utf-8 -*-
"""固定资产 SWP · 校验不变量。

**只在数据层判，不读 Excel 的缓存值** —— openpyxl 写进去的是公式字符串，
没有缓存值可读；要读 Excel 算出来的数就必须开本机 Excel 重算（见 recalc.py）。
这里判的是「喂给 Excel 的数据本身闭不闭合」，Excel 那一层交给重算脚本。

任何一条不过就 ``ok=False``，CLI 退出码非 0 —— 不出「看起来成功」的底稿。
"""

from __future__ import annotations

import re
from decimal import Decimal as D

from . import k01, moves, prep

TOL = D("0.01")


def _sum(rows, field) -> D:
    return sum((r[field] for r in rows), D("0"))


def check(prep_res, mvs, syn) -> dict:
    t = prep_res["totals"]
    def sec_sum(sec, kinds=None):
        """某段（可按 add/reduce 过滤）的合计。不再写死 R13..R27 —— 方式是可变的。"""
        out = D("0")
        for w in moves.WAY_ORDER[sec]:
            wd = moves.WAY_DEFS.get((sec, w))
            if wd and (kinds is None or wd.kind in kinds):
                out += mvs["totals"].get((sec, w), D("0"))
        return out
    items = []

    def ck(name, got, want, note="", soft=False):
        ok = abs(D(str(got)) - D(str(want))) <= TOL
        items.append({"name": name, "ok": ok, "got": str(got), "want": str(want),
                      "note": note, "soft": bool(soft)})

    def ckb(name, cond, extra=""):
        """布尔断言（不涉及金额的那些）。"""
        items.append({"name": name, "ok": bool(cond), "got": str(extra) if not cond else "",
                      "want": "", "note": ""})

    # 1) 原值：年初 + 购置 + 转入 − 处置 − 转出 = 年末
    ck("原值闭合", t["fa_beg"] + sec_sum("fa"), t["fa_end"],
       "K.01 表1 原值各方式行加总")
    # 2) 折旧
    ck("折旧闭合", t["dep_beg"] + sec_sum("dep"), t["dep_end"], "年初 + 各方式行")
    # 3) 减值
    ck("减值闭合", t["imp_beg"] + sec_sum("imp"), t["imp_end"], "年初 + 各方式行")

    # 4) 清单合计 = K.01 行
    fa = syn["fa"]
    ck("FA 期初原值", _sum(fa, "fa_beg"), t["fa_beg"])
    ck("FA 期末原值", _sum(fa, "fa_end"), t["fa_end"])
    ck("FA 期初折旧", _sum(fa, "dep_beg"), t["dep_beg"])
    ck("FA 期末折旧", _sum(fa, "dep_end"), t["dep_end"])
    ck("FA 期初减值", _sum(fa, "imp_beg"), t["imp_beg"])
    ck("FA 期末减值", _sum(fa, "imp_end"), t["imp_end"])
    prov = moves.amounts(mvs, "dep", "计提")
    dep_disp = -moves.amounts(mvs, "dep", "处置或报废")
    imp_disp = -moves.amounts(mvs, "imp", "处置或报废")
    ck("FA 本期计提", _sum(fa, "prov_dep"), prov)
    ck("FA 本期减少折旧", _sum(fa, "reduce_dep"), dep_disp)

    mv, dp = syn["moves"], syn["disposals"]
    ck("变动清单 = 各增加方式", _sum(mv, "orig"), sec_sum("fa", ("add",)))
    ck("处置清单 = 各减少方式", _sum(dp, "orig"), -sec_sum("fa", ("reduce",)))
    ck("处置清单折旧", _sum(dp, "dep"), dep_disp)
    ck("处置清单减值", _sum(dp, "imp"), imp_disp)

    # 5) 表4
    # 不设硬门禁：折旧费用按「科目名含折旧」归集，与 K.01 计提行的差额是
    # 真实业务去向（制造业折旧先入生产成本、随存货结转），或账上折旧费科目
    # 记串。这类差必须由人看，但不能因此不出稿 —— 底稿就是用来放差异的。
    ck("表4 = 计提", mvs["table4"]["total"], prov,
       "折旧费用与计提行勾稽；差额需在表4 差异行解释", soft=True)

    # 新增：方式完备 + 清单方式集合与表1 计划一致
    bad_way = sorted(k for k in mvs["totals"] if k not in moves.WAY_DEFS)
    ckb("方式都在 WAY_DEFS 里", not bad_way, bad_way)
    list_ways = {r["way"] for r in syn["moves"]} | {r["way"] for r in syn["disposals"]}
    plan_ways = set()
    for sec, s_ways in mvs["plan"].items():
        if sec != "fa":
            continue
        for w in s_ways:
            wd = moves.WAY_DEFS.get((sec, w))
            plan_ways.add(wd.list_way or wd.name if wd else w)
    ckb("清单方式 == 表1 计划方式", list_ways == plan_ways, (sorted(list_ways), sorted(plan_ways)))

    # 6) 净值
    ck("净值 期初", t["fa_beg"] - t["dep_beg"] - t["imp_beg"], prep_res["net_beg"])
    ck("净值 期末", t["fa_end"] - t["dep_end"] - t["imp_end"], prep_res["net_end"])

    # 7) 明细层
    # 红字冲回的凭证（原值为负，如「调整上年误转固」）本身就是负数行，
    # 不算「净值异常」—— 它冲减的是账上已有的资产，账务上就该是负的。
    red = [r for r in fa if r["add"] < 0]
    bad_nbv = [r for r in fa if r not in red and r["nbv_end"] < -TOL]
    bad_dep = [r for r in fa if r not in red and r["dep_end"] < -TOL]
    over = [r for r in fa if r not in red and r["life"] > 0 and (r["fa_beg"] + r["add"]) > 0
            and r["dep_end"] > (r["fa_beg"] + r["add"]) * D(str(1 - r["resid"])) + TOL]
    items.append({"name": "明细净值非负", "ok": not bad_nbv, "got": len(bad_nbv), "want": 0})
    items.append({"name": "明细折旧非负", "ok": not bad_dep, "got": len(bad_dep), "want": 0})
    items.append({"name": "折旧不超可折旧额", "ok": not over, "got": len(over), "want": 0})

    # 8) 类别
    cats = {r["cat"] for r in fa}
    order = list(prep_res["cat_order"])
    items.append({"name": "类别合法", "ok": cats <= set(order) and "未分类" not in cats,
                  "got": sorted(cats), "want": order})

    fails = [i for i in items if not i["ok"] and not i.get("soft")]
    return {"ok": not fails, "items": items, "fails": fails}


def check_prior(prep_res, prior_cats) -> dict:
    """上年 K.01 的年末必须等于本年 TB 期初，逐类别逐科目。"""
    fails = []
    for c in prep_res.get("cat_order") or list(prep_res["cats"]):
        for f, key in (("fa", "fa_beg"), ("dep", "dep_beg"), ("imp", "imp_beg")):
            a, b = float(prior_cats[c][f]), float(prep_res["cats"][c][key])
            if abs(a - b) > 0.005:
                fails.append("%s %s: 上年 %.2f ≠ 本年TB期初 %.2f" % (c, f, a, b))
    return {"ok": not fails, "fails": fails}


def check_cat_union(prep_res, mvs) -> dict:
    """TB 类别集合 vs 序时账类别集合，只在一侧的类别逐个列证据。

    **为什么必须互校**：K.01 的年初来自 TB 期末（上年页），变动来自序时账，
    期末 = 年初 + 变动。两边类别集不一致时，期末必然对不上，但表现只是
    「差异若干元」—— 看起来像金额算错，实际根因是分类口径分叉，而口径这种
    东西不看凭证号是查不出来的。这里提前拦住，并给两侧的行号/凭证号。
    """
    tb = set(prep_res["cats"])
    jnl = set(mvs.get("cat_seen") or ())
    only_tb = sorted(tb - jnl)
    only_jnl = sorted(jnl - tb)

    rows_tb = {}
    for c in only_tb:
        rows_tb[c] = [
            {"company": r["company_name"], "acct": r["acct_code"],
             "src_row": r["src_row"],
             "open": str(r["open_bal"]), "end": str(r["end_bal"])}
            for r in prep_res["rows"] if prep.category(r) == c
        ][:8]
    rows_jnl = {}
    for c in only_jnl:
        rows_jnl[c] = [
            {"company": ln["company_name"], "acct": ln["acct_code"],
             "date": str(ln["vdate"]), "vno": str(ln["vno"]),
             "amount": str(ln["debit"] - ln["credit"]),
             "summary": (ln["summary"] or "")[:30]}
            for ln in mvs["lines"] if moves.line_cat(ln) == c
        ][:8]

    parts = []
    if only_jnl:
        parts.append("序时账出现 TB 没有的类别：%s（TB 无期初期末，K.01 年初会是 0）"
                     % "、".join(only_jnl))
    if only_tb:
        parts.append("TB 有、序时账没有发生额：%s（本期没动，正常，但 K.01 变动行会是 0）"
                     % "、".join(only_tb))

    # **只对 only_jnl 硬报错。** 反方向是正常业务：一个类别本期一笔业务都没有
    # （不提折旧、不购置、不处置），序时账里就一条记录都不会有，而 TB 里余额还在。
    # 硬要求两边对称会把这种客户直接挡在门外。
    return {"ok": not only_jnl, "only_tb": only_tb, "only_jnl": only_jnl,
            "rows_tb": rows_tb, "rows_jnl": rows_jnl,
            "message": "；".join(parts)}


def summary(res: dict) -> str:
    ok = sum(1 for i in res["items"] if i["ok"])
    n = len(res["items"])
    if res["ok"]:
        return "校验通过：%d/%d 项" % (ok, n)
    lines = ["校验未通过：%d/%d 项" % (ok, n)]
    for f in res["fails"]:
        lines.append("  ✗ %s  实际 %s  期望 %s  %s"
                     % (f["name"], f["got"], f["want"], f.get("note", "")))
    return "\n".join(lines)


# --------------------------------------------------------------- sheet 层校验

def check_sheets(prep_res, s: dict) -> dict:
    """读回刚写好的本年页 / 上年页 / TB 页，判 10 条。

    与 ``check()`` 正交：那个判「喂给 Excel 的数据对不对」，这个判「写进 Excel 的
    东西对不对」。在 ``wb.save()`` 之前跑，不过就不出文件。

    ``s`` 需要带：``k01_ws / k01_lay / lead_ws / lead_lay / prior_k01 / prior_lead /
    tb / cat_map``，都是内存里的对象，不读磁盘。
    """
    from openpyxl.utils import column_index_from_string
    from openpyxl.utils import get_column_letter as gl

    items = []

    def ck(name, cond, extra=""):
        items.append({"name": name, "ok": bool(cond), "extra": str(extra)})

    kws, lay = s["k01_ws"], s["k01_lay"]
    pws, plws = s["prior_k01"], s["prior_lead"]
    tb = s["tb"]
    cats = prep_res["cats"]
    # 用**全部**类别，不做「非零才要」的过滤。过滤会丢掉「期初期末都是 0、
    # 中间全处置掉了」的类别 —— 它的变动金额不为 0，而列一丢，合计仍平。
    order = list(prep_res["cat_order"])
    blocks = s.get("blocks") or []
    used = order

    # 1) 上年 K.01 的年末行逐格 == 本年期初
    bad = []
    for i, cat in enumerate(used):
        money = lay["cat_cols"][i] - 1
        for key, f in (("fa", "fa_beg"), ("dep", "dep_beg"), ("imp", "imp_beg")):
            row = lay["rows"][key]["end"]
            v = pws.cell(row, money).value
            if not isinstance(v, (int, float)) or abs(v - float(cats[cat][f])) > 0.005:
                bad.append("%s/%s=%r≠%s" % (cat, key, v, cats[cat][f]))
    ck("上年 K.01 年末行 = 本年期初（%d 格）" % (len(used) * 3), not bad, bad[:4])

    # 2/3) 上年页残留公式必须为 0
    def nf(ws):
        return sum(1 for c in ws._cells.values() if c.data_type == "f")
    ck("上年 K.01 无残留公式", nf(pws) == 0, nf(pws))
    ck("上年 Lead 无残留公式", nf(plws) == 0, nf(plws))

    # 4) 上年 Lead 的 I 列 == 三行合计
    lr = s["lead_lay"]["rows"]
    bad2 = []
    for key, f in (("fa", "fa"), ("dep", "dep"), ("imp", "imp")):
        row = lr[key]
        v = plws.cell(row, 9).value
        tot = sum(s["prior_cats"][c][f] for c in s["prior_cats"])
        if not isinstance(v, (int, float)) or abs(v - tot) > 0.005:
            bad2.append("%s I%d=%r≠%.2f" % (key, row, v, tot))
    ck("上年 Lead I 列 = 上年合计", not bad2, bad2[:3])

    # 5) 三方一致：本年 E12 公式指向上年页年末行，且值与上年 Lead I 列相等
    c0 = lay["cat_cols"][0] - 1
    L = gl(c0)
    f12 = str(kws.cell(lay["rows"]["fa"]["beg"], c0).value or "")
    ok5 = (f12.startswith("='上年 K.01") and ("$%s$%d" % (L, lay["rows"]["fa"]["end"])) in f12)
    # 比的是「合计」对「合计」：上年 K.01 的 E18 只是机器设备一列，
    # 要把所有已用类别横向加起来，才能和 Lead 的三行合计对。
    v_up = sum((pws.cell(lay["rows"]["fa"]["end"], lay["cat_cols"][i] - 1).value or 0)
               for i in range(len(used)))
    v_lead = plws.cell(lr["fa"], 9).value
    ok5 = ok5 and isinstance(v_lead, (int, float)) and abs(v_up - v_lead) < 0.005
    ck("本年年初 ↔ 上年 K.01 ↔ 上年 Lead 三方一致", ok5, f12[:48])

    # 6) TB 页行数与重复标记
    rows = tb["last"] - tb["first"] + 1
    mcol = [c for c in range(1, 20) if tb["cols"]["counted"] == gl(c)][0]
    n_no = sum(1 for r in range(tb["first"], tb["last"] + 1) if tb["sheet"].cell(r, mcol).value == "否")
    # 行数不能写死 44（那是瑞云的数）。按「去重后行数 + 被删的重复行」核，
    # 换客户、加类别都能自然跟上。
    want_rows = len(prep_res["rows"]) + len(s["prep_trace"]["dropped"])
    ck("TB 页 %d 行、重复 %d 行" % (want_rows, len(s["prep_trace"]["dropped"])),
       rows == want_rows and n_no == len(s["prep_trace"]["dropped"]), (rows, n_no, want_rows))

    # 7) 用 TB 页的数据在 Python 里做等价 SUMIFS，反算期末矩阵
    cc = tb["cols"]
    cat_col = [i for i in range(1, 20) if gl(i) == cc["cat"]][0]
    end_col = [i for i in range(1, 20) if gl(i) == cc["end_bal"]][0]
    code_col = [i for i in range(1, 20) if gl(i) == cc["acct_code"]][0]
    cnt_col = mcol
    calc = {c: {k: 0.0 for k in ("fa_end", "dep_end", "imp_end")} for c in order}
    for r in range(tb["first"], tb["last"] + 1):
        if tb["sheet"].cell(r, cnt_col).value != "是":
            continue
        cat = tb["sheet"].cell(r, cat_col).value
        code = str(tb["sheet"].cell(r, code_col).value or "")
        amt = tb["sheet"].cell(r, end_col).value or 0
        if cat not in calc:
            # 空类别 + 金额为 0 = 零余额主体的占位行（如瑞影广州分公司），
            # prep.aggregate 那边也是 skip_zero 处理，这里跟着跳过。
            # 但**非空**的未知类别必须报错 —— 那是分类口径分叉，跳过就是少算。
            if not cat and not amt:
                continue
            raise ValueError("TB 页出现未经确认的类别 %r（第 %d 行，金额 %s），"
                             "不能静默跳过 —— 跳过就是在少算" % (cat, r, amt))
        # 一级科目比对：客户可能用明细科目（1601.01 固定资产_电子设备），
        # 只写 ``code == "1601"`` 会一行都匹配不到，反算全 0 却不是真差异。
        top = code.split(".")[0].strip()
        if top == "1601":
            calc[cat]["fa_end"] += amt
        elif top == "1602":
            calc[cat]["dep_end"] -= amt
        elif top == "1603":
            calc[cat]["imp_end"] -= amt
    bad3 = [c for c in order for k in calc[c]
            if abs(calc[c][k] - float(cats[c][k])) > 0.01]
    ck("TB 页反算 == 期末矩阵（12 格）", not bad3, bad3[:4])

    # 8) 本年 TB 核对块的公式形态
    tb_ok = True
    for row_key, acct in (("fa", "1601"), ("dep", "1602"), ("imp", "1603")):
        v = str(kws.cell(lay["tb"][row_key], c0).value or "")
        form_ok = v.startswith("=SUMIFS('TB") or v.startswith("=-SUMIFS('TB")
        # 条件可能是 "1601" 也可能是 "1601*"（客户用明细科目时必须带通配符），
        # 所以只校验前缀，不要求闭合引号。
        if not (form_ok and ('"%s' % acct) in v and '"是"' in v):
            tb_ok = False
    ck("本年 TB 核对块 = SUMIFS 引用 TB 页", tb_ok)

    # 8.5) ★ 年末行必须「恰好」覆盖年初 + 全部变动行 + 占位行
    #
    # 这是「漏行但合计仍平」的唯一可编程防线。公式平移把 +E17 变成 +E19 会跳过
    # 新插的行、年末少算却照样能平 —— 只有逐行号比对才抓得住。
    from openpyxl.formula import Tokenizer

    def _row_refs(f):
        out = set()
        if not isinstance(f, str) or not f.startswith("="):
            return out
        for tok in Tokenizer(f).items:
            if tok.type == "OPERAND" and tok.subtype == "RANGE":
                m = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)$", tok.value)
                if m:
                    out.add(int(m.group(2)))
        return out

    bad_cover = []
    lay_sec = lay.get("sec", {})
    for sec_key in ("fa", "dep", "imp"):
        sc = lay_sec.get(sec_key)
        if not sc:
            continue
        want = set(range(sc["beg"], sc["blank"] + 1))
        # 列号必须从 layout 推，不能写死 5/28。4 个类别时合计块在 Q..T，
        # 写死 AB 会读到一块空列、返回空集合，「恰好相等」反而更容易过 —— 假通过。
        cols = ((lay["cat_cols"][0] - 1, "首块账面"),
                (k01.total_cols(lay)["aud"], "审定合计"))
        for col, tag in cols:
            got = _row_refs(kws.cell(sc["end"], col).value)
            if got != want:
                bad_cover.append("%s/%s 缺%s 多%s" % (
                    sec_key, tag, sorted(want - got), sorted(got - want)))
    ck("年末行恰好覆盖年初+全部变动行", not bad_cover, bad_cover[:3])

    # 8.6) 两页几何一致：本年插了几行，上年页必须一样
    geom = []
    prior_lay = s.get("prior_lay")
    if prior_lay:
        for sec_key in ("fa", "dep", "imp"):
            a, b = lay_sec.get(sec_key), prior_lay.get("sec", {}).get(sec_key)
            if not a or not b:
                continue
            if (a["beg"], a["blank"], a["end"]) != (b["beg"], b["blank"], b["end"]) \
                    or [w for _r, w in a["mv"]] != [w for _r, w in b["mv"]]:
                geom.append("%s: 本年 %s / 上年 %s" % (sec_key, a, b))
    ck("上年页与本年页几何一致", not geom, geom[:2])

    # 8.7) 列块维度 —— 这一组是「漏一块但合计仍平」唯一能自动发现的防线。
    #
    # 插列之后最危险的失效是：新块的列建出来了、数据也填了，但横向合计
    # `Z12=SUM(E12,H12,…)` 还是 7 项没带上它。这时 Z 少一块、AB（审定合计）
    # 也少一块，而 CHECK 比的是「AB − Lead」，两边一起少 → 恒 0；
    # 表3 比的也是同样的两个数 → 恒 0。文件能打开、没有错误值、全绿。
    bl = list(s.get("blocks") or [])
    got_blocks = [b for b in bl if b.get("cat")]
    k_names = list(lay.get("cat_cols") or [])
    p_names = list(prior_lay.get("cat_cols") or []) if prior_lay else k_names

    ck("K.01 列块数 == 类别数", len(k_names) == len(bl) and len(got_blocks) == len(order),
       (len(k_names), len(bl), len(got_blocks), len(order)))

    heads = [str(kws.cell(lay["cat_row"], c).value or "").strip() for c in k_names]
    want_cats = [b["cat"] for b in got_blocks]
    ck("列块表头顺序 == 类别顺序", heads[:len(want_cats)] == want_cats,
       (heads[:len(want_cats)], want_cats))

    ck("两页列块位置一致", k_names == p_names, (k_names, p_names))
    ck("两页列块数一致", len(k_names) == len(p_names), (len(k_names), len(p_names)))
    ck("两页合并区一致",
       sorted(map(str, kws.merged_cells.ranges)) == sorted(map(str, pws.merged_cells.ranges)),
       (sorted(map(str, kws.merged_cells.ranges)), sorted(map(str, pws.merged_cells.ranges))))

    def _col_refs(f):
        out = set()
        if not isinstance(f, str) or not f.startswith("="):
            return out
        for tok in Tokenizer(f).items:
            if tok.type == "OPERAND" and tok.subtype == "RANGE":
                m = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)$", tok.value)
                if m:
                    out.add(column_index_from_string(m.group(1)))
        return out

    total_book = k_names[-1] + 2 if k_names else 26
    bad_ac = []
    if got_blocks:
        # 表1 主体（年初行）三个合计列都要覆盖全部块
        for off, key in ((0, "money"), (1, "name"), (2, "aud")):
            want = {b[key] for b in bl}
            gotc = _col_refs(kws.cell(lay["rows"]["fa"]["beg"], total_book + off).value)
            if gotc != want:
                bad_ac.append("R%d/%s 缺%s 多%s" % (
                    lay["rows"]["fa"]["beg"], key,
                    sorted(want - gotc), sorted(gotc - want)))
        # TB 核对行只有**账面列**有合计（模板的 AA/AB 在这几行本来就是空的），
        # 所以只校 Z 列，不能拿三列一起去要求它
        want_book = {b["money"] for b in bl}
        for row in (lay["tb"]["fa"], lay["tb"]["dep"], lay["tb"]["imp"], lay["tb"]["nbv"]):
            gotc = _col_refs(kws.cell(row, total_book).value)
            if gotc != want_book:
                bad_ac.append("R%d/TB 缺%s 多%s" % (
                    row, sorted(want_book - gotc), sorted(gotc - want_book)))
    ck("横向合计恰好覆盖全部块列", not bad_ac, bad_ac[:3])

    # 8.8) 列结构：块与合计块之间不许夹空列，也不许残留 `【…】` 占位块。
    #
    # 留着的空块对 SUM 的贡献是 0，所有金额校验都会过；只有「列区间逐列点名」
    # 才看得出中间多了 3 组永远填不满的列。这是本次改造的核心验收。
    tc = k01.total_cols(lay)
    span = set(range(k_names[0] - 1, tc["check"] + 1)) if k_names else set()
    want_span = set()
    for b in bl:
        want_span |= {b["money"], b["name"], b["aud"]}
    want_span |= {tc["book"], tc["adj"], tc["aud"], tc["check"]}
    ck("表1 列区间内没有空列", span == want_span,
       sorted(span - want_span) or sorted(want_span - span))

    ph = [c for c in span if "【" in str(kws.cell(lay["cat_row"], c).value or "")]
    ck("表1 无残留【…】占位块", not ph, ph)

    ck("列块数 == 类别数 == 有名字的列数",
       len(k_names) == len(bl) == len(order)
       and all(str(kws.cell(lay["cat_row"], c).value or "").strip() for c in k_names),
       (len(k_names), len(bl), len(order)))

    # 8.9) 表2 / 表3 的行结构：数据行数 == 类别数，合计行紧接其后。
    t2_first = lay["t2"] + 1
    t2_sum = None
    for r in range(t2_first, min(t2_first + 40, kws.max_row) + 1):
        if str(kws.cell(r, 2).value or "").strip() == "合计":
            t2_sum = r
            break
    ck("表2 数据行数 == 类别数", t2_sum is not None and (t2_sum - t2_first) == len(bl),
       (t2_first, t2_sum, len(bl)))

    bad_t2 = []
    if t2_sum is not None:
        for i, b in enumerate(bl):
            v = str(kws.cell(t2_first + i, 2).value or "")
            if v != "=%s%d" % (gl(b["name"]), lay["cat_row"]):
                bad_t2.append((t2_first + i, v))
    ck("表2 逐行引用对应块表头", not bad_t2, bad_t2[:3])

    # 9) 本年 Lead 的 J 列公式
    f47 = str(s["lead_ws"].cell(lr["fa"], s["lead_lay"]["cols"]["prior"]).value or "")
    ck("本年 Lead J 列引用上年 Lead",
       f47.startswith("='上年 K.00 Lead Sheet'!$I$"), f47[:48])

    # 10) SUMIFS 范围不吃合计行
    # SUMIFS 的范围止于数据末行，页末那两行合计（原始 / 去重后）必须在范围外，
    # 否则 TB → K.01 → TB 成环，Excel 会弹出循环引用警告。
    ck("TB 范围不含页末合计行", tb["last"] < tb["foot"], (tb["last"], tb["foot"]))

    # 11) 调整汇总表：五列表头 + 审计调整列是 SUMIFS 而不是常数 0
    #
    # 这一整块都在防同一件事：**把 H 列改回常数 0**。改回去不会报错，也不会让
    # K.01 变红，只是复核人填了调整却进不到审定数里 —— 差错完全静默。
    lws, adj = s["lead_ws"], (s.get("lead_adj") or {})
    hdr = adj.get("header")
    heads = [str(lws.cell(hdr, c).value or "") for c in range(2, 7)] if hdr else []
    ck("调整汇总表五列表头齐全",
       heads == ["调整类型", "主体", "调整说明", "科目名称", "调整金额"], heads)

    hcol = s["lead_lay"]["cols"]["adj_aud"]
    d0 = adj.get("d0") or 0
    bad_adj = []
    for row in s["lead_lay"]["rows"].values():
        f = str(lws.cell(row, hcol).value or "")
        # 固定资产的科目名称在 C 列（货币资金在 D 列），条件值必须是 $C{row}
        if not (f.startswith("=SUMIFS($F$%d:$F$" % d0) and ("$C%d)" % row) in f):
            bad_adj.append((row, f[:40]))
    ck("Lead 审计调整列 = SUMIFS 引用调整汇总表", not bad_adj, bad_adj[:3])

    fcol = s["lead_lay"]["cols"]["adj_bs"]
    bad_f = [(r, lws.cell(r, fcol).value) for r in s["lead_lay"]["rows"].values()
             if lws.cell(r, fcol).value not in (0, 0.0)]
    ck("Lead 账表调整列仍为常数 0（未误改）", not bad_f, bad_f[:3])

    bad_zero = []
    for r in range(d0, (adj.get("d1") or 0) + 1):
        if lws.cell(r, 6).value not in (None, ""):
            bad_zero.append((r, lws.cell(r, 6).value))
    ck("调整汇总表数据区为空（空表 → SUMIFS=0）", not bad_zero, bad_zero[:3])

    fails = [i for i in items if not i["ok"]]
    return {"ok": not fails, "items": items, "fails": fails}
