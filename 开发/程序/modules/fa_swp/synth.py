# -*- coding: utf-8 -*-
"""固定资产 SWP · 三张清单的构造（变动清单 / 处置清单 / FA list）。

**这一层造的是「颗粒度」，不是「金额」。**

金额全部来自瑞云真实账套：变动和处置的每一行都是一张真凭证（真实日期、真实凭证号、
真实摘要、真实金额），期初存量的原值/折旧/减值合计严格等于 TB 去重后的期初数。
造出来的只有「这台设备叫什么、放哪个部门、什么时候启用、折旧几年」这些明细属性，
而且名称全部从客户自己的凭证摘要里取，不是编的。

所以三张表能同时满足：
    Σ变动的原值 == K.01 R13 + R14
    Σ处置的原值 == K.01 R15 + R16
    FA list 期初 == K.01 R12，期末 == K.01 R18

分配算法只有一条铁律：**先定合计，再拆明细，最后一项吸收尾差**。
绝不「先生成随机数再四舍五入凑总数」—— 那会差几分钱，CHECK 永远不为 0。
"""

from __future__ import annotations

import datetime as dt
import hashlib
import re
import random
from decimal import Decimal as D

from . import moves, prep

CENTS = D("0.01")

#: 内置的类别元数据（折旧年限 / 残值率 / 部门池 / 存放地点 / 项数 / 编码前缀）。
#:
#: 保留这一份是为了**瑞云改造前后产物逐格一致**，同时给「配置里没写但确实是
#: 常见类别」的情况一个好默认。换客户不需要改这里：把类别写进
#: ``data/fa_swp_cats.json`` 的 ``cats`` 就会覆盖；连配置都没有的类别走
#: ``defaults`` 的通用兜底，并在底稿上标注「元数据未配置」。
_DEFAULT_META = {
    "机器设备": {
        "life_months": (60, 120),
        "resid_rates": (0.0, 0.03, 0.05),
        "dept": ("云渲染运维部", "IDC 运维部", "技术中心", "云游戏业务部"),
        "loc": ("深圳百旺信机房", "贵安电信机房", "雅安机房", "鄂尔多斯 T4 机房", "金华移动机房"),
        "count": 32,
        "code_prefix": "JQ",
    },
    "电子设备": {
        "life_months": (36, 60),
        "resid_rates": (0.0, 0.03, 0.05),
        "dept": ("研发中心", "行政部", "财务部", "信息技术部"),
        "loc": ("深圳总部", "北京分公司", "成都分公司", "杭州分公司"),
        "count": 26,
        "code_prefix": "DZ",
    },
    "其他设备": {
        "life_months": (36, 60),
        "resid_rates": (0.0, 0.03, 0.05),
        "dept": ("行政部", "技术中心"),
        "loc": ("深圳总部", "广州分公司"),
        "count": 22,
        "code_prefix": "QT",
    },
    # 运输设备总额只有 36 万，硬凑 20 项会拆成一堆 1.8 万的「运输设备」，
    # 反而是假数据，所以降到 8 项。
    "运输设备": {
        "life_months": (48, 96),
        "resid_rates": (0.05,),
        "dept": ("行政部", "物流部"),
        "loc": ("深圳总部",),
        "count": 8,
        "code_prefix": "YS",
    },
}


def count_of(fa_beg) -> int:
    """期初原值 → 造几项资产。**期初为 0 就是 0 项。**

    不能给固定项数：它直接决定 FA list 的观感。2.66 亿拆成 3 项 = 8000 万/台，
    一眼就是编的；拆 500 项又让文件大到打不开。按「单项 50 万」估算再夹逼，
    比写死四个数字可解释得多，也不会因为换客户就失真。
    """
    if not fa_beg or float(fa_beg) <= 0:
        return 0
    d = prep.cfg().get("defaults") or {}
    avg = float(d.get("avg_item_value") or 500000)
    lo = int(d.get("count_min") or 3)
    hi = int(d.get("count_max") or 40)
    n = int(round(float(fa_beg) / avg)) if avg else hi
    n = max(lo, min(hi, n))
    n = min(n, max(1, int(float(fa_beg) // 10000)))     # 单项不低于 1 万，防一堆碎设备
    return max(1, n)


def meta(cat: str, fa_beg=None) -> dict:
    """一个类别的造明细参数。配置 > 内置 > 通用兜底，三级。

    ``configured`` 为 False 时表示「这个类别没有任何人配置过」，
    调用方（`k01.fill_table1`）会在表头加批注，Notes 里也写一行 ——
    不能让一份用「5–10 年、行政部」凑出来的明细看起来像客户台账。
    """
    conf = prep.cfg()
    d = conf.get("defaults") or {}
    out = {
        "life_months": tuple(d.get("life_months") or (60, 120)),
        "resid_rates": tuple(d.get("resid_rates") or (0.0, 0.03, 0.05)),
        "dept": tuple(d.get("dept") or ("行政部",)),
        "loc": tuple(d.get("loc") or ("公司所在地",)),
        "count": None,
        "code_prefix": str(d.get("code_prefix") or ""),
        "configured": False,
    }
    for src in (_DEFAULT_META.get(cat), (conf.get("cats") or {}).get(cat)):
        if not src:
            continue
        out["configured"] = True
        for k in ("life_months", "resid_rates", "dept", "loc"):
            if src.get(k):
                out[k] = tuple(src[k])
        if src.get("code_prefix"):
            out["code_prefix"] = str(src["code_prefix"])
        if src.get("count"):
            out["count"] = int(src["count"])
    if out["count"] is None:
        out["count"] = count_of(fa_beg)
    return out


# ------------------------------------------------------------------ 工具

def _rng(*parts) -> random.Random:
    """固定种子：同样的输入每次跑出同样的明细，便于复核与回归。"""
    key = "|".join(str(p) for p in parts).encode("utf-8")
    return random.Random(int(hashlib.md5(key).hexdigest()[:12], 16))


def alloc(total: D, n: int, rng: random.Random, sigma: float = 0.75,
          weights=None) -> list[D]:
    """把 total 拆成 n 份，**合计精确等于 total**。

    前 n-1 份按分取整，最后一份 = total − 已分配，尾差全吃在最后一份上，
    这样 Excel 里 SUM 出来一分不差。

    ``weights`` 是「每份最多能装多少」的参考权重（一般传原值）。不给权重就是
    纯对数正态各分一份 —— 那会出现「原值 1 万的小设备分了 5 万折旧」，
    净值直接变成负数，K.01 表2 与表1 一起崩。折旧、减值一律带权重。
    """
    if n <= 0:
        return []
    if n == 1:
        return [total]
    w = [rng.lognormvariate(0.0, sigma) for _ in range(n)]
    if weights:
        w = [a * float(b) for a, b in zip(w, weights)]
    s = sum(w) or 1.0
    out, acc = [], D("0")
    for i in range(n - 1):
        v = (total * D(str(w[i] / s))).quantize(CENTS)
        out.append(v)
        acc += v
    out.append(total - acc)
    return out


def _month_add(d: dt.date, months: int) -> dt.date:
    y, m = divmod(d.month - 1 + months, 12)
    return dt.date(d.year + y, m + 1, 1)


def _months_between(d0: dt.date, d1: dt.date) -> int:
    return max(0, (d1.year - d0.year) * 12 + (d1.month - d0.month))


def dep_of(orig: D, resid: float, life: int, start: dt.date, asof: dt.date) -> D:
    """直线法到 asof 的累计折旧，封顶到 (1−残值率)×原值。"""
    if orig <= 0 or life <= 0:
        return D("0")
    used = _months_between(start, asof)
    if used <= 0:
        return D("0")
    v = orig * D(str(1 - resid)) * D(used) / D(life)
    cap = orig * D(str(1 - resid))
    return min(v, cap).quantize(CENTS)


def sample_names(lines: list[dict], cats: tuple, per_cat: int, seed: str) -> dict:
    """从客户真实凭证摘要里取样，做成「XX机房-YY台设备」形式的名称池。

    不自己编设备名 —— 编出来的名字一眼假，而且和客户台账对不上号。
    """
    pools = {c: [] for c in cats}
    for ln in lines:
        cat = moves.line_cat(ln)
        if cat not in pools:
            continue
        s = (ln["summary"] or "").strip()
        for sep in ("承接-", "结转---", "结转-", "调整-"):
            if s.startswith(sep):
                s = s[len(sep):]
                break
        if "#" in s:
            s = s.split("#", 1)[1].lstrip("，,、 ")
        s = s.split("（")[0].strip(" -，,")
        s = re.sub(r"^(等[，,、]?|、)+", "", s).strip()
        if 4 <= len(s) <= 40:
            pools[cat].append(s)
    out = {}
    for c in cats:
        uniq = sorted(set(pools[c]))
        rng = _rng("name", seed, c)
        if len(uniq) >= per_cat:
            out[c] = rng.sample(uniq, per_cat)
        elif uniq:
            out[c] = [uniq[i % len(uniq)] for i in range(per_cat)]
        else:
            out[c] = ["%s-第%02d批" % (c, i + 1) for i in range(per_cat)]
    return out


# ------------------------------------------------------- 变动 / 处置（真实凭证）

def aggregate_vouchers(lines: list[dict], ctx: dict) -> tuple[list[dict], list[dict]]:
    """把序时账按「凭证 × 类别 × 行项目」聚合成清单行。

    一行 = 一笔凭证在一个类别上的一个业务动作。这是审计清单的自然粒度：
    凭证号就是原始单据号，能直接回查。
    """
    acc = {}
    for ln in lines:
        sec, way, amt = moves.classify(ln, ctx)
        if not sec or amt == 0:
            continue
        # 只有原值段进口径清单；折旧/减值行只用来定 K.01 的数，不进口径清单
        if sec != "fa":
            continue
        cat = moves.line_cat(ln)
        if not cat:
            continue
        k = (way, cat, ln["voucher_key"])
        if k not in acc:
            acc[k] = {
                "row": way, "cat": cat, "vkey": ln["voucher_key"],
                "date": ln["vdate"], "no": str(ln["vno"]),
                "company": ln["company_name"], "summary": ln["summary"],
                "amount": D("0"),
            }
        acc[k]["amount"] += amt

    mv, dp = [], []
    for (way, cat, _vk), r in sorted(acc.items(), key=lambda kv: (kv[1]["date"], kv[1]["vkey"], kv[0][0])):
        r = dict(r)
        r["amount"] = r["amount"].quantize(CENTS)
        if r["amount"] == 0:
            continue
        # 清单里的词与表1 的行名可以不同（内部重分类 ↔ 内部转出），
        # 唯一翻译点就是 Way.list_way，别在别处再写一遍字符串。
        w = moves.WAY_DEFS[("fa", way)]
        r["way"] = w.list_way or w.name
        if w.kind == "add":
            mv.append(r)
        else:
            r["amount"] = -r["amount"]          # 减少统一记正数
            dp.append(r)
    return mv, dp


# ------------------------------------------------------------------ FA list

# ---------------------------------------------------------- 客户台账（真卡片）

#: 台账字段顺序：load_cards 的 select 与 _cards_rows 的下标必须一致。
CARD_COLS = ("src_row, company_name, category_name, asset_code, asset_name, dept_name, "
             "location, start_date, life_months, residual_rate, orig_cost, acc_dep, "
             "impair, status")


def load_cards(conn) -> dict | None:
    """从库里读客户资产台账：``fa_beg`` 期初、``fa`` 期末。

    两张表都空就返回 None，调用方回落到「按序时账造明细」那条老路。

    配对键用 ``src_row``（原表行号），不用卡片编码 —— 客户台账里同一张卡会被
    拆成多行、编码重复，拿编码配对会把两张卡缝成一张，原值直接翻倍。
    """
    nb = conn.execute("select count(*) from fa_beg").fetchone()[0]
    ne = conn.execute("select count(*) from fa").fetchone()[0]
    if not nb and not ne:
        return None
    beg = {r[0]: r for r in conn.execute("select " + CARD_COLS + " from fa_beg").fetchall()}
    end = {r[0]: r for r in conn.execute("select " + CARD_COLS + " from fa").fetchall()}
    return {"beg": beg, "end": end}


def _cards_rows(prep_res: dict, cards: dict, balance_date: dt.date) -> list[dict]:
    """台账两张表按行号配对，拼成 FA list 行。

    变动额一律用**净额**：客户台账一行只给期初、期末两个数，中间怎么变看不到。
    净额法的代价是「同一行既增又减」会被抵掉，但合计一分不差，且这正是台账的
    口径 —— 明细与台账要对得上，就不能在这层自己拆发生额。
    """
    cats = prep_res["cats"]
    out = []
    for row_no in sorted(set(cards["beg"]) | set(cards["end"])):
        b = cards["beg"].get(row_no)
        e = cards["end"].get(row_no)
        ref = e or b
        cat = (ref[2] or "").strip()
        if cat not in cats:
            raise ValueError("台账里出现 TB 没有的类别 %r（第 %s 行，卡片 %s）"
                             % (cat, row_no, ref[3]))
        fa_beg = D(b[10] or 0) if b else D("0")
        dep_beg = D(b[11] or 0) if b else D("0")
        imp_beg = D(b[12] or 0) if b else D("0")
        fa_end = D(e[10] or 0) if e else D("0")
        dep_end = D(e[11] or 0) if e else D("0")
        imp_end = D(e[12] or 0) if e else D("0")
        out.append({
            "no": 0, "code": (ref[3] or "").strip(), "name": (ref[4] or "").strip(),
            "cat": cat, "dept": (ref[5] or "").strip(), "loc": (ref[6] or "").strip(),
            "start": ref[7] or balance_date, "life": int(ref[8] or 0),
            "resid": float(ref[9] or 0),
            "fa_beg": fa_beg.quantize(CENTS), "dep_beg": dep_beg.quantize(CENTS),
            "imp_beg": imp_beg.quantize(CENTS),
            "nbv_beg": (fa_beg - dep_beg - imp_beg).quantize(CENTS),
            "add": max(D("0"), fa_end - fa_beg).quantize(CENTS),
            "reduce": max(D("0"), fa_beg - fa_end).quantize(CENTS),
            "prov_dep": D("0"), "reduce_dep": D("0"),
            "fa_end": fa_end.quantize(CENTS), "dep_end": dep_end.quantize(CENTS),
            "imp_end": imp_end.quantize(CENTS),
            "nbv_end": (fa_end - dep_end - imp_end).quantize(CENTS),
        })
    return out


def _from_cards(prep_res, mvs, balance_date, cards, seed) -> dict:
    """走客户台账：明细一行不少地照搬，本期计提/转出按 K.01 的合计倒摊。

    台账只给「期末累计折旧」，没给「本期计提」和「本期转出折旧」两列，所以：
        净增 = 期末折旧 − 期初折旧          （台账口径，合计一定等于折旧净增）
        转出折旧 = 按 K.01 该类折旧转出额，摊给本期有原值减少的行
        本期计提 = 净增 + 转出折旧
    三者仍满足 期末 = 期初 + 计提 − 转出，逐行、逐类、合计都对得上。
    """
    fa_rows = _cards_rows(prep_res, cards, balance_date)
    mv_rows, dp_rows = aggregate_vouchers(mvs["lines"], mvs["ctx"])
    order = list(prep_res["cat_order"])
    cats = prep_res["cats"]

    # ---- 台账口径对齐：底稿以**账**为准
    # 客户台账与科目余额表差几千块是常事（卡片科目改过、期初录入早于调账），
    # 但 K.01 的年初/年末必须等于总账，否则整个复核链条第一格就断。
    # 差额定点补到该类金额最大的卡片上，并留痕 —— 不重分类别、不改总额。
    adj = {}
    for cat in order:
        rows = [r for r in fa_rows if r["cat"] == cat]
        if not rows:
            continue
        for f, key in (("fa_beg", "fa_beg"), ("dep_beg", "dep_beg"),
                       ("fa_end", "fa_end"), ("dep_end", "dep_end")):
            want = D(cats[cat][key]).quantize(CENTS)
            got = sum((r[f] for r in rows), D("0")).quantize(CENTS)
            gap = (want - got).quantize(CENTS)
            if gap == 0:
                continue
            big = max(rows, key=lambda r: abs(r[f]))
            if big[f] + gap < 0:
                raise ValueError("%s 的台账%s比账少 %.2f，最大的卡片也装不下，"
                                 "先确认台账期间是否正确" % (cat, f, -float(gap)))
            big[f] = (big[f] + gap).quantize(CENTS)
            adj.setdefault(cat, {})[f] = str(gap)
        for r in rows:
            r["add"] = max(D("0"), r["fa_end"] - r["fa_beg"]).quantize(CENTS)
            r["reduce"] = max(D("0"), r["fa_beg"] - r["fa_end"]).quantize(CENTS)
            r["nbv_beg"] = (r["fa_beg"] - r["dep_beg"] - r["imp_beg"]).quantize(CENTS)
            r["nbv_end"] = (r["fa_end"] - r["dep_end"] - r["imp_end"]).quantize(CENTS)


    # 本期转出折旧：逐类锁到 K.01 的折旧「处置或报废」行
    for cat in order:
        rows = [r for r in fa_rows if r["cat"] == cat and r["reduce"] > 0]
        target = (-moves.by_cat(mvs, "dep", "处置或报废", cat)).quantize(CENTS)
        if not rows:
            if target:
                raise ValueError("%s 账上有折旧转出 %.2f，但台账里没有原值减少的卡片"
                                 % (cat, float(target)))
            continue
        for r, v in zip(rows, alloc(target, len(rows), _rng("cd", seed, cat), 0.25,
                                    [r["reduce"] for r in rows])):
            # 单行转出折旧封顶到该行**期初**折旧：被处置的卡片期末折旧是 0，
            # 拿 dep_end 当上限会把额度全压没，然后整类报「摆不平」。
            r["reduce_dep"] = min(v, max(D("0"), r["dep_beg"])).quantize(CENTS)
        gap = (target - sum((r["reduce_dep"] for r in rows), D("0"))).quantize(CENTS)
        if gap:
            for r in sorted(rows, key=lambda x: -(x["dep_beg"] - x["reduce_dep"])):
                if gap <= 0:
                    break
                room = (r["dep_beg"] - r["reduce_dep"]).quantize(CENTS)
                if room <= 0:
                    continue
                take = min(room, gap)
                r["reduce_dep"] += take
                gap -= take
        if gap:
            raise ValueError("%s 台账转出折旧摆不平，仍差 %.2f" % (cat, float(gap)))

    for i, r in enumerate(fa_rows, 1):
        r["no"] = i
        r["prov_dep"] = (r["dep_end"] - r["dep_beg"] + r["reduce_dep"]).quantize(CENTS)

    # 处置清单的折旧/减值：与老路一致，按类别锁到 K.01
    dp_out = []
    for i, r in enumerate(dp_rows, 1):
        dp_out.append({
            "no": i, "cat": r["cat"], "date": r["date"], "vno": r["no"], "vkey": r["vkey"],
            "code": "", "name": r["summary"], "way": r["way"], "orig": r["amount"],
            "dep": D("0"), "imp": D("0"), "company": r["company"],
        })
    for cat in order:
        rows = [r for r in dp_out if r["cat"] == cat]
        if not rows:
            continue
        dep_total = (-moves.by_cat(mvs, "dep", "处置或报废", cat)).quantize(CENTS)
        for r, v in zip(rows, alloc(dep_total, len(rows), _rng("dpd", seed, cat), 0.3,
                                    [x["orig"] for x in rows])):
            r["dep"] = v

    mv_out = []
    for i, r in enumerate(mv_rows, 1):
        mv_out.append({
            "no": i, "cat": r["cat"], "date": r["date"], "vno": r["no"], "vkey": r["vkey"],
            "code": "", "name": r["summary"], "dept": "", "way": r["way"],
            "orig": r["amount"], "dep_in": D("0"), "company": r["company"],
        })

    return {"fa": fa_rows, "moves": mv_out, "disposals": dp_out,
            "trace": {"source": "cards", "opening_n": len(fa_rows),
                      "mv_n": len(mv_out), "dp_n": len(dp_out),
                      "cards_adj": adj, "gap_prov": "0", "gap_reduce": "0"}}


def build(prep_res: dict, mvs: dict, balance_date: dt.date, seed: str = "ruiyun",
          cards: dict | None = None) -> dict:
    """组装三张清单。返回的行都是 dict，键名固定，供 k01/cli 直接写表。

    ``cards`` 有值（客户台账已导入 fa/fa_beg）就照搬台账明细，没有才走
    「序时账反推明细」——后者只在客户拿不出台账时兜底，明细是造出来的。
    """
    cats = prep_res["cats"]
    lines = mvs["lines"]
    ctx = mvs["ctx"]
    # 类别顺序与元数据都由数据决定。原来这里用模块级固定元组 + 四张硬编码表，
    # 出现第 5 个类别会直接 KeyError —— 换客户就没法跑。
    order = list(prep_res["cat_order"])
    M = {c: meta(c, cats[c]["fa_beg"] if c in cats else 0) for c in order}

    if cards:
        return _from_cards(prep_res, mvs, balance_date, cards, seed)

    mv_rows, dp_rows = aggregate_vouchers(lines, ctx)

    # ---- 期初存量：先按类生成资产卡，原值/折旧/减值合计锁死到 TB 期初
    names = sample_names(lines, order, max((M[c]["count"] for c in order), default=1), seed)
    opening = []
    for cat in order:
        m = cats[cat]
        n = M[cat]["count"]
        rng = _rng("open", seed, cat)
        origs = alloc(m["fa_beg"], n, rng)
        # 处置部分的折旧单独切出来：Σ处置折旧 必须等于 K.01 的本期转出额，
        # 否则 FA list 的期末折旧凑不上 TB。
        disp_orig_cat = sum((r["amount"] for r in dp_rows if r["cat"] == cat), D("0"))
        for i, orig in enumerate(origs):
            life = rng.randint(*M[cat]["life_months"])
            resid = rng.choice(M[cat]["resid_rates"])
            start = _month_add(balance_date, -rng.randint(6, 84))
            opening.append({
                "cat": cat, "orig": orig, "life": life, "resid": resid,
                "start": start, "dept": rng.choice(M[cat]["dept"]),
                "loc": rng.choice(M[cat]["loc"]),
                "name": names[cat][i % len(names[cat])],
                "code": "ZCLB%s%04d" % (M[cat]["code_prefix"], i + 1),
                "disposed": D("0"), "dep_disp": D("0"),
            })

    # ---- 处置摊销：挑几项整笔处置，不是把处置额摊到每个资产上
    #
    # 这里必须「挑整项」：如果按原值比例摊给每一项，88 个资产全都会被标成
    # 「部分处置」，于是「仍然在用的资产」几乎为空，期末折旧/减值会瞬间塌到 0。
    # 真实台账也是这样——处置的是具体某几台设备，不是每台都卖一部分。
    for cat in order:
        d_total = sum((r["amount"] for r in dp_rows if r["cat"] == cat), D("0"))
        pool = [a for a in opening if a["cat"] == cat]
        if d_total <= 0 or not pool:
            continue
        # 注意别用 order 当局部变量名 —— 它是外层的「类别顺序」，
        # 被覆盖后下一次 `for cat in order` 拿到的是资产 dict，报 unhashable。
        picks = list(pool)
        _rng("disp", seed, cat).shuffle(picks)
        acc = D("0")
        for a in picks:
            if acc >= d_total:
                break
            take = min(a["orig"], (d_total - acc).quantize(CENTS))
            if take <= 0:
                break
            a["disposed"] = take
            acc += take
        if acc < d_total - CENTS:
            raise ValueError("%s 的处置额 %.2f 超过期初资产总额，无法分配" % (cat, float(d_total)))

    # ---- 期初折旧 / 减值：处置部分与留存部分分别锁死合计
    #
    # 四个字段必须分开存：dep_beg 是期初，dep_end 是期末（= 期初 + 本期计提）。
    # 早期版本只用一个 "dep"，后面算期末计提时写回去把期初值覆盖了，
    # 结果 FA list 的「期初累计折旧」显示的是期末数，K.01 表2 直接对不上。
    for cat in order:
        m = cats[cat]
        pool = [a for a in opening if a["cat"] == cat]
        disp = [a for a in pool if a["disposed"] > 0]
        keep = [a for a in pool if a["disposed"] <= 0]

        # 处置折旧目标**只认 K.01 的 R21**，不要按原值比例兜底：
        # 电子设备有一笔 2,695.58 的内部转出，账上没有任何折旧转出（R21=0），
        # 按比例兜底会给它硬摊 1,789.88 的折旧，计提行就多出这个数，K.01 立刻差 1,789。
        r21 = moves.by_cat(mvs, "dep", "处置或报废", cat)
        target_disp_dep = (-r21).quantize(CENTS) if r21 else D("0")
        target_keep_dep = (m["dep_beg"] - target_disp_dep).quantize(CENTS)

        for a in pool:
            a["dep_beg"] = D("0")
            a["imp_beg"] = D("0")
        if disp:
            for a, v in zip(disp, alloc(target_disp_dep, len(disp), _rng("dd", seed, cat), 0.25,
                                        [a["disposed"] for a in disp])):
                a["dep_beg"] = v
        if keep:
            for a, v in zip(keep, alloc(target_keep_dep, len(keep), _rng("dk", seed, cat), 0.25,
                                        [a["orig"] for a in keep])):
                a["dep_beg"] = v

        # 减值：被处置的资产身上要有减值（随资产一起转销），其余留给还在用的。
        # 全堆在留存的资产上，处置清单的减值列就会全是 0，K.01 R27 对不上。
        r27 = moves.by_cat(mvs, "imp", "处置或报废", cat)
        disp_imp = (-r27).quantize(CENTS) if r27 else D("0")
        if m["imp_beg"] > 0:
            if disp and disp_imp > 0:
                for a, v in zip(disp, alloc(disp_imp, len(disp), _rng("dpi", seed, cat), 0.25,
                                            [a["orig"] for a in disp])):
                    a["imp_beg"] = v
            rest = (m["imp_beg"] - sum((a["imp_beg"] for a in disp), D("0"))).quantize(CENTS)
            if keep and rest > 0:
                # 减值只能落在「还有账面空间」的资产上：原值 − 已提折旧 > 0
                for a, v in zip(keep, alloc(rest, len(keep), _rng("imp", seed, cat), 0.25,
                                            [max(D("0"), a["orig"] - a["dep_beg"]) for a in keep])):
                    a["imp_beg"] = v
        for a in pool:
            a["dep_end"] = a["dep_beg"]
            a["imp_end"] = a["imp_beg"]

    # ---- 本期新增：直接取真实凭证金额
    new_assets = []
    for r in mv_rows:
        new_assets.append({
            "cat": r["cat"], "orig": r["amount"], "way": r["way"], "date": r["date"],
            "vkey": r["vkey"], "no": r["no"], "name": r["summary"], "company": r["company"],
            "dept": "", "loc": "", "life": 0, "resid": 0.0, "start": r["date"],
            "code": "", "dep": D("0"), "imp": D("0"), "disposed": D("0"),
        })
    for i, a in enumerate(new_assets, 1):
        rng = _rng("new", seed, a["vkey"], a["cat"])
        a["life"] = rng.randint(*M[a["cat"]]["life_months"])
        a["resid"] = rng.choice(M[a["cat"]]["resid_rates"])
        a["dept"] = rng.choice(M[a["cat"]]["dept"])
        a["loc"] = rng.choice(M[a["cat"]]["loc"])
        a["code"] = "ZCLB%s%04d" % (M[a["cat"]]["code_prefix"], 9000 + i)
        a["dep_end"] = dep_of(a["orig"], a["resid"], a["life"], a["start"], balance_date)
        a["imp_end"] = D("0")

    # ---- 期末折旧：K.01 R20 的计提额扣掉新增已提部分，余额按期末原值比例摊给存量
    total_dep_new = sum((a["dep_end"] for a in new_assets), D("0"))
    total_r20 = moves.amounts(mvs, "dep", "计提")
    keep_budget = total_r20 - total_dep_new
    if keep_budget < 0:
        # 新增资产按理论折旧提得比全年计提还多，说明账上的折旧提得偏少；
        # 这时按比例压缩新增折旧，保持合计与 TB 一致，并留痕。
        ratio = (total_r20 / total_dep_new) if total_dep_new else D("0")
        for a in new_assets:
            a["dep_end"] = (a["dep_end"] * ratio).quantize(CENTS)
        total_dep_new = sum((a["dep_end"] for a in new_assets), D("0"))
        keep_budget = total_r20 - total_dep_new

    for cat in order:
        keep = [a for a in opening if a["cat"] == cat and a["disposed"] <= 0]
        if not keep:
            continue
        base = sum((a["orig"] for a in keep), D("0"))
        cat_budget = (keep_budget * base / sum((a["orig"] for a in opening
                                                if a["disposed"] <= 0), D("0"))).quantize(CENTS) if base else D("0")
        for a, v in zip(keep, alloc(cat_budget, len(keep), _rng("kd", seed, cat), 0.3)):
            a["dep_end"] = a["dep_end"] + v
            cap = a["orig"] * D(str(1 - a["resid"]))
            a["dep_end"] = min(a["dep_end"], cap).quantize(CENTS)

    # 按压封顶造成的尾差，补到最大的那项（只补还有空间的）
    for cat in order:
        keep = [a for a in opening if a["cat"] == cat and a["disposed"] <= 0]
        if not keep:
            continue
        want = (cats[cat]["dep_end"] - sum((a["dep_end"] for a in new_assets if a["cat"] == cat), D("0"))).quantize(CENTS)
        gap = want - sum((a["dep_end"] for a in keep), D("0"))
        if gap:
            keep.sort(key=lambda x: x["orig"], reverse=True)
            for a in keep:
                cap = a["orig"] * D(str(1 - a["resid"]))
                room = cap - a["dep_end"]
                if gap > 0 and room > 0:
                    take = min(room, gap)
                    a["dep_end"] += take
                    gap -= take
                elif gap < 0:
                    take = max(-a["dep_end"], gap)
                    a["dep_end"] += take
                    gap -= take
                if gap == 0:
                    break
        if gap and gap > 0:
            # keep 装不下它那份期末折旧：通常是本期有「内部重分类转出」——
            # 资产整体划到别的类别去了，留下的原值变小、但要背的折旧没跟着走。
            # 新增资产本来就是合成的，把差额挪给它：类别合计不变、计提合计也不变
            # （want = 类别目标 − 新增折旧，补完以后 keep + 新增 仍等于类别目标）。
            pool_new = [x for x in new_assets if x["cat"] == cat]
            if pool_new:
                pool_new.sort(key=lambda x: x["orig"], reverse=True)
                pool_new[0]["dep_end"] = (pool_new[0]["dep_end"] + gap).quantize(CENTS)
                gap = D("0")
        if gap:
            # 该类一笔新增都没有，只能报出来：账面折旧率与模拟的年限/残值不一致。
            raise ValueError("%s 期末折旧无法分配，仍差 %.2f（keep=%d 项、原值 %s、容量 %s）"
                             % (cat, float(gap), len(keep),
                                sum((x["orig"] for x in keep), D("0")),
                                sum((x["orig"] * D(str(1 - x["resid"])) for x in keep), D("0"))))

    # ---- 折旧收口：期初折旧不能超过「原值 − 已分减值」
    # 按原值加权分配仍可能把某台小设备的折旧分到超过它自己的原值
    #（权重是概率性的，不是硬约束）。收口一次，差额挪给还有空间的资产。
    def _fit_dep(pool, field, imp_field, target, label):
        for a in pool:
            a[field] = min(a[field], max(D("0"), a["orig"] - a[imp_field]))
        gap = (target - sum((a[field] for a in pool), D("0"))).quantize(CENTS)
        if gap > 0:
            for a in sorted(pool, key=lambda x: -(x["orig"] - x[imp_field] - x[field])):
                if gap <= 0:
                    break
                room = (a["orig"] - a[imp_field] - a[field]).quantize(CENTS)
                if room <= 0:
                    continue
                take = min(room, gap)
                a[field] += take
                gap -= take
        if gap != 0:
            raise ValueError("%s %s 收口后仍差 %.2f" % (label, field, float(gap)))

    # ---- 减值收口：不能让「原值 − 折旧 − 减值」出现负数
    #
    # 折旧封顶在 (1−残值率)×原值，残值率 0 的资产折旧提满就等于原值，
    # 再叠一笔减值，净值就成了负数。真实台账不会这样（减值后按新账面价值重算折旧）。
    # 这里按剩余空间重新收口，差额挪给还有空间的资产；挪不动就报错，不摊平。
    def _fit_imp(pool, field_imp, field_dep, target, label):
        for a in pool:
            room = max(D("0"), a["orig"] - a[field_dep])
            a[field_imp] = min(a[field_imp], room)
        gap = (target - sum((a[field_imp] for a in pool), D("0"))).quantize(CENTS)
        if gap > 0:
            for a in sorted(pool, key=lambda x: -(x["orig"] - x[field_dep] - x[field_imp])):
                if gap <= 0:
                    break
                room = (a["orig"] - a[field_dep] - a[field_imp]).quantize(CENTS)
                if room <= 0:
                    continue
                take = min(room, gap)
                a[field_imp] += take
                gap -= take
        if gap != 0:
            raise ValueError("%s %s 无法在净值非负的前提下摆平，仍差 %.2f" % (label, field_imp, float(gap)))

    for cat in order:
        pool = [a for a in opening if a["cat"] == cat]
        if not pool:
            continue
        # 期初：空间按 dep_beg 算
        _fit_imp(pool, "imp_beg", "dep_beg", cats[cat]["imp_beg"], cat)
        # 期末：未处置的才有期末减值；空间按 dep_end 算
        keep = [a for a in pool if a["disposed"] <= 0]
        if keep:
            _fit_imp(keep, "imp_end", "dep_end", cats[cat]["imp_end"], cat)
        for a in pool:
            if a["disposed"] > 0:
                a["imp_end"] = D("0")

    for cat in order:
        pool = [a for a in opening if a["cat"] == cat]
        if pool:
            _fit_dep(pool, "dep_beg", "imp_beg",
                     sum((a["dep_beg"] for a in pool), D("0")), cat)

    # ---- 组装 FA list 行
    fa_rows = []
    idx = 0
    for a in opening:
        idx += 1
        disp = a["disposed"]
        end_orig = a["orig"] - disp
        # 本期计提 / 本期减少折旧要分列：处置资产的折旧是「转出」不是「计提」，
        # 混在一列里 K.01 的计提行和处置行就分不开了。
        if disp > 0:
            end_dep, end_imp = D("0"), D("0")
            prov_dep, reduce_dep = D("0"), a["dep_beg"]
        else:
            end_dep, end_imp = a["dep_end"], a["imp_end"]
            prov_dep, reduce_dep = (a["dep_end"] - a["dep_beg"]).quantize(CENTS), D("0")
        fa_rows.append({
            "no": idx, "code": a["code"], "name": a["name"], "cat": a["cat"],
            "dept": a["dept"], "loc": a["loc"], "start": a["start"],
            "life": a["life"], "resid": a["resid"],
            "fa_beg": a["orig"], "dep_beg": a["dep_beg"], "imp_beg": a["imp_beg"],
            "nbv_beg": a["orig"] - a["dep_beg"] - a["imp_beg"],
            "add": D("0"), "reduce": disp,
            "prov_dep": prov_dep, "reduce_dep": reduce_dep,
            "fa_end": end_orig, "dep_end": end_dep, "imp_end": end_imp,
            "nbv_end": end_orig - end_dep - end_imp,
        })
    for a in new_assets:
        idx += 1
        fa_rows.append({
            "no": idx, "code": a["code"], "name": a["name"], "cat": a["cat"],
            "dept": a["dept"], "loc": a["loc"], "start": a["start"],
            "life": a["life"], "resid": a["resid"],
            "fa_beg": D("0"), "dep_beg": D("0"), "imp_beg": D("0"), "nbv_beg": D("0"),
            "add": a["orig"], "reduce": D("0"),
            "prov_dep": a["dep_end"], "reduce_dep": D("0"),
            "fa_end": a["orig"], "dep_end": a["dep_end"], "imp_end": D("0"),
            "nbv_end": a["orig"] - a["dep_end"],
        })

    # ---- 清单行（含编码/部门，供 K.02 抽样用）
    mv_out = []
    for i, r in enumerate(mv_rows, 1):
        mv_out.append({
            "no": i, "cat": r["cat"], "date": r["date"], "vno": r["no"], "vkey": r["vkey"],
            "code": "ZCLB%s%04d" % (M[r["cat"]]["code_prefix"], 9000 + i),
            "name": r["summary"], "dept": (M[r["cat"]]["dept"] or ("",))[0], "way": r["way"],
            "orig": r["amount"], "dep_in": D("0"), "company": r["company"],
        })

    dp_out = []
    for i, r in enumerate(dp_rows, 1):
        orig = r["amount"]
        # 处置资产的折旧/减值按该类平均转出率取，保证合计等于 K.01
        dp_out.append({
            "no": i, "cat": r["cat"], "date": r["date"], "vno": r["no"], "vkey": r["vkey"],
            "code": "ZCLB%s%04d" % (M[r["cat"]]["code_prefix"], 1000 + i),
            "name": r["summary"], "way": r["way"], "orig": orig,
            "dep": D("0"), "imp": D("0"), "company": r["company"],
        })

    # 处置清单的折旧/减值：按类别合计锁死到 K.01 R21 / R27
    for cat in order:
        rows = [r for r in dp_out if r["cat"] == cat]
        if not rows:
            continue
        dep_total = -moves.by_cat(mvs, "dep", "处置或报废", cat)
        imp_total = -moves.by_cat(mvs, "imp", "处置或报废", cat)
        for r, v in zip(rows, alloc(dep_total, len(rows), _rng("dpd", seed, cat), 0.3,
                                    [x["orig"] for x in rows])):
            r["dep"] = v
        if imp_total:
            for r, v in zip(rows, alloc(imp_total, len(rows), _rng("dpi", seed, cat), 0.3,
                                        [x["orig"] for x in rows])):
                r["imp"] = v
        # 单行不能出现「折旧 + 减值 > 处置原值」；被压掉的零头补到还有空间的行上，
        # 否则合计会比 K.01 的折旧转出少几分钱，CHCHK 就不为 0 了。
        for r in rows:
            cap = max(D("0"), r["orig"] - r["imp"])
            r["dep"] = min(r["dep"], cap)
        gap = (dep_total - sum((r["dep"] for r in rows), D("0"))).quantize(CENTS)
        if gap > 0:
            for r in sorted(rows, key=lambda x: -(x["orig"] - x["imp"] - x["dep"])):
                if gap <= 0:
                    break
                room = (r["orig"] - r["imp"] - r["dep"]).quantize(CENTS)
                if room <= 0:
                    continue
                take = min(room, gap)
                r["dep"] += take
                gap -= take
        if gap != 0:
            raise ValueError("%s 处置清单折旧无法摆平，仍差 %.2f" % (cat, float(gap)))

    res = {"fa": fa_rows, "moves": mv_out, "disposals": dp_out,
           "trace": {"opening_n": len(opening), "new_n": len(new_assets),
                     "mv_n": len(mv_out), "dp_n": len(dp_out)}}
    # 逐项 quantize 会攒出几分钱，统一在最大的一行上吃掉（超过 1 元就报错，不摊平）
    res["trace"]["gap_prov"] = str(close_gap(res["fa"], "prov_dep", moves.amounts(mvs, "dep", "计提")))
    res["trace"]["gap_reduce"] = str(close_gap(res["fa"], "reduce_dep", -moves.amounts(mvs, "dep", "处置或报废")))
    return res


def close_gap(rows: list[dict], field: str, target: D) -> D:
    """把某一列凑到 target。逐项 quantize 会累计出几分钱，统一在这里吃掉。

    只在**最大的一行**上调整，且只调几分钱 —— 如果缺口很大说明上游算错了，
    那不是尾差，应该报错而不是悄悄摊平。
    """
    if not rows:
        return D("0")
    gap = (target - sum((r[field] for r in rows), D("0"))).quantize(CENTS)
    if gap == 0:
        return gap
    if abs(gap) > D("1"):
        raise ValueError("%s 合计与目标差 %.2f，超出尾差范围，不允许摊平" % (field, float(gap)))
    big = max(rows, key=lambda r: abs(r[field]))
    big[field] = (big[field] + gap).quantize(CENTS)
    return gap
