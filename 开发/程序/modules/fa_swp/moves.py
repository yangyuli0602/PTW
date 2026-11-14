# -*- coding: utf-8 -*-
"""固定资产 SWP · 序时账打标层。

把 1601/1602/1603 的每一行，按**凭证借贷对家**归到 K.01 表1 的某个「方式」。

为什么用凭证对家而不是摘要文字：瑞云的摘要是业务话术
（「承接-2024-11-549#，云渲染-雅安模块化机房建设二期（15组）（替换：25组过滤器）」），
同一句话可能对应购置、在建转入或改造；而借贷对家是硬证据——
借 1601 / 贷 1604 就是在建工程转入，一个字都不用读。

**「方式」是这一段最核心的抽象**：它既是清单里的「新增方式 / 处置方式」字段，
也是 K.01 表1 的行标签。表1 的行不再写死，而是由「数据里出现了哪些方式」决定 ——
模板原有的四个方式永远保留（金额 0 也留着），数据里出现新方式就插行。

判定规则（顺序即优先级）：

    1601 借方（净额 > 0）
        本凭证 1604.x 有贷方 → 摘要「承接/改造」= 改造转入，否则 = 在建工程转入
        其余                 → 购置
    1601 贷方（净额 < 0，含以负借方表示的冲回）
        摘要「调整/冲回/转固」 → 转固冲回
        本凭证 1602 有借方    → 处置或报废
        其余                  → 内部重分类
    1602 / 1603：凭证含 1601 → 随资产转销（处置或报废），否则 → 计提

三个必须记住的坑：

1. **1604 是明细科目**（1604.02 / 1604.03），用 ``= '1604'`` 精确匹配会一行都取不到，
   结果全部 1.3 亿的转固被算成「购置」。必须 ``like '1604%'``。
2. **``debit`` / ``credit`` 会出现负数**（红字冲销）。一律用 ``debit - credit`` 的带符号
   净额判方向，不要写 ``if debit:``。
3. **摘要只在前缀上判**。正文里出现「转固」不等于这笔是转固 —— 那 2 笔红字冲回的
   摘要以「调整-」开头，正文才提到转固；只认 ``startswith``。

闭合性由构造保证：每一笔 1601 的借贷都归进某一 (段, 方式)，一分钱不留在外面，
所以 ``Σ(该段各方式) == 净增`` 恒成立。跑完按段自检，不成立直接抛。
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal as D

from . import prep


# ------------------------------------------------------------------ 方式定义

#: 每段的规范顺序（也是表1 插入新行时的排序依据：先增后减）
WAY_ORDER = {
    "fa": ("购置", "在建工程转入", "改造转入", "转固冲回", "处置或报废", "内部重分类"),
    "dep": ("计提", "处置或报废", "内部重分类"),
    "imp": ("计提", "处置或报废", "内部重分类"),
}

#: 模板 K.01 表1 **本来就有**的方式。金额为 0 也必须在表1 里留一行 ——
#: 这是「保留模板原有变动类别」的落点，靠它而不是靠删行来保证版面稳定。
TEMPLATE_WAYS = {
    "fa": ("购置", "在建工程转入", "处置或报废", "内部重分类"),
    "dep": ("计提", "处置或报废", "内部重分类"),
    "imp": ("计提", "处置或报废", "内部重分类"),
}


@dataclass(frozen=True)
class Way:
    sec: str            # fa / dep / imp
    name: str           # 表1 的行标签
    kind: str           # add / reduce
    src: str            # mv（变动清单）/ dp（处置清单）/ fa（FA list）/ const
    field: str          # 取哪一列
    sign: int           # +1 / -1
    list_way: str = ""  # 清单里用的词（空 = 与表1 相同）


#: 方式 → 数据来源。**这是「表1 行 ↔ 清单字段」唯一映射点**。
#: 清单写「内部转出」，表1 写「内部重分类」，就靠 list_way 连起来；
#: 没有这张表，SUMIFS 匹配不到，某行恒 0 而合计看着还是对的。
WAY_DEFS = {
    ("fa", "购置"):       Way("fa", "购置", "add", "mv", "orig", +1),
    ("fa", "在建工程转入"): Way("fa", "在建工程转入", "add", "mv", "orig", +1),
    ("fa", "改造转入"):    Way("fa", "改造转入", "add", "mv", "orig", +1),
    ("fa", "转固冲回"):    Way("fa", "转固冲回", "add", "mv", "orig", +1),
    ("fa", "处置或报废"):  Way("fa", "处置或报废", "reduce", "dp", "orig", -1),
    ("fa", "内部重分类"):  Way("fa", "内部重分类", "reduce", "dp", "orig", -1, "内部转出"),
    ("dep", "计提"):       Way("dep", "计提", "add", "fa", "prov_dep", +1),
    ("dep", "处置或报废"): Way("dep", "处置或报废", "reduce", "dp", "dep", -1),
    ("dep", "内部重分类"): Way("dep", "内部重分类", "reduce", "dp", "dep", -1, "内部转出"),
    ("imp", "计提"):       Way("imp", "计提", "add", "const", "", +1),
    ("imp", "处置或报废"): Way("imp", "处置或报废", "reduce", "dp", "imp", -1),
    ("imp", "内部重分类"): Way("imp", "内部重分类", "reduce", "dp", "imp", -1, "内部转出"),
}

#: 摘要前缀 → 业务方式。只认前缀，不认正文。
BIZ_WORDS = (
    ("改造转入", ("承接", "改造")),
    ("转固冲回", ("调整", "冲回")),
    ("在建工程转入", ("结转", "在建工程", "转固")),
)

#: 表4 的费用归集：科目前缀 → 模板行名。顺序敏感，5602 必须排在 66 之前。
#: 5101 / 5001 是「生产成本」类 —— 制造业把折旧先计入制造费用再进产品成本，
#: 不认这两个前缀，那部分折旧会掉进「其他」，表4 与计提行的差就没法解释。
EXP_BUCKETS = (
    ("制造费用", ("6401", "5101", "5001")),
    ("销售费用", ("6601",)),
    ("管理费用", ("6602", "5602")),
    ("研发费用", ("6605", "6604")),
)
EXP_OTHER = "其他"

CTX_SQL = """
select voucher_key,
       sum(case when acct_code like '1601%' then coalesce(debit, 0) else 0 end) as fa_dr,
       sum(case when acct_code like '1601%' then coalesce(credit, 0) else 0 end) as fa_cr,
       sum(case when acct_code like '1602%' then coalesce(debit, 0) else 0 end) as dep_dr,
       sum(case when acct_code like '1602%' then coalesce(credit, 0) else 0 end) as dep_cr,
       sum(case when acct_code like '1604%' then coalesce(credit, 0) else 0 end) as cip_cr,
       sum(case when acct_code like '1604%' then coalesce(debit, 0) else 0 end) as cip_dr
  from journal
 group by 1
"""

LINE_SQL = """
select company_name, acct_code, acct_name, acct_l1, summary,
       vdate, vno, voucher_key, period,
       coalesce(aux_text, '') as aux_text,
       coalesce(debit, 0) as debit,
       coalesce(credit, 0) as credit,
       src_row
  from journal
 where acct_code like '1601%' or acct_code like '1602%' or acct_code like '1603%'
 order by vdate, voucher_key, src_row
"""


def _d(v) -> D:
    if v is None:
        return D("0")
    if isinstance(v, D):
        return v
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "infinity", "-infinity", "inf", "-inf"):
        return D("0")
    return D(s)


def load_ctx(conn) -> dict[str, dict]:
    """凭证级上下文：一张凭证里 1601/1602/1604 各自的借贷发生额。"""
    cur = conn.execute(CTX_SQL)
    cols = [d[0] for d in cur.description]
    out = {}
    for rec in cur.fetchall():
        r = dict(zip(cols, rec))
        # 必须先把凭证号取出来再建字典：订阅赋值是「先算右边、再算下标」，
        # 写成一行的 r.pop 会把凭证号一起喂给 _d。
        vk = r.pop("voucher_key")
        out[vk] = {k: _d(v) for k, v in r.items()}
    return out


def load_lines(conn) -> list[dict]:
    cur = conn.execute(LINE_SQL)
    cols = [d[0] for d in cur.description]
    out = []
    for rec in cur.fetchall():
        r = dict(zip(cols, rec))
        r["debit"] = _d(r["debit"])
        r["credit"] = _d(r["credit"])
        r["src_row"] = int(r["src_row"])
        out.append(r)
    return out


def line_cat(line: dict) -> str:
    """一行的类别。**委托 `prep.category`**，与 TB 侧共用同一个解析器。

    这里以前是第二份平行实现（只读 aux_text、自己判 1603、自己查兜底表），
    只要两边有一处改了、另一处没改，TB 侧和序时账侧就会给出**不同的类别集合**，
    而 K.01 的期初来自 TB、变动来自序时账 —— 结果是期末对不上，
    却看不出是分类口径分叉造成的。
    """
    return prep.category(line)


def biz_of(summary: str) -> str:
    """摘要前缀 → 业务方式。空串 = 认不出（由调用方按借贷对家兜底）。"""
    s = (summary or "").lstrip()
    for way, prefixes in BIZ_WORDS:
        if s.startswith(prefixes):
            return way
    return ""


def classify(line: dict, ctx: dict) -> tuple[str, str, D]:
    """返回 (段, 方式, 带符号金额)。段为空串表示归不进。"""
    net = line["debit"] - line["credit"]
    if net == 0:
        return "", "", D("0")
    c = ctx.get(line["voucher_key"], {})
    acct = line["acct_code"]
    kind = prep.acct_kind(acct)          # 瑞云是一级科目，海柔是 1601.03 这种明细
    biz = biz_of(line["summary"])

    if kind == "fa":
        if net > 0:
            if c.get("cip_cr", 0) > 0:
                # 借 1601 / 贷 1604：承接类（部件更换、机房改造）与结转类分开
                return "fa", ("改造转入" if biz == "改造转入" else "在建工程转入"), net
            return "fa", "购置", net
        # 贷 1601（含以负借方表示的冲回）
        if biz == "转固冲回":
            return "fa", "转固冲回", net
        if c.get("dep_dr", 0) > 0:
            return "fa", "处置或报废", net
        return "fa", "内部重分类", net

    if kind == "dep":
        # 判据是凭证不是摘要：凭证里有 1601 的，折旧在随资产转销；没有的才是真计提。
        # 瑞云有一批「计提-红冲」行 credit 为负，用摘要判会漏。
        if c.get("fa_dr", 0) > 0 or c.get("fa_cr", 0) > 0:
            return "dep", "处置或报废", -net
        return "dep", "计提", -net

    # 其余能进到这里的只有 1603 减值准备（kind == "imp"）
    if kind != "imp":
        return "", "", D("0")
    if c.get("fa_dr", 0) > 0 or c.get("fa_cr", 0) > 0:
        return "imp", "处置或报废", -net
    return "imp", "计提", -net


def exp_bucket(code: str, name: str) -> str:
    """折旧费用归到模板表4 的哪一行。按科目前缀，不按科目名。"""
    c4 = (code or "")[:4]
    for bucket, prefixes in EXP_BUCKETS:
        if c4 in prefixes:
            return bucket
    return EXP_OTHER


def table4(conn, lines: list[dict], ctx: dict, z20_target: D) -> dict:
    """表4：折旧费用与损益类科目勾稽。

    **范围两层筛**：

    1. 凭证层：只取「含 1602 且不含任何 1601」的凭证 —— 这些是计提凭证，
       带 1601 的是处置（折旧在转销，不是费用）。
    2. 科目层：只认**科目名含「折旧」且不含「使用权」**的费用科目。

    第 2 层不是可选项。制造业客户的月度结转凭证是「一张凭证里折旧、摊销、
    使用权资产折旧、办公费全在一起」，只按凭证筛，表4 会变成当月全部费用
    （海柔实测 2,764 万 vs 计提 478 万）。折旧费也可能挂在下级科目
    （5101.15 制造费用_折旧费），所以前缀不能只认 64/66/5602。

    ``z20_target`` 传 K.01 计提行的合计；两者不相等是正常的（折旧进了存货、
    在建工程就没走损益），差异由调用方披露，不在这里摊平。
    """
    pure = {vk for vk, c in ctx.items()
            if c.get("dep_cr", 0) != 0 and c.get("fa_dr", 0) == 0 and c.get("fa_cr", 0) == 0}

    buckets = {b: D("0") for b, _ in EXP_BUCKETS}
    buckets[EXP_OTHER] = D("0")
    details = {b: [] for b in buckets}
    z20 = z20_target

    if pure:
        marks = ",".join("?" * len(pure))
        rows = conn.execute(
            "select acct_code, acct_name, sum(coalesce(debit,0)) - sum(coalesce(credit,0)) amt "
            "from journal where voucher_key in (%s) and acct_code not like '1602%%' "
            "  and (acct_code like '64%%' or acct_code like '66%%' or acct_code like '5602%%' "
            "       or acct_code like '5101%%' or acct_code like '5001%%') "
            "group by 1, 2 having abs(sum(coalesce(debit,0)) - sum(coalesce(credit,0))) > 0.005 "
            "order by 3 desc" % marks, list(pure)).fetchall()
        for code, name, amt in rows:
            n = name or ""
            if "折旧" not in n or "使用权" in n or "投资性房地产" in n:
                continue        # 摊销费 / 使用权资产折旧 / 办公费，都不是固定资产折旧
            b = exp_bucket(code, name)
            buckets[b] += _d(amt)
            details[b].append({"code": code, "name": name, "amount": str(amt)})

    total = sum(buckets.values(), D("0"))
    return {
        "buckets": buckets, "details": details, "total": total,
        "z20": z20, "diff": total - z20, "vouchers": len(pure),
    }


# ------------------------------------------------------------------ 汇总

def amounts(mvs: dict, sec: str, way: str) -> D:
    """某个 (段, 方式) 的合计。找不到返回 0。"""
    return mvs["totals"].get((sec, way), D("0"))


def by_cat(mvs: dict, sec: str, way: str, cat: str) -> D:
    """某个 (段, 方式, 类别) 的金额。"""
    return mvs["grid"].get((sec, way), {}).get(cat, D("0"))


def build(conn, prep_res: dict) -> dict:
    """跑完整分类，返回 K.01 各 (段, 方式) × 类别的矩阵，并自检闭合。"""
    lines = load_lines(conn)
    ctx = load_ctx(conn)

    grid = {}
    unknown = []
    counts = {}
    cat_seen = set()
    order = prep_res["cat_order"]

    for ln in lines:
        sec, way, amt = classify(ln, ctx)
        if not sec:
            if ln["debit"] - ln["credit"] != 0:
                unknown.append(ln)
            continue
        cat = line_cat(ln)
        if not cat:
            unknown.append(ln)
            continue
        key = (sec, way)
        # 用**本次数据的**类别顺序初始化，不用模块级常量。写死元组时，
        # 第 5 个类别会在下面那行 KeyError —— 或者更糟，如果只改了初始化
        # 没改这里，金额会被算进一个不存在的键。
        grid.setdefault(key, {c: D("0") for c in order})
        if cat not in grid[key]:
            raise ValueError(
                "序时账里出现 TB 没有的类别 %r（凭证明细 %s %s），"
                "先确认是分类错还是 TB 漏了该类" % (cat, ln["vdate"], ln["vno"]))
        grid[key][cat] += amt
        cat_seen.add(cat)
        counts[key] = counts.get(key, 0) + 1

    if unknown:
        sample = [(r["company_name"], r["acct_code"], str(r["debit"]), str(r["credit"]),
                   (r["summary"] or "")[:30]) for r in unknown[:5]]
        raise ValueError("有 %d 行固定资产序时账归不进 K.01 方式：%s" % (len(unknown), sample))

    totals = {k: sum(v.values(), D("0")) for k, v in grid.items()}

    bad = set(totals) - set(WAY_DEFS)
    if bad:
        raise ValueError("出现未定义的方式，需要先在 WAY_DEFS 里补上：%s" % sorted(bad))

    # 按段闭合自检
    t = prep_res["totals"]
    problems = []
    for sec, beg, end in (("fa", "fa_beg", "fa_end"),
                          ("dep", "dep_beg", "dep_end"),
                          ("imp", "imp_beg", "imp_end")):
        got = sum((totals.get((sec, w), D("0")) for w in WAY_ORDER[sec]), D("0"))
        want = t[end] - t[beg]
        if abs(got - want) > D("0.01"):
            problems.append({"kind": "%s_delta" % sec, "calc": str(got), "expect": str(want)})

    # 本期 plan：数据里出现且非 0 的方式（模板原有方式无论有没有数都保留）
    plan = {sec: set() for sec in WAY_ORDER}
    for (sec, way), amt in totals.items():
        if abs(amt) > D("0.005"):
            plan[sec].add(way)

    t4 = table4(conn, lines, ctx, totals.get(("dep", "计提"), D("0")))
    if abs(t4["diff"]) > D("0.01"):
        problems.append({"kind": "table4_diff", "diff": str(t4["diff"]),
                         "total": str(t4["total"]), "z20": str(t4["z20"])})

    return {
        "grid": grid,
        "totals": totals,
        "plan": plan,
        # 序时账里实际命中过的类别集合。与 TB 侧比，只出现在一边就说明
        # 分类口径分叉了 —— 见 checks.check_cat_union。
        "cat_seen": cat_seen,
        "counts": {"%s|%s" % k: v for k, v in counts.items()},
        "table4": t4,
        "problems": problems,
        "lines": lines,
        "ctx": ctx,
    }
