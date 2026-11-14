# -*- coding: utf-8 -*-
"""固定资产 SWP · 数据准备层。

**这一层只做一件事：把 TB 和序时账整理成一张「类别 × 三科目 × 期初期末」的目标矩阵，
并证明它是对的。** 后面的 K.01、三张清单、Lead 全部以这张矩阵为唯一基准。

为什么需要这一层（三个坑，都是瑞云实测踩出来的）：

1. **TB 重复导入**。瑞云 TB 里 1601/1602/1603 被整块复制了一遍：母公司
   ``src_row`` 1225-1228 与 1230-1233 内容完全相同，期初期末精确翻倍。
   不去重的话，「与 TB 核对一致」核对的其实是一个翻倍的数，底稿看着平、实际全错。
   去重键必须含 ``aux2``（辅助核算编码），否则「机器设备」和「电子设备」
   金额相同时会被当成重复误删。

2. **期初不可信**。瑞影科技是本年新设主体，TB 期初写了 50,306.76、期末也是
   50,306.76，但序时账本年明明有 53,002.34 借方。按
   ``期初 = 期末 − 序时账净额`` 反推，真实期初是 0。
   这类主体不修正，K.01 的 CHECK 永远差那 5 万。

3. **类别不在科目上**。瑞云的 1601 只有一级科目、没有明细科目，
   类别藏在辅助核算里：``tb.aux1='机器设备'``，或 ``journal.aux_text='ZCLB02_SYS/机器设备'``。
   两处分隔符不同（竖线 / 斜杠），必须走同一个解析器。

修正完之后有一条硬不变量，本模块在返回前会自检：

    Σ期末 − Σ期初 == Σ序时账净额         （逐科目，精确到分）

这条不成立就说明去重或反推错了，直接抛异常，不往下走。
"""

from __future__ import annotations

import io
import json
import os
import re
from decimal import Decimal as D

#: 已知类别的**优先顺序**。它只决定列块谁放前面，**不再决定「有哪几类」** ——
#: 有哪几类由数据决定（TB 辅助核算 ∪ 序时账辅助核算，见 `load`）。
#: 为什么需要它：只按辅助核算编码排的话，瑞云的运输设备（ZCLB04）会排到
#: 其他设备（ZCLB05）前面，与历史底稿的列序相反，跨年不可比。
DEFAULT_PRIORITY = ("机器设备", "电子设备", "其他设备", "运输设备")

#: 伪类别：辅助核算里出现、但语义上不是资产类别。
#: 瑞影和广州分公司把**科目名**塞进了辅助核算位（`aux1='固定资产'`、
#: `aux1='累计折旧'`），天真的 distinct 扫描会把它们当成两个真类别，
#: 凭空多出两块，且 1601/1602 的同一批金额被算两遍。
PSEUDO_CATS = ("", "固定资产", "累计折旧", "减值准备", "累计折旧准备",
               "工程物资", "固定资产清理", "未分类", "合计")

#: 辅助核算里的编码段：`ZCLB02_SYS`、`ZCLB02`、`1601`。
#: 只用来拿排序键，**不用于识别类别** —— 瑞影的 aux1 就是裸中文名，
#: 只认编码会把绝大多数行判成「无类别」。
_SYS_RE = re.compile(r"_SYS$", re.I)
_CODE_RE = re.compile(r"^[A-Z]{2,6}\d{2,4}$")
_SPLIT_RE = re.compile(r"[|/]")
_AUX_CODE_RE = re.compile(r"(ZCLB\d{2,4})")

#: 三个科目：1601 原值 / 1602 累计折旧 / 1603 减值准备。
#: 注意减值准备是 1603，不是 1605（1605 是工程物资，瑞云余额为 0）。
ACCTS = ("1601", "1602", "1603")

#: 矩阵字段。金额一律**正数**（与 K.01 模板口径一致：原值/折旧/减值都填正，
#: 净值 = 原值 − 折旧 − 减值）。TB 里 1602/1603 是负数，进来就翻正。
FIELDS = ("fa_beg", "fa_end", "dep_beg", "dep_end", "imp_beg", "imp_end")


# ------------------------------------------------------------------ 配置

def _default_cfg() -> dict:
    return {
        "version": 0,
        "priority_order": list(DEFAULT_PRIORITY),
        "pseudo_categories": list(PSEUDO_CATS),
        "imp_category": "机器设备",
        "fallback_by_company": {"深圳市瑞影科技有限公司": "电子设备"},
        "defaults": {
            "life_months": [60, 120],
            "resid_rates": [0.0, 0.03, 0.05],
            "dept": ["行政部"],
            "loc": ["公司所在地"],
            "code_prefix": "",
            "avg_item_value": 500000,
            "count_min": 3,
            "count_max": 40,
        },
        "cats": {},
    }


_CFG: dict = {}
_ACTIVE: frozenset = frozenset()


def init(data_dir: str | None = None) -> dict:
    """加载 ``data_dir/fa_swp_cats.json``；文件不存在就用内置默认。

    配置里只有三件事和「账套」有关：优先顺序、伪类别名单、
    减值准备归哪一类 + 无辅助核算主体的兜底。其余（折旧年限、部门、
    编码前缀）是造明细用的，缺了用通用兜底，不会挡住出数。
    """
    global _CFG, PSEUDO_CATS
    conf = _default_cfg()
    pth = os.path.join(data_dir, "fa_swp_cats.json") if data_dir else None
    if pth and os.path.isfile(pth):
        with io.open(pth, encoding="utf-8") as f:
            ext = json.load(f)
        for k, v in ext.items():
            if k in ("cats", "defaults") and isinstance(v, dict):
                conf[k] = dict(conf.get(k) or {}, **v)
            else:
                conf[k] = v
    _CFG = conf
    if conf.get("pseudo_categories"):
        PSEUDO_CATS = tuple(conf["pseudo_categories"])
    return conf


def cfg(key: str | None = None):
    if not _CFG:
        init(None)
    return _CFG if key is None else _CFG.get(key)


def freeze(order) -> frozenset:
    """把「已确认的类别集合」冻成模块级状态，供 totals 拦脏键。"""
    global _ACTIVE
    _ACTIVE = frozenset(order)
    return _ACTIVE


def active_cats() -> frozenset:
    return _ACTIVE


# ------------------------------------------------------------------ 类别解析

#: 明细科目形态：``1601.01`` / ``1602.03``。瑞云是一级科目（``1601``），
#: 海柔是明细科目且类别直接写在科目名里（``固定资产_房屋及建筑物``）——
#: 两种都要认，否则换个客户就一行都取不到。
_KIND_RE = re.compile(r"^(160[123])(?:\.|$)")
_NAME_CAT_RE = re.compile(r"^[^_]+_(.+)$")


def acct_kind(code: str) -> str:
    """科目 → 段：``fa`` 原值 / ``dep`` 折旧 / ``imp`` 减值 / 空串。

    只认 1601/1602/1603 及其明细（``1601.01``）。1604 在建工程、1606 清理
    不在此列 —— 前缀匹配要求「1601/1602/1603 后面跟点或结束」就是为了不把 1604 吃进来。
    """
    m = _KIND_RE.match((code or "").strip())
    return {"1601": "fa", "1602": "dep", "1603": "imp"}.get(m.group(1), "") if m else ""


def cat_from_name(acct_name: str) -> str:
    """从科目名后缀取类别：``固定资产_房屋及建筑物`` → ``房屋及建筑物``。

    只对「一级名 + 下划线 + 类别」这种形态生效；``固定资产减值准备`` 这类没有
    下划线、或后缀本身是伪类别的，一律返回空串，交给调用方兜底，不硬猜。
    """
    m = _NAME_CAT_RE.match((acct_name or "").strip())
    if not m:
        return ""
    tail = m.group(1).split("_")[-1].strip()
    if not tail or tail in set(PSEUDO_CATS):
        return ""
    return tail


def cat_of(aux_text: str, acct_name: str = "") -> str:
    """从辅助核算文本里取类别。**纯解析，不依赖任何类别白名单。**

    这一点是这次改造的关键：原实现是 ``if p in CATS``，于是「识别」需要先知道
    类别集合，而集合又来自数据 —— 鸡生蛋。现在改成「排掉不像类别的段，
    第一个剩下的就是类别」，`load()` 之前调用也正确。

    认两种写法：``机器设备|ZCLB02_SYS``（竖线，TB）、
    ``ZCLB02_SYS/机器设备``（斜杠，序时账）。取不到返回空串，由调用方决定
    归集还是报错 —— 这里不猜。
    """
    s = (aux_text or "").strip()
    if not s:
        return ""
    acct = (acct_name or "").strip()
    pseudo = set(PSEUDO_CATS)
    for part in _SPLIT_RE.split(s):
        p = part.strip()
        if not p:
            continue
        if _SYS_RE.search(p) or _CODE_RE.match(p):
            continue                        # ZCLB02_SYS / ZCLB02
        if p.isdigit():
            continue                        # 1601 / 1602
        if p in pseudo:
            continue                        # 固定资产 / 累计折旧
        if acct and p == acct:
            continue                        # 与本行科目同名 → 伪类别
        return p
    return ""


def aux_code(row: dict) -> str:
    """取辅助核算编码（``ZCLB02``），只用于列块排序。取不到返回空串。"""
    m = _AUX_CODE_RE.search("%s|%s" % (row.get("aux2") or "", row.get("aux_text") or ""))
    return m.group(1) if m else ""


def category(row: dict) -> str:
    """一行的类别。优先 aux1（tb 已解析好），再退 aux_text，最后按主体兜底。"""
    acct_name = row.get("acct_name", "")
    c = (cat_of(row.get("aux1"), acct_name)
         or cat_of(row.get("aux_text"), acct_name)
         or cat_from_name(acct_name))
    if c:
        return c
    if acct_kind(row["acct_code"]) == "imp":
        # 减值准备完全没有辅助核算，归属只能靠配置 —— 瑞云的依据是
        # 1603 的 47 张凭证全部与 1601「机器设备」的处置凭证同凭证号。
        return cfg("imp_category") or ""
    return (cfg("fallback_by_company") or {}).get(row["company_name"], "")


def order_cats(cats, cat_code: dict | None = None) -> list[str]:
    """已知类别按优先序在前，新类别按辅助核算编码（ZCLB02<03<…）追加。

    **不要用 set**：set 的迭代顺序不稳定，同样输入两次跑出不同列序，
    跨年底稿就不可复现了。
    """
    cats = list(dict.fromkeys(cats))          # 去重且保序
    code = cat_code or {}
    pri = [c for c in (cfg("priority_order") or ()) if c in cats]
    rest = [c for c in cats if c not in pri]
    rest.sort(key=lambda c: (0, code[c], c) if code.get(c) else (1, "", c))
    return pri + rest


def cat_codes(rows: list[dict]) -> dict[str, str]:
    """类别 → 辅助核算编码。同一类别多行时取第一个有编码的。"""
    out = {}
    for r in rows:
        c = category(r)
        if not c or c in out:
            continue
        code = aux_code(r)
        if code:
            out[c] = code
    return out


# ------------------------------------------------------------------ 读 TB

TB_SQL = """
select company_name,
       acct_code,
       coalesce(acct_name, '') as acct_name,
       coalesce(aux1, '') as aux1,
       coalesce(aux2, '') as aux2,
       coalesce(aux_text, '') as aux_text,
       coalesce(open_bal, 0) as open_bal,
       coalesce(end_bal, 0) as end_bal,
       coalesce(cur_debit, 0) as cur_debit,
       coalesce(cur_credit, 0) as cur_credit,
       src_row
  from tb
 where acct_code like '1601%' or acct_code like '1602%' or acct_code like '1603%'
 order by acct_code, company_name, src_row
"""


def load_tb(conn) -> list[dict]:
    """取 TB 里固定资产三个科目的全部行（含重复行，去重交给 dedupe）。"""
    cur = conn.execute(TB_SQL)
    cols = [d[0] for d in cur.description]
    out = []
    for rec in cur.fetchall():
        r = dict(zip(cols, rec))
        for k in ("open_bal", "end_bal", "cur_debit", "cur_credit"):
            r[k] = D(str(r[k]))
        r["src_row"] = int(r["src_row"])
        out.append(r)
    return out


def _dedupe_key(r: dict) -> tuple:
    """去重键：主体 + 科目 + 辅助核算编码 + 四个金额，全等才算重复。

    ``aux2`` 必须在内。瑞云母公司有 4 行 1601（机器/电子/其他/运输），
    其中「运输设备」期初期末都是 361,790.12 —— 若只按主体+科目+金额去重，
    会把两个真实类别合并掉。
    """
    return (r["company_name"], r["acct_code"], r["aux2"],
            r["open_bal"], r["end_bal"], r["cur_debit"], r["cur_credit"])


def dedupe(rows: list[dict]) -> tuple[list[dict], list[dict]]:
    """整块重复只保留 src_row 最小的那条。返回 (保留, 删除)。"""
    seen, kept, dropped = {}, [], []
    for r in sorted(rows, key=lambda x: (x["acct_code"], x["company_name"], x["src_row"])):
        k = _dedupe_key(r)
        if k in seen:
            r = dict(r)
            r["_dup_of"] = seen[k]
            dropped.append(r)
        else:
            seen[k] = r["src_row"]
            kept.append(r)
    return kept, dropped


# --------------------------------------------------------------- 序时账净额

JNET_SQL = """
select company_name, acct_code,
       sum(coalesce(debit, 0)) - sum(coalesce(credit, 0)) as net
  from journal
 where acct_code like '1601%' or acct_code like '1602%' or acct_code like '1603%'
 group by 1, 2
"""


def journal_net(conn) -> dict[tuple[str, str], D]:
    """每个 (主体, 科目) 的本年发生净额（借 − 贷）。"""
    out = {}
    for comp, acct, net in conn.execute(JNET_SQL).fetchall():
        out[(comp, acct)] = D(str(net or 0))
    return out


def reconcile_opening(rows: list[dict], jnet: dict) -> tuple[list[dict], list[dict]]:
    """用序时账反推期初，修正 TB 里不可信的 open_bal。

    ``expect_open = end_bal − net``（原值科目）；
    1602/1603 是贷方科目，TB 里存的本来就是负数，公式同样成立，
    因为 ``net`` 也是带符号的（贷方为负）。

    只在该 (主体, 科目) 的合计差额 > 0.01 时修正，差额记到该组最后一行上，
    并写进 adjusts 返回 —— 让「我改过哪一格」看得见，不静默修改。
    """
    groups: dict[tuple, list[dict]] = {}
    for r in rows:
        groups.setdefault((r["company_name"], r["acct_code"]), []).append(r)

    adjusts = []
    for (comp, acct), grp in groups.items():
        ob = sum((r["open_bal"] for r in grp), D("0"))
        eb = sum((r["end_bal"] for r in grp), D("0"))
        net = jnet.get((comp, acct), D("0"))
        expect = eb - net
        diff = ob - expect
        if abs(diff) <= D("0.01"):
            continue
        last = grp[-1]
        adjusts.append({
            "company": comp,
            "acct": acct,
            "tb_open": str(ob),
            "expect_open": str(expect),
            "adjust": str(-diff),
            "applied_to_src_row": last["src_row"],
            "reason": "序时账反推：期末 − 本年发生净额",
        })
        last["open_bal"] = last["open_bal"] - diff
    return rows, adjusts


# ------------------------------------------------------------------ 汇总

def aggregate(rows: list[dict], trace: list | None = None) -> dict:
    """按类别汇总成目标矩阵，金额全部翻正。

    返回 ``{类别: {fa_beg: Decimal, ...}}``，另把「未归类」的行抛出来。
    """
    m = {}
    unknown = []
    for r in rows:
        cat = category(r)
        if not cat:
            # 余额全 0 的主体（如瑞影广州分公司，本期无固定资产）
            # 不影响任何金额，跳过并留痕；有余额还归不进类别才是真问题。
            if (r["open_bal"] == 0 and r["end_bal"] == 0
                    and r["cur_debit"] == 0 and r["cur_credit"] == 0):
                if trace is not None:
                    trace.append({
                        "kind": "skip_zero", "company": r["company_name"],
                        "acct": r["acct_code"], "to": "",
                        "amount": "0", "reason": "余额与发生额均为 0，不参与汇总",
                    })
                continue
            unknown.append(r)
            continue
        # 懒建键：类别不再预置。预置一个固定元组就等于把「有哪几类」写死，
        # 数据里第 5 个类别会被 KeyError 挡下来（或者更糟，被静默丢掉）。
        if cat not in m:
            m[cat] = {f: D("0") for f in FIELDS}
        acct = r["acct_code"]
        kind = acct_kind(acct)
        if kind == "fa":
            m[cat]["fa_beg"] += r["open_bal"]
            m[cat]["fa_end"] += r["end_bal"]
        elif kind == "dep":
            # TB 里折旧是负数，翻正
            m[cat]["dep_beg"] += -r["open_bal"]
            m[cat]["dep_end"] += -r["end_bal"]
        else:
            m[cat]["imp_beg"] += -r["open_bal"]
            m[cat]["imp_end"] += -r["end_bal"]
        if trace is not None:
            acct_name = r.get("acct_name", "")
            if not cat_of(r.get("aux1"), acct_name) and not cat_of(r.get("aux_text"), acct_name):
                trace.append({
                    "kind": "cat_fallback",
                    "company": r["company_name"],
                    "acct": acct,
                    "to": cat,
                    "amount": str(r["end_bal"]),
                    "reason": "无类别辅助核算，按主体归集" if acct != "1603"
                              else "减值准备无辅助核算，按配置归集",
                })
    return {"cats": m, "unknown": unknown}


def totals(m: dict) -> dict:
    """矩阵合计。

    遍历 ``m`` 本身的键，**不是**遍历一个固定元组 —— 遍历固定元组时，矩阵里
    多出来的类别会被静默漏算：合计变小、而 K.01 的横向合计是按列块加的，
    最终表现为「表1 合计 ≠ TB 合计」，很晚才炸、且看不出原因。
    用 ``_ACTIVE`` 反向拦脏键（未经过 `load()` 确认的类别），两头都堵住。
    """
    if _ACTIVE:
        bad = set(m) - _ACTIVE
        if bad:
            raise ValueError("矩阵里出现未经确认的类别键：%s" % sorted(bad))
    return {f: sum((m[c][f] for c in m), D("0")) for f in FIELDS}


def check_invariants(rows: list[dict], jnet: dict) -> dict:
    """硬不变量：逐科目「Σ期末 − Σ期初 == Σ序时账净额」，返回差异清单。

    （原来这里还声称要校验「矩阵合计 == EXPECT_RUIYUN」，但那组硬编码基准
    从未被真正比较过 —— 是死代码，已删。瑞云基准现在只活在回归测试里，
    生产代码不该认识某个客户的具体金额。）
    """
    problems = []
    for acct in ACCTS:
        grp = [r for r in rows if r["acct_code"] == acct]
        ob = sum((r["open_bal"] for r in grp), D("0"))
        eb = sum((r["end_bal"] for r in grp), D("0"))
        net = sum((v for (c, a), v in jnet.items() if a == acct), D("0"))
        diff = (eb - ob) - net
        if abs(diff) > D("0.01"):
            problems.append({
                "kind": "journal_mismatch", "acct": acct,
                "tb_delta": str(eb - ob), "journal_net": str(net), "diff": str(diff),
            })
    return {"ok": not problems, "problems": problems}


def load(conn) -> dict:
    """一站式：读 → 去重 → 反推 → 归类 → 排序 → 自检。返回可直接用的结果包。

    ``cat_order`` 是**本次唯一权威的类别顺序**，下游（moves / synth / k01 /
    lead / checks）一律用它，不再有模块级常量。它在 `totals` 之后才 freeze，
    所以 load 内部的合计不会被自己拦；外部再拿矩阵来算就会检查脏键。
    """
    raw = load_tb(conn)
    kept, dropped = dedupe(raw)
    jnet = journal_net(conn)
    kept, adjusts = reconcile_opening(kept, jnet)

    trace = []
    agg = aggregate(kept, trace)
    if agg["unknown"]:
        names = sorted({(r["company_name"], r["acct_code"]) for r in agg["unknown"]})
        raise ValueError("有固定资产行归不进任何类别，不能猜：%s" % (names,))

    inv = check_invariants(kept, jnet)
    if not inv["ok"]:
        raise ValueError("去重/反推后与序时账不勾稽：%s" % (inv["problems"],))

    m = agg["cats"]
    t = totals(m)
    codes = cat_codes(kept)
    order = order_cats(list(m), codes)
    freeze(order)

    return {
        "rows": kept,
        "cats": m,
        "cat_order": order,
        "cat_code": codes,
        "totals": t,
        "net_beg": t["fa_beg"] - t["dep_beg"] - t["imp_beg"],
        "net_end": t["fa_end"] - t["dep_end"] - t["imp_end"],
        "trace": {
            "dropped": [
                {"company": r["company_name"], "acct": r["acct_code"],
                 "src_row": r["src_row"], "dup_of": r["_dup_of"],
                 "open": str(r["open_bal"]), "end": str(r["end_bal"])}
                for r in dropped
            ],
            "adjusts": adjusts,
            "cat_rules": trace,
            "invariants": inv,
            "cat_order": order,
            "cat_code": codes,
        },
    }
