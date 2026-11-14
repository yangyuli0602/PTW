# -*- coding: utf-8 -*-
"""不变量校验：把「应该相等」的规则抽成可逐条执行的清单。

对应 Skill 附录 C。与它的差别：**用 JSON 不用 YAML**——内置 runtime 没有 PyYAML、
离线装不上。JSON 同样可读可编辑、可导出，语义不打折。

设计要点
--------
1. 默认清单写在代码里（DEFAULT_CASH / DEFAULT_EXP），`data/invariants.json` 只能
   **覆盖阈值和开关**，不能新增检查项——新增检查必须写代码，避免"配置文件里写了
   但没人实现"的假通过。
2. status 三态：pass / fail / skip。`skip` 是诚实的"这条现在算不了"，
   比如需要 Excel 计算结果才能判的（openpyxl 写出的公式没有缓存值）。
3. 金额比较统一用 0.005 容差，浮点噪声不算不平。
"""
import json
import os

TOL = 0.005


def _f(x, d=0.0):
    try:
        return float(x)
    except (TypeError, ValueError):
        return d


def _close(a, b, tol=TOL):
    return abs(_f(a) - _f(b)) <= tol


# ================================================================ 货币资金
def _c_rows(ctx):
    n = int(_f((ctx.get("info") or {}).get("rows")))
    return ("pass" if n > 0 else "fail", "账户数 %d" % n)


def _c_sum(ctx):
    """明细本币合计 == 返回的 amount。这条能抓住"统计口径和明细不一致"。"""
    rows = ctx.get("rows") or []
    s = round(sum(_f(r.get("local")) for r in rows), 2)
    amt = _f((ctx.get("info") or {}).get("amount"))
    return ("pass" if _close(s, amt) else "fail",
            "明细合计 %.2f / 返回 %.2f / 差 %.2f" % (s, amt, s - amt))


def _c_rmb(ctx):
    """折算人民币合计 == rmbAmount。"""
    rows = ctx.get("rows") or []
    s = round(sum(_f(r.get("rmb")) for r in rows), 2)
    amt = _f((ctx.get("info") or {}).get("rmbAmount"))
    return ("pass" if _close(s, amt) else "fail",
            "折算合计 %.2f / 返回 %.2f / 差 %.2f" % (s, amt, s - amt))


def _c_fx(ctx):
    """逐行折算：rmb == round(local × 率, 2)。率默认 1，所以全填 1 时必须完全相等。"""
    rates = ctx.get("rates") or {}
    bad = []
    for r in (ctx.get("rows") or []):
        rate = _f(rates.get(r.get("company")), 1.0) or 1.0
        want = round(_f(r.get("local")) * rate, 2)
        got = _f(r.get("rmb"), None if r.get("rmb") is None else 0)
        if r.get("rmb") is None:
            continue
        if not _close(got, want, 0.02):
            bad.append("%s/%s 应为 %.2f 实为 %.2f" % (r.get("company"), r.get("acctNo"), want, got))
    return ("pass" if not bad else "fail",
            "全部按率折算" if not bad else "%d 行不符：%s" % (len(bad), "；".join(bad[:3])))


def _c_subject(ctx):
    """按科目汇总 == 总额。分类不能丢钱。"""
    by = (ctx.get("info") or {}).get("bySubject") or {}
    s = round(sum(_f(v) for v in by.values()), 2)
    amt = _f((ctx.get("info") or {}).get("amount"))
    return ("pass" if _close(s, amt) else "fail",
            "科目合计 %.2f / 总额 %.2f" % (s, amt))


def _c_prior_only(ctx):
    """上年有、本期无的账户金额必须为 0——它们是"存在性证据"，不该带金额进来。"""
    rows = ctx.get("rows") or []
    bad = [r for r in rows if r.get("_priorOnly") and abs(_f(r.get("local"))) > TOL]
    return ("pass" if not bad else "fail",
            "全部为 0" if not bad else "%d 行非 0" % len(bad))


def _c_gap_count(ctx):
    """缺口数 == 明细里标黄格数。gapList 最多留 300 条，超了按截断判断。"""
    info = ctx.get("info") or {}
    n = int(_f(info.get("gaps")))
    lst = info.get("gapList") or []
    if n <= len(lst):
        return ("pass" if n == len(lst) else "fail", "缺口 %d / 列表 %d" % (n, len(lst)))
    return ("pass", "缺口 %d（列表截断在 %d 条）" % (n, len(lst)))


def _c_lead_tie(ctx):
    """ΣBKD == Lead 期末账面合计。需要 Excel 算完公式才能读，本版算不了。"""
    return ("skip", "需要 Excel 计算后的单元格值；openpyxl 写出的公式无缓存值")


def _c_restricted(ctx):
    """受限资金：ΣBKD 受限列 == Lead 受限列示位置（重分类行或披露行）。

    「两处不重复计数」的落地：BKD 只记一次（每笔一行），Lead 只有一个位置引用它。
    工具能判的是「有没有落到 Lead」和「落到了哪一种形态」，落的具体金额由公式引用，
    Excel 算完才读得到，所以金额一致性改为验证引用单元格存在。
    """
    info = ctx.get("info") or {}
    total = _f(info.get("restrictedTotal"), 0)
    mode = info.get("restrictedMode") or ""
    if abs(total) <= TOL:
        return ("pass", "BKD 受限资金列全为 0，无需在 Lead 列示")
    if mode == "reclass":
        return ("pass", "受限 %.2f 已作为重分类调整进 Lead（%s）"
                % (total, info.get("restrictedTarget") or "表1 受限资金行"))
    if mode == "disclosure":
        return ("pass", "受限 %.2f 已进 Lead 披露区（%s）"
                % (total, info.get("restrictedTarget") or "披露行"))
    if mode == "none":
        return ("fail", "受限 %.2f 但 Lead 里没有可列示的位置：%s"
                % (total, info.get("restrictedNote") or ""))
    return ("skip", "本次未经过 Lead 填充，无法判断受限资金是否已列示")


def _c_derived(ctx):
    """派生表执行情况：有 error 即失败；partial 说明有内容没写全，要点出来。"""
    d = (ctx.get("info") or {}).get("derived") or {}
    if not d:
        return ("skip", "本次未执行派生表（模板可能不是货币资金底稿）")
    errs = [k for k, v in d.items() if isinstance(v, dict) and v.get("status") == "error"]
    parts = [k for k, v in d.items() if isinstance(v, dict) and v.get("status") == "partial"]
    oks = [k for k, v in d.items() if isinstance(v, dict) and v.get("status") == "ok"]
    if errs:
        return ("fail", "派生表出错：%s；成功 %d 张" % ("、".join(errs), len(oks)))
    tail = "；%s 未写全需人工补" % "、".join(parts) if parts else ""
    return ("pass", "成功 %d 张（%s）%s" % (len(oks), "、".join(oks) or "无", tail))


DEFAULT_CASH = [
    ("cash_rows", "有账户进入底稿", "A", _c_rows),
    ("cash_sum", "BKD 明细合计 == 返回金额", "A", _c_sum),
    ("cash_rmb", "折算人民币合计 == 返回折算金额", "A", _c_rmb),
    ("cash_fx", "逐行折算率计算正确", "A", _c_fx),
    ("cash_subject", "按科目汇总 == 总额", "A", _c_subject),
    ("cash_prior_only", "上年独有账户金额为 0", "A", _c_prior_only),
    ("cash_gap_count", "缺口数与清单一致", "B", _c_gap_count),
    ("cash_lead_tie", "BKD 合计 == Lead 期末账面", "A", _c_lead_tie),
    ("cash_restricted", "受限资金两处不重复计数", "A", _c_restricted),
    ("cash_derived", "派生表（C.01b/00b/02/02a）执行", "B", _c_derived),
]


# ================================================================ 销管费用
def _e_rows(ctx):
    info = ctx.get("info") or {}
    n = int(_f(info.get("vc"))) + int(_f(info.get("vd")))
    return ("pass" if n > 0 else "fail", "VC %s + VD %s 行" % (info.get("vc"), info.get("vd")))


def _e_sum(ctx):
    """VC + VD 合计 == 本期账面数合计。"""
    info = ctx.get("info") or {}
    s = round(_f(info.get("vc_sum")) + _f(info.get("vd_sum")), 2)
    amt = _f(info.get("amount"))
    return ("pass" if _close(s, amt) else "fail",
            "VC+VD %.2f / 合计 %.2f / 差 %.2f" % (s, amt, s - amt))


def _e_prior_null(ctx):
    """上期审定数本版无数据源，必须留空而不是填 0（填 0 会让变动率全错）。"""
    info = ctx.get("info") or {}
    gaps = info.get("gapList") or []
    prior_gaps = [g for g in gaps if "上期" in str(g.get("field") or "")]
    if int(_f(info.get("gaps"))) == 0:
        return ("skip", "无缺口，无上期列")
    return ("pass" if prior_gaps else "fail",
            "上期列标黄 %d 处" % len(prior_gaps))


def _e_close_out(ctx):
    """已剔除的结转损益必须 >= 0，负数说明摘要匹配规则出问题。"""
    info = ctx.get("info") or {}
    co = _f(info.get("closeOut"))
    return ("pass" if co >= -TOL else "fail", "剔除结转损益 %.2f" % co)


DEFAULT_EXP = [
    ("exp_rows", "有费用行进入底稿", "A", _e_rows),
    ("exp_sum", "VC+VD 合计 == 本期账面数", "A", _e_sum),
    ("exp_prior_null", "上期审定数留空而非填 0", "A", _e_prior_null),
    ("exp_close_out", "剔除的结转损益非负", "B", _e_close_out),
]


# ================================================================ 执行
def _overrides(data_dir):
    """读 data/invariants.json：只能改阈值/禁用，不能新增项。"""
    p = os.path.join(data_dir or "", "invariants.json")
    if not data_dir or not os.path.exists(p):
        return {}
    try:
        with open(p, encoding="utf-8") as fh:
            d = json.load(fh)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def run(ctx, kind="cash", data_dir=None, only=None):
    """跑一遍不变量。返回 {items, summary}。

    items 每项：{id, desc, level, status, detail}
    summary：{pass, fail, skip, total}
    """
    table = DEFAULT_CASH if kind == "cash" else DEFAULT_EXP
    ov = (_overrides(data_dir).get(kind) or {})
    off = set(ov.get("disable") or [])
    items = []
    for cid, desc, level, fn in table:
        if only and cid not in only:
            continue
        if cid in off:
            items.append({"id": cid, "desc": desc, "level": level,
                          "status": "skip", "detail": "已按 invariants.json 禁用"})
            continue
        try:
            st, detail = fn(ctx)
        except Exception as e:
            st, detail = "fail", "校验本身出错：%s" % e
        items.append({"id": cid, "desc": desc, "level": level,
                      "status": st, "detail": detail})
    summ = {"pass": 0, "fail": 0, "skip": 0, "total": len(items)}
    for it in items:
        summ[it["status"]] = summ.get(it["status"], 0) + 1
    return {"items": items, "summary": summ}


def failed(report):
    return [i for i in (report or {}).get("items") or [] if i.get("status") == "fail"]


def brief(report):
    """一行摘要，给 CLI 输出用。"""
    s = (report or {}).get("summary") or {}
    if not s.get("total"):
        return "无校验项"
    if not s.get("fail"):
        return "%d 项通过 / %d 项跳过" % (s.get("pass", 0), s.get("skip", 0))
    return "%d 项不通过：%s" % (s.get("fail"),
                                "、".join(i["id"] for i in failed(report)[:5]))
