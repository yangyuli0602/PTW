# -*- coding: utf-8 -*-
"""回归基准：把一次「正确的」生成结果记下来，以后每次改完代码都能证明输出没变。

对应 Skill §3.4 的"回归不变量"。为什么必须单独存：`_fill_bkd`、`expense`
改一行就可能让某个数偏移 0.01，而人眼看不出——只有和基准逐个比才发现。

基准只记**关键数**，不记全量明细：全量会把 baseline 变成第二个数据库，反而没人看。
"""
import json
import os
from datetime import datetime

BASE_VERSION = 1
TOL_AMOUNT = 0.005      # 金额容差（分）
TOL_RATIO = 1e-6


def _f(x, d=None):
    try:
        return float(x)
    except (TypeError, ValueError):
        return d


def capture(kind, info):
    """从生成结果里抽取要长期盯住的那几个数。"""
    info = info or {}
    vals = {}
    if kind == "cash":
        vals = {
            "rows": info.get("rows"),
            "gaps": info.get("gaps"),
            "amount": _f(info.get("amount")),
            "rmbAmount": _f(info.get("rmbAmount")),
            "fxCount": info.get("fxCount"),
            "bySubject": {k: _f(v) for k, v in (info.get("bySubject") or {}).items()},
            "stats": {k: _f((info.get("stats") or {}).get(k))
                      for k in ("matched", "newCount", "missCount", "priorOnlyCount",
                                "sumTb", "sumPrior", "diff")},
            "cutoff": {k: (info.get("cutoff") or {}).get(k)
                       for k in ("enabled",)},
        }
        c = info.get("cutoff") or {}
        st = c.get("stat") or {}
        vals["cutoff"].update({"before": len(c.get("before") or []),
                               "after": len(c.get("after") or []),
                               "crossCount": st.get("crossCount"),
                               "threshold": _f(st.get("threshold"))})
    else:
        vals = {
            "vc": info.get("vc"), "vd": info.get("vd"),
            "other": info.get("other"),
            "amount": _f(info.get("amount")),
            "vc_sum": _f(info.get("vc_sum")), "vd_sum": _f(info.get("vd_sum")),
            "gaps": info.get("gaps"),
            "closeOut": _f(info.get("closeOut")),
        }
    return {"baseVersion": BASE_VERSION, "kind": kind,
            "capturedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "values": vals}


# ---------------------------------------------------------------- 存/取
def path_for(data_dir, kind):
    d = os.path.join(data_dir, "baseline")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, kind + ".json")


def save(base, path):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(base, fh, ensure_ascii=False, indent=1)
    return path


def load(path):
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as fh:
            d = json.load(fh)
    except Exception:
        return None
    if (d or {}).get("baseVersion") != BASE_VERSION:
        return None
    return d


# ---------------------------------------------------------------- 比对
def _walk(a, b, prefix, out):
    if isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b)):
            _walk(a.get(k), b.get(k), prefix + "." + k if prefix else k, out)
        return
    # 数值比较
    fa, fb = _f(a), _f(b)
    if fa is not None or fb is not None:
        if fa is None or fb is None:
            out.append({"key": prefix, "base": a, "cur": b, "level": "changed"})
            return
        d = fb - fa
        if abs(d) <= TOL_AMOUNT:
            return
        out.append({"key": prefix, "base": fa, "cur": fb, "delta": round(d, 6),
                    "pct": (round(d / fa * 100, 4) if abs(fa) > TOL_RATIO else None),
                    "level": "changed"})
        return
    if a != b:
        out.append({"key": prefix, "base": a, "cur": b, "level": "changed"})


def compare(cur, base):
    """cur/base 都是 capture() 的结果。返回差异列表，空表示完全一致。"""
    if not base:
        return [{"key": "-", "level": "info", "detail": "还没有基准，本次结果已记录"}]
    out = []
    _walk(base.get("values") or {}, (cur or {}).get("values") or {}, "", out)
    return [d for d in out if d.get("key")]


def summarize(diffs):
    if not diffs:
        return "与基准完全一致"
    if len(diffs) == 1 and diffs[0].get("level") == "info":
        return diffs[0].get("detail") or "已建立基准"
    return "%d 处与基准不同" % len(diffs)
