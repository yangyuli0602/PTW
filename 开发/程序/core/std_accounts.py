# -*- coding: utf-8 -*-
"""标准一级科目表：内置一份，客户可在 data/ 放同名文件覆盖 / 扩充。

为什么独立成模块：clean.py 是流水线执行器，不该知道文件系统；这份表以后映射页、
报表页也可能要用。加载带 mtime 缓存，避免每步清洗都读盘。
"""
import json
import os

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 程序/
BUILTIN = os.path.join(_ROOT, "rules", "std_accounts.json")
OVERRIDE_NAME = "std_accounts.json"

_CACHE = {}          # (内置信, 覆盖路径, 覆盖信) -> (mapping, info)
_CACHE_MAX = 8


def _sig(path):
    try:
        st = os.stat(path)
        return (st.st_mtime_ns, st.st_size)
    except OSError:
        return None


def _parse(path):
    """一份表 → {code: name}。支持 {"accounts": {"1001": "库存现金"}} 或
    {"accounts": [{"code": ..., "name": ...}]} 两种写法。

    坏文件 / 坏条目一律跳过：这份表只决定「名称好不好看」，绝不能让它把清洗整条挂掉。
    """
    try:
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)
    except Exception:
        return {}
    acc = raw.get("accounts") if isinstance(raw, dict) else raw
    if isinstance(acc, dict):
        items = [{"code": k, "name": v} for k, v in acc.items()]
    elif isinstance(acc, list):
        items = acc
    else:
        return {}
    out = {}
    for x in items:
        if not isinstance(x, dict):
            continue
        code = str(x.get("code") or "").strip()
        name = str(x.get("name") or "").strip()
        if code and name:
            out[code] = name
    return out


def override_path(data_dir):
    return os.path.join(data_dir, OVERRIDE_NAME) if data_dir else ""


def load(data_dir=None, with_info=False):
    """合并顺序：内置 → data/ 覆盖（同码覆盖、新码追加）。"""
    op = override_path(data_dir)
    key = (_sig(BUILTIN), op, _sig(op) if op else None)
    hit = _CACHE.get(key)
    if hit is None:
        base = _parse(BUILTIN)
        n_base = len(base)
        ov = _parse(op) if op and os.path.exists(op) else {}
        base.update(ov)
        info = {"builtin": n_base, "override": len(ov),
                "override_path": op if ov else "", "total": len(base)}
        hit = (base, info)
        if len(_CACHE) >= _CACHE_MAX:
            _CACHE.clear()
        _CACHE[key] = hit
    return hit if with_info else hit[0]


def sql_map(mapping, chunk=800):
    """{code: name} → DuckDB 的 map 表达式。

    用 map 而不是内联 CASE：查表 O(1)、SQL 长度只随条目线性增长；也不用临时表，
    不会在连接里留东西。键唯一由 dict 保证（map_from_entries 遇重复键会报错）。
    """
    if not mapping:
        return "CAST(NULL AS MAP(VARCHAR, VARCHAR))"
    items = list(mapping.items())
    parts = []
    for i in range(0, len(items), chunk):
        ks = ", ".join("'%s'" % str(k).replace("'", "''") for k, _ in items[i:i + chunk])
        vs = ", ".join("'%s'" % str(v).replace("'", "''") for _, v in items[i:i + chunk])
        parts.append("map_from_entries(list_zip([%s], [%s]))" % (ks, vs))
    return parts[0] if len(parts) == 1 else "map_concat(%s)" % ", ".join(parts)
