# -*- coding: utf-8 -*-
"""版本指纹：把「底稿模板长什么样」固化成可比对的一小段数据。

用途（对应 Skill §3.1 步骤 4、§7）：
  1. 每次生成后落一份指纹；下次跑之前先比对
  2. 指纹不一致 → 只输出差异报告，不写数（`--strict` 时）
  3. 报告要能说清"是哪张表、哪一列、从什么变成什么"，而不是一句"结构变了"

只记**结构**，不记数据：表名、表头行、表头文字的哈希、行列数、外部链接目标、
参数（TE/期间）。数据行变了不算漂移——同一模板填不同客户的数据，指纹必须一样。
"""
import hashlib
import io
import json
import os
from datetime import datetime

#: 指纹格式版本。字段含义变了就 +1，老指纹直接失效而不是误比对
FP_VERSION = 1


def _norm(s):
    """表头文字归一：去空白/换行/全角空格，避免「期末余额（本币）」和「期末余额(本币)」被当成不同。"""
    if s is None:
        return ""
    t = str(s).replace("\u3000", " ").replace("\xa0", " ")
    t = t.replace("（", "(").replace("）", ")").replace("：", ":")
    return "".join(t.split())


def _header_row(ws, limit=20):
    """表头行 = 前 limit 行里非空单元格最多的那一行。

    不依赖任何 spec：指纹只关心"结构稳定"，不关心"这列是什么语义"。
    语义映射是 _locate_bkd_cols 的事，两件事不要混。
    """
    best, best_n = 1, -1
    for r in range(1, min(limit, (ws.max_row or 1)) + 1):
        n = 0
        for c in range(1, min(ws.max_column or 1, 60) + 1):
            if _norm(ws.cell(r, c).value):
                n += 1
        if n > best_n:
            best, best_n = r, n
    return best


def _headers(ws, row):
    out = []
    for c in range(1, (ws.max_column or 0) + 1):
        t = _norm(ws.cell(row, c).value)
        if t:
            out.append("%d:%s" % (c, t))
    return out


def sheet_fingerprint(ws):
    hr = _header_row(ws)
    heads = _headers(ws, hr)
    h = hashlib.sha1("\u0001".join(heads).encode("utf-8")).hexdigest()[:16]
    return {"headerRow": hr, "headersHash": h, "headers": heads,
            "rows": ws.max_row or 0, "cols": ws.max_column or 0}


def of_workbook(wb, params=None):
    """从 openpyxl 工作簿提指纹。"""
    sheets = []
    for ws in wb.worksheets:
        try:
            fp = sheet_fingerprint(ws)
        except Exception as e:                       # 图表页等异常表：记下来但不阻断
            fp = {"error": str(e)[:120]}
        fp["name"] = ws.title
        sheets.append(fp)
    links = []
    try:
        for l in (wb._external_links or []):
            tgt = getattr(l, "file_link", None) or getattr(l, "Target", None)
            if tgt:
                links.append(os.path.basename(str(tgt)))
    except Exception:
        pass
    return {"fpVersion": FP_VERSION,
            "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "sheets": sheets,
            "links": sorted(set(links)),
            "params": params or {}}


def of_bytes(data, params=None):
    """从生成的 xlsx 字节提指纹。生成器拿到的就是 bytes，不用再落盘读。"""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), data_only=False, read_only=False)
    try:
        return of_workbook(wb, params=params)
    finally:
        try:
            wb.close()
        except Exception:
            pass


def diff(old, new):
    """比对两份指纹，返回结构化差异。

    三种差异分开报：多了/少了整张 sheet、表头行变了、表头文字变了（逐列说清从什么变成什么）。
    行数变化单独一类，因为它通常是"填了不同客户的数据"，多数情况不该拦。
    """
    out = []
    o = {s.get("name"): s for s in (old or {}).get("sheets") or []}
    n = {s.get("name"): s for s in (new or {}).get("sheets") or []}
    for name in sorted(set(o) - set(n)):
        out.append({"kind": "sheet_removed", "sheet": name, "level": "block",
                    "detail": "模板里没有这张表了"})
    for name in sorted(set(n) - set(o)):
        out.append({"kind": "sheet_added", "sheet": name, "level": "info",
                    "detail": "模板新增这张表"})
    for name in sorted(set(o) & set(n)):
        a, b = o[name], n[name]
        if a.get("headersHash") and b.get("headersHash") and a["headersHash"] != b["headersHash"]:
            ha = {int(x.split(":", 1)[0]): x.split(":", 1)[1] for x in a.get("headers") or []}
            hb = {int(x.split(":", 1)[0]): x.split(":", 1)[1] for x in b.get("headers") or []}
            cols = []
            for c in sorted(set(ha) | set(hb)):
                if ha.get(c) != hb.get(c):
                    cols.append({"col": c, "old": ha.get(c), "new": hb.get(c)})
            out.append({"kind": "headers_changed", "sheet": name, "level": "block",
                        "headerRow": {"old": a.get("headerRow"), "new": b.get("headerRow")},
                        "cols": cols[:40],
                        "detail": "表头文字变了 %d 列" % len(cols)})
        elif a.get("headerRow") != b.get("headerRow"):
            out.append({"kind": "header_row_moved", "sheet": name, "level": "block",
                        "detail": "表头行 %s → %s" % (a.get("headerRow"), b.get("headerRow"))})
        if a.get("rows") != b.get("rows"):
            out.append({"kind": "row_count_changed", "sheet": name, "level": "info",
                        "detail": "行数 %s → %s（通常是数据量不同，不算结构漂移）"
                                  % (a.get("rows"), b.get("rows"))})
    if (old or {}).get("links") != (new or {}).get("links"):
        out.append({"kind": "links_changed", "level": "warn",
                    "detail": "外部链接 %s → %s" % ((old or {}).get("links"),
                                                    (new or {}).get("links"))})
    return out


def blocking(diffs):
    """会不会影响写入。header/sheet 变化算 blocker，行数和新增表不算。"""
    return [d for d in diffs or [] if d.get("level") == "block"]


# ---------------------------------------------------------------- 落盘
def path_for(data_dir, subject, tag=""):
    d = os.path.join(data_dir, "fp")
    os.makedirs(d, exist_ok=True)
    nm = subject + ("_" + tag if tag else "")
    return os.path.join(d, nm + ".json")


def save(fp, path):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(fp, fh, ensure_ascii=False, indent=1)
    return path


def load(path):
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as fh:
            d = json.load(fh)
    except Exception:
        return None
    if (d or {}).get("fpVersion") != FP_VERSION:
        return None
    return d


def save_and_diff(data_dir, subject, fp, tag=""):
    """存新指纹，同时和上一版比一次。返回 (path, diffs)。

    第一次跑（没有旧指纹）时 diffs 为空，不算漂移。
    """
    p = path_for(data_dir, subject, tag)
    old = load(p)
    diffs = diff(old, fp) if old else []
    save(fp, p)
    return p, diffs
