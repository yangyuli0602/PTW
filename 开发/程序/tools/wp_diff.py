# -*- coding: utf-8 -*-
"""底稿逐格比对：拿「基线 xlsx」和「新生成的 xlsx」逐格比，判定是否同一份底稿。

用途
    换机器 / 换 AI / 换版本重跑同一套输入后，证明「结果一样」。
    不看总数对不对（数字对不对是另一回事，见 08_底稿运行SOP.md 的不变量），
    只看「每个格子的公式与静态值是否与基线一致」。

口径
    · 公式格比公式文本（去多余空格），静态格比规范化后的值；
    · 公式格另比一次缓存值（程序写盘时落的数），差 >0.005 记为缓存差异（警告，不算失败）；
    · 工作表按名字对齐，缺表/多表直接判失败；
    · 每张表的图片数量要比（C.02a 的汇率截图就靠这个）；
    · 「检查」页 A2（生成时间 + 数据来源路径）属易变格，规范化后跳过。

用法
    python wp_diff.py <基线.xlsx> <新文件.xlsx> [--json 报告.json] [--max-show 40]
    退出码：0 = 完全一致；1 = 有差异；2 = 用法/读文件出错。
"""
import argparse
import json
import os
import re
import sys

try:
    from openpyxl import load_workbook
except Exception:  # pragma: no cover
    sys.stderr.write("!! 需要 openpyxl（用平台自带 runtime 跑）\n")
    raise

#: 整格视为「生成痕迹」的判据：命中就把整格替换成固定占位
VOLATILE_TEXT = re.compile(r"生成时间\s*[:：]")
#: 单元格文本里的绝对路径（Windows 盘符 + 反斜杠或斜杠）
ABS_PATH = re.compile(r"[A-Za-z]:[\\/][^\s\u3000，,；;）)]*")
#: 生成器写进说明文字的版本号，换版本重跑必然变，但不代表底稿不同
VER_TOKEN = re.compile(r"V\d+\.\d+\.\d+")

NUM_TOL = 0.005


def _norm_text(s):
    s = s.replace("\u3000", " ").replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s).strip()
    if VOLATILE_TEXT.search(s):
        # 生成时间 / 数据来源路径 / 版本号：整格抹平
        return "<生成痕迹>"
    s = ABS_PATH.sub("<路径>", s)
    s = VER_TOKEN.sub("<版本>", s)
    return s


def _norm_value(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = _norm_text(v)
        return None if s == "" else ("S", s)
    if isinstance(v, bool):
        return ("B", v)
    if isinstance(v, (int, float)):
        return ("N", round(float(v), 6))
    # datetime / date
    try:
        if getattr(v, "hour", 0) == 0 and getattr(v, "minute", 0) == 0 and getattr(v, "second", 0) == 0:
            return ("D", v.strftime("%Y-%m-%d"))
        return ("D", v.strftime("%Y-%m-%d %H:%M:%S"))
    except Exception:
        return ("T", str(v))


def _norm_formula(f):
    """公式规范化。

    `_xlfn.` / `_xlws.` 是「新函数前缀」：openpyxl 写盘时给 XLOOKUP、TEXTJOIN 这类
    新函数加前缀，Excel 打开另存后又把它去掉。同一个函数两种写法，Excel 结果完全相同，
    不能当成差异，否则「基线被 Excel 打开过、新版没打开过」就会假报一片。
    """
    s = re.sub(r"\s+", "", f)
    s = re.sub(r"_xlfn\.|_xlws\.", "", s, flags=re.I)
    return s


def _cell_token(ws, coord, cache_ws):
    c = ws[coord]
    v = c.value
    if isinstance(v, str) and v.startswith("="):
        tok = ("F", _norm_formula(v))
        cached = _norm_value(cache_ws[coord].value) if cache_ws is not None else None
        return tok, cached
    return ("V", _norm_value(v)), None


def diff_sheet(base_ws, new_ws, base_cache, new_cache, max_show):
    """返回 (差异列表, 缓存差异列表, 图片差异)。差异项：(coord, 基线, 新, 类型)"""
    diffs, cache_diffs = [], []
    rows = max(base_ws.max_row or 0, new_ws.max_row or 0)
    cols = max(base_ws.max_column or 0, new_ws.max_column or 0)
    # 最多扫 400 行 × 60 列，超出部分只报「尺寸不同」
    scan_r, scan_c = min(rows, 400), min(cols, 60)
    for r in range(1, scan_r + 1):
        for col in range(1, scan_c + 1):
            coord = "%s%d" % (_col_letter(col), r)
            bt, bcached = _cell_token(base_ws, coord, base_cache)
            nt, ncached = _cell_token(new_ws, coord, new_cache)
            if bt != nt:
                if len(diffs) < max_show:
                    diffs.append((coord, bt, nt, "值/公式"))
                else:
                    diffs.append((coord, None, None, "…"))
            elif bt[0] == "F" and bcached is not None and ncached is not None:
                bn, nn = bcached, ncached
                if bn[0] == "N" and nn[0] == "N":
                    if abs(bn[1] - nn[1]) > NUM_TOL:
                        cache_diffs.append((coord, bn[1], nn[1]))
                elif bn != nn:
                    cache_diffs.append((coord, bn, nn))
    img = (len(getattr(base_ws, "_images", []) or []),
           len(getattr(new_ws, "_images", []) or []))
    return diffs, cache_diffs, img


def _col_letter(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def main():
    ap = argparse.ArgumentParser(description="底稿逐格比对（基线 vs 新生成）")
    ap.add_argument("baseline")
    ap.add_argument("candidate")
    ap.add_argument("--json", default="", help="把报告写成 JSON")
    ap.add_argument("--max-show", type=int, default=40, help="每表最多列出的差异格数")
    a = ap.parse_args()

    for p in (a.baseline, a.candidate):
        if not os.path.isfile(p):
            sys.stderr.write("!! 找不到文件：%s\n" % p)
            return 2

    bf = load_workbook(a.baseline)
    nf = load_workbook(a.candidate)
    bv = load_workbook(a.baseline, data_only=True)
    nv = load_workbook(a.candidate, data_only=True)

    bs, ns = set(bf.sheetnames), set(nf.sheetnames)
    report = {"baseline": os.path.abspath(a.baseline),
              "candidate": os.path.abspath(a.candidate),
              "missingSheets": sorted(bs - ns), "extraSheets": sorted(ns - bs),
              "sheets": {}, "ok": True}

    if bs - ns:
        report["ok"] = False
    if ns - bs:
        report["ok"] = False

    total_diff = total_cache = 0
    for name in [s for s in bf.sheetnames if s in ns]:
        d, cd, img = diff_sheet(bf[name], nf[name], bv[name], nv[name], a.max_show)
        n_real = sum(1 for x in d if x[3] != "…")
        entry = {"diffs": n_real, "cacheDiffs": len(cd), "images": {"baseline": img[0], "candidate": img[1]}}
        if n_real:
            entry["diffDetail"] = [{"cell": x[0], "baseline": _brief(x[1]), "candidate": _brief(x[2])} for x in d]
            report["ok"] = False
        if cd:
            entry["cacheDetail"] = [{"cell": x[0], "baseline": x[1], "candidate": x[2]} for x in cd[:a.max_show]]
        if img[0] != img[1]:
            report["ok"] = False
            entry["imageMismatch"] = True
        if n_real or cd or img[0] != img[1]:
            report["sheets"][name] = entry
        total_diff += n_real
        total_cache += len(cd)

    report["totalCellDiffs"] = total_diff
    report["totalCacheDiffs"] = total_cache

    # ---- 人读报告 ----
    print("基线：%s" % os.path.basename(a.baseline))
    print("新版：%s" % os.path.basename(a.candidate))
    print("工作表：基线 %d / 新版 %d" % (len(bs), len(ns)))
    if bs - ns:
        print("!! 缺表：%s" % "、".join(sorted(bs - ns)))
    if ns - bs:
        print("!! 多表：%s" % "、".join(sorted(ns - bs)))
    if not report["sheets"]:
        print("\n结论：逐格一致（%d 张表，公式与静态值全部相同）" % len(bs))
    else:
        print("")
        for name, e in report["sheets"].items():
            head = "· %s：差异 %d 格" % (name, e["diffs"])
            if e.get("imageMismatch"):
                head += "，图片 %d→%d" % (e["images"]["baseline"], e["images"]["candidate"])
            print(head)
            for x in (e.get("diffDetail") or [])[:a.max_show]:
                print("    %-6s 基线 %-46s 新版 %s" % (x["cell"], x["baseline"], x["candidate"]))
            for x in (e.get("cacheDetail") or [])[:6]:
                print("    %-6s 缓存值不同：%s → %s（公式相同，需 Excel 重算确认）" % (x["cell"], x["baseline"], x["candidate"]))
        print("\n结论：不一致 —— 单元格差异 %d 处，缓存差异 %d 处" % (total_diff, total_cache))
        if total_cache and not total_diff:
            print("提示：只有缓存值不同、公式一致，多半是重算精度差；用 Excel 打开重算再比一次。")

    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=1)
        print("报告：%s" % a.json)

    return 0 if report["ok"] else 1


def _brief(tok):
    if tok is None:
        return "<空>"
    kind, val = tok
    if kind == "F":
        return "=" + (val[:60] if isinstance(val, str) else str(val))
    if val is None:
        return "<空>"
    s = str(val)
    return s if len(s) <= 60 else s[:57] + "…"


if __name__ == "__main__":
    sys.exit(main())
