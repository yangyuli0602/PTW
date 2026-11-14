# -*- coding: utf-8 -*-
"""PBC 输入指纹：跑底稿前先证明「喂进去的是同一份数据、同一份上年底稿」。

为什么需要
    实测两套账跑出来不一样，多数不是 AI 做得不对，而是输入不同（清洗口径被改过、
    上年底稿换了版本、期间错位）。光看文件名看不出来，必须看指纹。

用法
    生成：python input_fingerprint.py --data-dir "<数据目录>" [--prior "上年底稿.xlsx"] --out 指纹.json
    核对：python input_fingerprint.py --data-dir "<数据目录>" [--prior "..."] --check 指纹.json
          退出码 0 = 逐项一致；1 = 有差异；2 = 读不了

口径
    · 行数与金额一律走平台自己的取数函数（`core.read_tb_from_duckdb`），
      保证指纹与底稿同源 —— 自己另写一段 SQL 会得出「看起来差不多」的第二个口径。
    · 上年底稿解析失败不抛错，记 `error`，指纹照样能比（说明这份文件本身要先修）。
"""
import argparse
import hashlib
import io
import json
import os
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


def sha256(path, chunk=1 << 20):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def _platform_version(root):
    p = os.path.join(root, "server.py")
    try:
        with open(p, encoding="utf-8") as f:
            for line in f:
                if line.startswith("VERSION"):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    except Exception:
        pass
    return ""


def build(data_dir, prior_path):
    root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    sys.path.insert(0, root)
    try:
        import duckdb
        from modules.bkd import core as BKD
    except Exception as e:
        raise SystemExit("!! 取数依赖不可用：%s" % e)

    db = os.path.join(data_dir, "ledger.duckdb")
    if not os.path.isfile(db):
        raise SystemExit("!! 数据目录里找不到 ledger.duckdb：%s" % db)

    fp = {"tool": {"name": "input_fingerprint",
                   "platform": _platform_version(root),
                   "at": time.strftime("%Y-%m-%d %H:%M:%S")}}
    fp["db"] = {"file": "ledger.duckdb", "bytes": os.path.getsize(db), "sha256": sha256(db)}

    try:
        con = duckdb.connect(db, read_only=True)
    except Exception as e:
        raise SystemExit("!! 打不开库（多半被别的进程独占，先关掉再跑）：%s" % e)
    try:
        tabs = [r[0] for r in con.execute("SHOW TABLES").fetchall()]
        for t in ("tb", "journal"):
            if t in tabs:
                fp["db"]["%s_rows" % t] = con.execute("SELECT COUNT(*) FROM %s" % t).fetchone()[0]
            else:
                fp["db"]["%s_rows" % t] = None
        rows, meta = BKD.read_tb_from_duckdb("货币资金", conn=con)
        if meta.get("error"):
            fp["cash"] = {"error": meta["error"]}
        else:
            total = sum(float(r.get("local") or 0) for r in rows)
            comps = sorted({str(r.get("company") or "") for r in rows if r.get("company")})
            fp["cash"] = {"rows": len(rows), "local_sum": round(total, 2),
                          "companies": len(comps),
                          "companies_sample": comps[:3]}
        if "journal" in tabs:
            try:
                lo, hi = con.execute("SELECT MIN(period), MAX(period) FROM journal").fetchone()
                fp["cash"]["period_from"] = lo
                fp["cash"]["period_to"] = hi
            except Exception:
                pass
    finally:
        con.close()

    fp["prior"] = {"file": "", "error": ""}
    if prior_path:
        if not os.path.isfile(prior_path):
            fp["prior"]["error"] = "文件不存在：%s" % prior_path
        else:
            fp["prior"]["file"] = os.path.basename(prior_path)
            fp["prior"]["bytes"] = os.path.getsize(prior_path)
            fp["prior"]["sha256"] = sha256(prior_path)
            try:
                raw = open(prior_path, "rb").read()
                wb = BKD.load_wb_any(raw, data_only=True)
                best = (BKD.probe_prior_structure(wb).get("best") or {})
                if best:
                    prior, _ = BKD.parse_prior_xlsx(raw, [{"sheet": best["name"],
                                                           "header_row": best["header_row"]}])
                    fp["prior"]["sheet"] = best["name"]
                    fp["prior"]["rows"] = len(prior)
                    fp["prior"]["total"] = round(sum(float(p.get("local") or 0) for p in prior), 2)
                else:
                    fp["prior"]["error"] = "认不出上年底稿结构（找不到 BKD 表头）"
            except Exception as e:
                fp["prior"]["error"] = "%s: %s" % (type(e).__name__, str(e)[:120])
    return fp


def compare(cur, base):
    """逐项比 db 与 cash（prior 只在两边都有 sha256 时比）。返回差异列表。"""
    diffs = []

    def walk(a, b, path):
        if isinstance(b, dict) and isinstance(a, dict):
            for k in b:
                if k in ("at", "file", "platform", "companies_sample", "error"):
                    continue
                walk(a.get(k), b[k], path + [k])
        else:
            if a != b:
                diffs.append(("%s.%s" % (path[0], ".".join(path[1:])), b, a))

    walk(cur.get("db", {}), base.get("db", {}), ["db"])
    walk(cur.get("cash", {}), base.get("cash", {}), ["cash"])
    bp, cp = base.get("prior") or {}, cur.get("prior") or {}
    if bp.get("sha256") and cp.get("sha256"):
        if bp["sha256"] != cp["sha256"]:
            diffs.append(("prior.sha256", bp["sha256"][:12], cp["sha256"][:12]))
    elif bp.get("sha256") and not cp.get("sha256"):
        diffs.append(("prior.sha256", "有基线", "本次未提供 --prior"))
    return diffs


def main():
    ap = argparse.ArgumentParser(description="PBC 输入指纹（生成 / 核对）")
    ap.add_argument("--data-dir", required=True, help="含 ledger.duckdb 的数据目录")
    ap.add_argument("--prior", default="", help="上年底稿路径（可选）")
    ap.add_argument("--out", default="", help="写指纹 json 的路径")
    ap.add_argument("--check", default="", help="与已有指纹 json 比对")
    a = ap.parse_args()

    fp = build(a.data_dir, a.prior)

    if a.check:
        if not os.path.isfile(a.check):
            sys.stderr.write("!! 找不到基线指纹：%s\n" % a.check)
            return 2
        base = json.load(open(a.check, encoding="utf-8"))
        diffs = compare(fp, base)
        print("基线指纹：%s（平台 %s）" % (os.path.basename(a.check),
                                     (base.get("tool") or {}).get("platform") or "?"))
        print("本次：平台 %s" % fp["tool"]["platform"])
        if not diffs:
            print("\n结论：输入一致（行数、金额、期间、上年底稿指纹全部相同）")
            return 0
        print("\n!! 输入不一致 %d 处：" % len(diffs))
        for k, b, c in diffs:
            print("   %-22s 基线 %-24s 本次 %s" % (k, b, c))
        print("\n先确认是换了数据还是拿错了文件；换了就人工核过后重新生成指纹，别直接覆盖。")
        return 1

    if a.out:
        os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
        json.dump(fp, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("指纹已写入：%s" % a.out)
    print(json.dumps(fp, ensure_ascii=False, indent=1))
    if fp.get("cash", {}).get("error"):
        print("\n!! 取数失败：%s" % fp["cash"]["error"])
        return 2
    if fp.get("prior", {}).get("error"):
        print("\n!! 上年底稿有问题：%s（先用 Excel 打开另存修一遍）" % fp["prior"]["error"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
