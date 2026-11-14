# -*- coding: utf-8 -*-
"""按会话 jsonl 统计真实 token：用于对比「同一个人工任务，配置复用前后差多少」。

    # 跑任务前打快照
    token_report.py snapshot --out tmp/复跑_xxx/out/token_snapshot.json
    # 跑任务后出对比
    token_report.py report --date 2026-11-11 --window 20:00-23:59 \
        --snapshot tmp/复跑_xxx/out/token_snapshot.json --out tmp/复跑_xxx/out/token.json

口径要点（踩过的坑，别改）：
  · 请求键是 providerData.conversationRequestId，**没有**顶层 requestId
  · 同一步的 usage 会在 function_call 与 model-usage 两条记录里重复出现，
    必须按 (rid, stepSeq) 去重后求和，逐行累加会翻倍
  · usage.totalTokens 是「每一步的上下文总量」，一轮 80 步累加能到几千万；
    它是"流量"不是"花费"，报告里要写清
  · 会话文件是跨天的长会话，一个文件 ≠ 一次任务；基线必须同时给
    A 严格（窗口内 + 用户消息含关键词）/ B 上界（窗口内全部）/ C 人工剔除
"""
import argparse
import json
import os
import re
import sys
import time

SESS_DIR = os.path.join(os.path.expanduser("~"), ".workbuddy", "projects",
                        "d-桌面-FDE工作-01_PTW框架")


def files_of(d):
    out = []
    for f in sorted(os.listdir(d)):
        if f.endswith(".jsonl"):
            p = os.path.join(d, f)
            out.append((p, os.path.getsize(p), os.path.getmtime(p)))
    return out


def text_of(d):
    """取用户消息的文本内容。"""
    parts = []
    m = d.get("message") or {}
    for src in (d.get("content"), m.get("content")):
        if isinstance(src, str):
            parts.append(src)
        elif isinstance(src, list):
            for c in src:
                if isinstance(c, dict) and c.get("text"):
                    parts.append(str(c["text"]))
                elif isinstance(c, str):
                    parts.append(c)
    return " ".join(parts)


def scan(paths, offsets=None):
    """扫 jsonl，返回 {rid: {...}}。offsets 给定时只从偏移处读。"""
    rids = {}
    for p in paths:
        start = 0
        if offsets is not None:
            start = offsets.get(os.path.basename(p), 0)
            if start > os.path.getsize(p):
                start = 0                      # 文件被重写/截断 → 全量重扫
        with open(p, "rb") as fh:
            fh.seek(start)
            for raw in fh:
                try:
                    d = json.loads(raw)
                except Exception:
                    continue
                pd = d.get("providerData") or {}
                rid = pd.get("conversationRequestId") or d.get("requestId")
                if not rid:
                    continue
                r = rids.setdefault(rid, {"steps": {}, "tools": 0, "texts": [], "ts": []})
                ts = d.get("timestamp") or pd.get("timestamp")
                if isinstance(ts, (int, float)):
                    r["ts"].append(int(ts))
                if d.get("type") == "function_call":
                    r["tools"] += 1
                if d.get("type") == "message" and (
                        d.get("role") == "user" or (d.get("message") or {}).get("role") == "user"):
                    t = text_of(d)
                    if t:
                        r["texts"].append(t)
                u = pd.get("usage") or (d.get("message") or {}).get("usage") or {}
                st = pd.get("stepSeq")
                tot = u.get("totalTokens") or u.get("total_tokens")
                if st is not None and tot:
                    r["steps"][st] = max(int(r["steps"].get(st, 0)), int(tot))
    for r in rids.values():
        r["tokens"] = sum(r["steps"].values())
        r["n_steps"] = len(r["steps"])
        r["first_ts"] = min(r["ts"]) if r["ts"] else None
        r["last_ts"] = max(r["ts"]) if r["ts"] else None
        r["peak_ctx"] = max(r["steps"].values()) if r["steps"] else 0
        # 关键词匹配要用全文：前 400 字基本都是 system-reminder，截断会把用户原话切掉
        r["text"] = " ".join(r["texts"])[:20000]
    return rids


def in_window(ts, day, wfrom, wto):
    if not ts:
        return False
    t = time.localtime(ts / 1000.0 if ts > 1e11 else ts)
    if "%04d-%02d-%02d" % (t.tm_year, t.tm_mon, t.tm_mday) != day:
        return False
    hm = "%02d:%02d" % (t.tm_hour, t.tm_min)
    return (wfrom or "00:00") <= hm <= (wto or "23:59")


def agg(rids, keys):
    keys = list(keys)
    reqs = len(keys)
    steps = sum(rids[k]["n_steps"] for k in keys)
    tokens = sum(rids[k]["tokens"] for k in keys)
    tools = sum(rids[k]["tools"] for k in keys)
    peak = max((rids[k]["peak_ctx"] for k in keys), default=0)
    return {"requests": reqs, "steps": steps, "tokens": tokens, "tools": tools, "peak_ctx": peak}


def fmt(n):
    return "{:,}".format(int(n))


def pct(base, now):
    if not base:
        return "n/a"
    return "%.1f%%" % (100.0 * now / base) if now else "0.0%"


def print_table(title, rows):
    print("\n" + title)
    print("%-22s %8s %8s %16s %8s %12s" % ("口径", "请求", "步数", "totalTokens", "工具调用", "单步峰值"))
    for name, a in rows:
        print("%-22s %8s %8s %16s %8s %12s"
              % (name, fmt(a["requests"]), fmt(a["steps"]), fmt(a["tokens"]),
                 fmt(a["tools"]), fmt(a["peak_ctx"])))


def cmd_snapshot(args):
    fs = files_of(args.sessions)
    snap = {"sessions_dir": args.sessions, "taken_at_ms": int(time.time() * 1000),
            "files": {os.path.basename(p): sz for p, sz, _ in fs}}
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(snap, fh, ensure_ascii=False, indent=2)
    print("快照写入 %s（%d 个会话文件，共 %.1f MB）"
          % (args.out, len(fs), sum(s for _, s, _ in fs) / 1048576.0))


def cmd_candidates(args):
    rids = scan([p for p, _, _ in files_of(args.sessions)])
    rows = [(k, v) for k, v in rids.items() if in_window(v["first_ts"], args.date, args.wfrom, args.wto)]
    rows.sort(key=lambda x: x[1]["first_ts"] or 0)
    print("%-19s %-14s %9s %6s %6s  %s"
          % ("时间", "rid", "tokens", "步数", "工具", "首句"))
    for k, v in rows:
        t = time.strftime("%m-%d %H:%M:%S", time.localtime((v["first_ts"] or 0) / 1000.0))
        print("%-19s %-14s %9s %6d %6d  %s"
              % (t, k[-14:], fmt(v["tokens"]), v["n_steps"], v["tools"],
                 re.sub(r"\s+", " ", v["text"])[200:290]))


def cmd_report(args):
    paths = [p for p, _, _ in files_of(args.sessions)]
    rids = scan(paths)
    kws = [k.strip().lower() for k in (args.keywords or "").split(",") if k.strip()]

    def hit(v):
        t = (v["text"] or "").lower()
        return any(k in t for k in kws)

    win = [k for k, v in rids.items() if in_window(v["first_ts"], args.date, args.wfrom, args.wto)]
    strict = [k for k in win if (not kws or hit(rids[k]))]
    excl = set(x.strip() for x in (args.exclude_rids or "").split(",") if x.strip())
    manual = [k for k in strict if k not in excl]

    out = {"sessions_dir": args.sessions, "date": args.date,
           "window": [args.wfrom, args.wto], "keywords": kws,
           "baseline": {"A_strict": agg(rids, strict),
                        "B_window_all": agg(rids, win),
                        "C_manual": agg(rids, manual),
                        "excluded_rids": sorted(excl)},
           "this_run": None, "mode": None}

    print_table("基线（%s %s–%s）" % (args.date, args.wfrom, args.wto),
                [("A 严格（含关键词）", out["baseline"]["A_strict"]),
                 ("B 上界（窗口内全部）", out["baseline"]["B_window_all"]),
                 ("C 人工剔除后", out["baseline"]["C_manual"])])

    if args.snapshot and os.path.isfile(args.snapshot):
        snap = json.load(open(args.snapshot, encoding="utf-8"))
        t0 = snap.get("taken_at_ms", 0)
        inc = scan(paths, offsets=snap.get("files") or {})
        now_ids = [k for k, v in inc.items() if (v["first_ts"] or 0) >= t0 - 5000]
        mode = "append"
        if not now_ids:                      # 兜底：会话轮转/时间戳异常 → 用全量里首末时间窗
            inc2 = {k: v for k, v in rids.items() if (v["first_ts"] or 0) >= t0 - 5000}
            now_ids = list(inc2)
            mode = "time-filtered"
        cur = agg(inc, now_ids)
        out["this_run"] = cur
        out["mode"] = mode
        base = out["baseline"]["C_manual"] if excl else out["baseline"]["A_strict"]
        print_table("本次（%s 之后）" % time.strftime("%m-%d %H:%M:%S", time.localtime(t0 / 1000.0)),
                    [("本次", cur)])
        print("\n%-22s %16s %16s %10s" % ("对比", "基线", "本次", "占比"))
        for lab, b in (("A 严格", out["baseline"]["A_strict"]),
                       ("B 上界", out["baseline"]["B_window_all"]),
                       ("C 人工", out["baseline"]["C_manual"])):
            print("%-22s %16s %16s %10s"
                  % (lab, fmt(b["tokens"]), fmt(cur["tokens"]), pct(b["tokens"], cur["tokens"])))
        print("降幅（对 A）：%.1f%%    降幅（对 C）：%.1f%%"
              % (100 - 100.0 * cur["tokens"] / max(1, out["baseline"]["A_strict"]["tokens"]),
                 100 - 100.0 * cur["tokens"] / max(1, out["baseline"]["C_manual"]["tokens"])))
        print("步数：基线 A %s → 本次 %s    工具调用：基线 A %s → 本次 %s"
              % (fmt(out["baseline"]["A_strict"]["steps"]), fmt(cur["steps"]),
                 fmt(out["baseline"]["A_strict"]["tools"]), fmt(cur["tools"])))
        print("（mode=%s；totalTokens 是整轮累加、含 cache_read，是流量不是花费）" % mode)

        # 刷新快照，支持连续多次复跑
        snap["files"] = {os.path.basename(p): os.path.getsize(p) for p in paths}
        snap["taken_at_ms"] = int(time.time() * 1000)
        json.dump(snap, open(args.snapshot, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        json.dump(out, open(args.out, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print("\n已写入 %s" % args.out)


def main():
    ap = argparse.ArgumentParser(description="会话 token 统计与基线对比")
    ap.add_argument("--sessions", default=SESS_DIR)
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("snapshot", help="跑任务前打快照")
    s.add_argument("--sessions", default=SESS_DIR)
    s.add_argument("--out", required=True)
    s.set_defaults(func=cmd_snapshot)

    c = sub.add_parser("candidates", help="列出某天窗口内的请求，人工挑要剔除的")
    c.add_argument("--sessions", default=SESS_DIR)
    c.add_argument("--date", default=time.strftime("%Y-%m-%d"))
    c.add_argument("--window", default="00:00-23:59")
    c.set_defaults(func=cmd_candidates, wfrom=None, wto=None)

    r = sub.add_parser("report", help="出对比报告")
    r.add_argument("--sessions", default=SESS_DIR)
    r.add_argument("--date", default="2026-11-11")
    r.add_argument("--window", default="20:00-23:59")
    r.add_argument("--keywords", default="海柔,hairou,HAI ROBOTICS")
    r.add_argument("--snapshot", default="")
    r.add_argument("--exclude-rids", default="")
    r.add_argument("--out", default="")
    r.set_defaults(func=cmd_report, wfrom=None, wto=None)

    args = ap.parse_args()
    if getattr(args, "window", None):
        a, _, b = args.window.partition("-")
        args.wfrom, args.wto = a.strip() or "00:00", b.strip() or "23:59"
    args.func(args)


if __name__ == "__main__":
    main()
