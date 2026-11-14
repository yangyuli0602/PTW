# -*- coding: utf-8 -*-
"""固定资产 SWP · 命令行入口。

独立入口，不改 ``程序/ptw.py``（那里已经挂着货币资金和费用，动它要证明旧输出不变）。

    python -m modules.fa_swp.cli --data-dir tmp/xxx/data --date 2026-09-30

数据目录一律用 ``--data-dir`` 或环境变量 ``KZ_DATA_DIR`` 指过去，
绝不默认指向项目的 ``data/``。
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys

import duckdb

from . import moves


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="fa_swp", description="固定资产 SWP 底稿（K1）生成")
    ap.add_argument("--data-dir", default=os.environ.get("KZ_DATA_DIR", ""),
                    help="数据目录（含 ledger.duckdb 与 bkd_tpl），默认取 KZ_DATA_DIR")
    ap.add_argument("--date", default="2026-09-30", help="资产负债表日")
    ap.add_argument("--customer", default="深圳市瑞云科技股份有限公司（合并 5 主体）")
    ap.add_argument("--te", type=float, default=550000.0, help="可容忍误差")
    ap.add_argument("--sad", type=float, default=55000.0, help="名义金额")
    ap.add_argument("--standard", default="中国企业会计准则")
    ap.add_argument("--currency", default="人民币元")
    ap.add_argument("--seed", default="ruiyun", help="明细合成随机种子（固定则结果可复现）")
    ap.add_argument("--out", default="", help="产物目录，默认 <data>/bkd_out")
    ap.add_argument("--json", action="store_true", help="只打印结果 JSON")
    a = ap.parse_args(argv)

    if not a.data_dir:
        print("必须指定 --data-dir 或设置 KZ_DATA_DIR（不要指向项目 data/）", file=sys.stderr)
        return 2
    data_dir = os.path.abspath(a.data_dir)
    db = os.path.join(data_dir, "ledger.duckdb")
    if not os.path.isfile(db):
        print("数据目录里没有 ledger.duckdb：%s" % db, file=sys.stderr)
        return 2

    out_dir = os.path.abspath(a.out) if a.out else os.path.join(data_dir, "bkd_out")
    os.makedirs(out_dir, exist_ok=True)

    from . import build as B

    info = {"customer": a.customer, "balance_date": a.date, "te": a.te, "sad": a.sad,
            "standard": a.standard, "currency": a.currency}

    conn = duckdb.connect(db, read_only=True)
    try:
        res = B.build(conn, data_dir, out_dir, info,
                      dt.date.fromisoformat(a.date), a.seed)
    finally:
        conn.close()

    tr = res["trace"]
    if a.json:
        print(json.dumps(tr, ensure_ascii=False, indent=1, default=str))
        return 0

    print("底稿：%s" % res["path"])
    print("最新：%s" % res["latest"])
    print("上年K.01：%s" % res["prior"])
    print()
    print("原值  期初 %15s → 期末 %15s" % (tr["totals"]["fa_beg"], tr["totals"]["fa_end"]))
    print("折旧  期初 %15s → 期末 %15s" % (tr["totals"]["dep_beg"], tr["totals"]["dep_end"]))
    print("减值  期初 %15s → 期末 %15s" % (tr["totals"]["imp_beg"], tr["totals"]["imp_end"]))
    print("净值  期初 %15s → 期末 %15s" % (tr["net"]["beg"], tr["net"]["end"]))
    print()
    print("K.01 表1 各方式（行数由数据决定，模板原有方式保留）：")
    for sec, label in (("fa", "原值"), ("dep", "折旧"), ("imp", "减值")):
        parts = ["%s %s" % (w, tr["k01"]["%s|%s" % (sec, w)])
                 for w in moves.WAY_ORDER[sec] if ("%s|%s" % (sec, w)) in tr["k01"]]
        print("  %s  %s" % (label, " / ".join(parts) or "—"))
    print("  插入行：%s" % (tr.get("row_plan") or "无"))
    print("表4 折旧费用：%s" % tr["table4"])
    print("清单行数：%s" % tr["synth"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
