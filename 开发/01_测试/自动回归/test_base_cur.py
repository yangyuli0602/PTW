# -*- coding: utf-8 -*-
"""清洗步骤「删除综合本位币行」（V1.2.96）回归。

**为什么要有这个测试**：多币别账套的余额表里，「综合本位币」汇总行与「人民币」
「美元」这些具体币别行并存，两份都算会把同一笔钱算两遍（瑞云实测 14.23 亿会变成
28.46 亿）。而这一行**看不出具体币别**，对原币核对没有任何用处。所以清洗阶段就删，
下游不用每处再判一次。

三条防线：
1. 双视图场景确实删掉了汇总行、具体币别行一行不少；
2. 整表都是汇总行时不删（否则把表清空，比翻倍更糟）；
3. 没有币别列时不报错、不误删。

跑法：
    程序/runtime/python/python.exe 01_测试/自动回归/test_base_cur.py
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "程序"))

import duckdb                                          # noqa: E402

from core import clean                                 # noqa: E402

FAILS = []


def ck(name, cond, extra=""):
    if cond:
        print("  \u2713 %s" % name)
    else:
        print("  \u2717 %s  %s" % (name, extra))
        FAILS.append(name)


def mk(rows, cols_sql):
    con = duckdb.connect()
    con.execute(cols_sql)
    for i, r in enumerate(rows, 1):
        con.execute("INSERT INTO t VALUES (%s)"
                    % ",".join(["?"] * (len(r) + 1)), [i] + list(r))
    return con


COLS6 = ('''CREATE TABLE t(_src_row BIGINT, company VARCHAR, acct_code VARCHAR,
            acct_name VARCHAR, currency VARCHAR, dr VARCHAR, cr VARCHAR)''')


def main():
    print("=== 1. 双视图：综合本位币 + 人民币 + 美元 ===")
    con = mk([
        ("甲公司", "1002", "银行存款", "综合本位币", "800", "0"),
        ("甲公司", "1002", "银行存款", "人民币",     "700", "0"),
        ("甲公司", "1002", "银行存款", "美元",       "100", "0"),
        ("甲公司", "6602", "管理费用", "综合本位币", "0", "800"),
        ("甲公司", "6602", "管理费用", "人民币",     "0", "700"),
        ("甲公司", "6602", "管理费用", "美元",       "0", "100"),
    ], COLS6)
    cols = clean.columns_of(con, "t")
    steps = clean.preset("generic", cols, "tb")
    ck("科目表预设自带 drop_base_cur",
       "drop_base_cur" in [s["type"] for s in steps],
       [s["type"] for s in steps])

    clean.run(con, "t1", "t", steps, keep=True, sample=0)
    got = dict(con.execute(
        "SELECT currency, count(*) FROM stg_t1 GROUP BY 1").fetchall())
    ck("汇总行全删", "综合本位币" not in got, got)
    ck("人民币行一行不少", got.get("人民币") == 2, got)
    ck("美元行一行不少", got.get("美元") == 2, got)
    reasons = [str(r[0]) for r in con.execute(
        'SELECT DISTINCT reason FROM "_dropped_t1"').fetchall()]
    ck("剔除留痕写明原因", any("综合本位币" in x for x in reasons), reasons)
    con.close()

    print()
    print("=== 2. 整表都是汇总行 → 不能删空 ===")
    con = mk([
        ("甲公司", "1002", "银行存款", "综合本位币", "800", "0"),
        ("甲公司", "6602", "管理费用", "综合本位币", "0", "800"),
    ], COLS6)
    cols = clean.columns_of(con, "t")
    res = clean.run(con, "t2", "t", clean.preset("generic", cols, "tb"),
                    keep=True, sample=0)
    n = con.execute("SELECT count(*) FROM stg_t2").fetchone()[0]
    ck("保留不清空", n == 2, n)
    ck("步骤记录为跳过并说明原因",
       any("保留" in str(s.get("note") or "") for s in res["steps"]),
       [s.get("note") for s in res["steps"] if s.get("note")])
    con.close()

    print()
    print("=== 3. 没有币别列 → 跳过、不报错、不误删 ===")
    con = mk([("甲公司", "1002", "100")],
             '''CREATE TABLE t(_src_row BIGINT, company VARCHAR,
                acct_code VARCHAR, dr VARCHAR)''')
    cols = clean.columns_of(con, "t")
    res = clean.run(con, "t3", "t", clean.preset("generic", cols, "tb"),
                    keep=True, sample=0)
    ck("不报错", not res.get("error"), res.get("error"))
    ck("数据还在", con.execute("SELECT count(*) FROM stg_t3").fetchone()[0] == 1)
    con.close()

    print()
    if FAILS:
        print("失败 %d 项：%s" % (len(FAILS), FAILS))
        return 1
    print("全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
