# -*- coding: utf-8 -*-
"""整列取反（clean.negate_col）的边界检查：数字翻号、文本/空/0 不动、列序不变。"""
import os, sys
ROOT = os.path.dirname(os.path.abspath(__file__))
while not os.path.isdir(os.path.join(ROOT, "程序")):   # 往上找项目根，别靠层数猜
    parent = os.path.dirname(ROOT)
    if parent == ROOT:
        break
    ROOT = parent
sys.path.insert(0, os.path.join(ROOT, "程序"))
import duckdb
from core import clean

fails = []
_checks = [0]
def ck(name, cond, got=None):
    _checks[0] += 1
    print(("  ok   " if cond else "  FAIL ") + name + ("" if cond else " | %r" % (got,)))
    if not cond:
        fails.append(name)

def run_mode(mode, rows, col="amt"):
    c = duckdb.connect()
    c.execute("CREATE TABLE t(_src_row BIGINT, amt VARCHAR, memo VARCHAR)")
    for i, (a, m) in enumerate(rows, 1):
        c.execute("INSERT INTO t VALUES (?,?,?)", [i, a, m])
    plan = clean._b_negate_col({"col": col, "mode": mode}, ["_src_row", "amt", "memo"], "t")
    assert plan["kind"] == "select", plan
    sql = plan["sql"].replace("{t}", "t")
    return c.execute(sql + " ORDER BY _src_row").fetchall(), sql

ROWS = [("1,234.00", "a"), ("-98.50", "b"), ("0", "c"), ("合计", "d"), (None, "e"), ("+7", "f")]

print("=== 1. flip：正变负、负变正 ===")
r, sql = run_mode("flip", ROWS)
ck("千分位保留写法并加负号", r[0][1] == "-1,234.00", r[0][1])
ck("负数去符号变正", r[1][1] == "98.50", r[1][1])
ck("0 保持原样", r[2][1] == "0", r[2][1])
ck("文本行不动", r[3][1] == "合计", r[3][1])
ck("空值仍为空", r[4][1] is None, r[4][1])
ck("正号写法变负号", r[5][1] == "-7", r[5][1])

print("=== 2. to_negative：正数取反，负数不动 ===")
r, _ = run_mode("to_negative", ROWS)
ck("正数变负", r[0][1] == "-1,234.00", r[0][1])
ck("负数原样", r[1][1] == "-98.50", r[1][1])
ck("0 原样", r[2][1] == "0", r[2][1])
ck("文本原样", r[3][1] == "合计", r[3][1])

print("=== 3. to_positive：负数取反，正数不动 ===")
r, _ = run_mode("to_positive", ROWS)
ck("正数原样", r[0][1] == "1,234.00", r[0][1])
ck("负数转正", r[1][1] == "98.50", r[1][1])
ck("0 原样", r[2][1] == "0", r[2][1])
ck("文本原样", r[3][1] == "合计", r[3][1])

print("=== 4. 其它列与列序不受影响 ===")
r, sql = run_mode("flip", ROWS)
ck("_src_row 未变", [x[0] for x in r] == [1, 2, 3, 4, 5, 6], [x[0] for x in r])
ck("memo 未变", [x[2] for x in r] == ["a", "b", "c", "d", "e", "f"], [x[2] for x in r])
ck("SQL 里列序保持 _src_row, amt, memo",
   sql.index('"_src_row"') < sql.index('"amt"') < sql.index('"memo"'), sql[:120])

print("=== 5. 未选列 / 列不存在 → noop ===")
ck("空列名 noop", clean._b_negate_col({"col": "", "mode": "flip"}, ["amt"], "t")["kind"] == "noop")
ck("列不存在 noop", clean._b_negate_col({"col": "zzz", "mode": "flip"}, ["amt"], "t")["kind"] == "noop")

print("=== 6. 走完整 run() 流水线 ===")
c = duckdb.connect()
c.execute("CREATE TABLE raw_x(_src_row BIGINT, amt VARCHAR, memo VARCHAR)")
for i, (a, m) in enumerate(ROWS, 1):
    c.execute("INSERT INTO raw_x VALUES (?,?,?)", [i, a, m])
out = clean.run(c, "sidtest", "raw_x", [{"type": "negate_col", "params": {"col": "amt", "mode": "flip"}}])
ck("流水线无 error", not out.get("error"), out.get("error"))
ck("结果列 = amt/memo（_src_row 是内部列，不对外）",
   list(out.get("columns") or []) == ["amt", "memo"], out.get("columns"))
vals = c.execute('SELECT amt FROM stg_sidtest ORDER BY _src_row').fetchall()
ck("落库值已取反", [v[0] for v in vals] == ["-1,234.00", "98.50", "0", "合计", None, "-7"],
   [v[0] for v in vals])

print("\n%s  %d 项通过，失败 %d" % (
    "ALL PASS" if not fails else "FAILED", _checks[0] - len(fails), len(fails)))
sys.exit(1 if fails else 0)
