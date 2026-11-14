# -*- coding: utf-8 -*-
"""按科目编码派生一级科目（clean.derive_l1）的边界检查。

覆盖：位数推断（只有 8 位明细也能推成 4）/ 分隔符编码 / 名称三级优先级 /
两种兜底与 note 计数 / 三种 noop 原因 / 列序与行数不变 / 重名防撞 / 整条流水线 / 覆盖文件。
"""
import json
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.abspath(__file__))
while not os.path.isdir(os.path.join(ROOT, "程序")):   # 往上找项目根，别靠层数猜
    parent = os.path.dirname(ROOT)
    if parent == ROOT:
        break
    ROOT = parent
sys.path.insert(0, os.path.join(ROOT, "程序"))

import duckdb
from core import clean, std_accounts

fails = []
_checks = [0]


def ck(name, cond, got=None):
    _checks[0] += 1
    print(("  ok   " if cond else "  FAIL ") + name + ("" if cond else " | %r" % (got,)))
    if not cond:
        fails.append(name)


def mk(rows, cols=("_src_row", "code", "name", "amt")):
    """建一张最小原始表：_src_row / code / name / amt。"""
    c = duckdb.connect()
    c.execute('CREATE TABLE t(_src_row BIGINT, code VARCHAR, name VARCHAR, amt VARCHAR)')
    for r in rows:
        c.execute("INSERT INTO t VALUES (?,?,?,?)", list(r))
    return c, list(cols)


def run_plan(c, plan):
    return c.execute(plan["sql"].replace("{t}", '"t"')).fetchall()


STD = std_accounts.load(None)

print("=== 1. 8 位明细 + 4 位父行混存，手工指 4 位 ===")
c, cols = mk([(1, "1002", "银行存款", "0"), (2, "10020101", "工行孝陵卫支行", "1")])
plan = clean._b_derive_l1({"col": "code", "name_col": "name", "length": 4,
                           "_conn": c, "_std": STD}, cols, "t")
rows = run_plan(c, plan)
ck("取到 2 行（不删行）", len(rows) == 2, len(rows))
ck("父行 1002 → 银行存款", rows[0][4:] == ("1002", "银行存款"), rows[0])
ck("明细 10020101 → 1002 银行存款", rows[1][4:] == ("1002", "银行存款"), rows[1])

print("=== 2. 只有 8 位明细行，位数自动推断 ===")
c, cols = mk([(1, "10020101", "工行", "1"), (2, "660201", "办公费", "2"),
              (3, "11220101", "客户甲", "3")])
n, why = clean._auto_l1_len(c, "t", "code", STD)
ck("推断出 4 位（不是 8）", n == 4, (n, why))
plan = clean._b_derive_l1({"col": "code", "name_col": "name", "length": 0,
                           "_auto_l1_len": n, "_conn": c, "_std": STD}, cols, "t")
rows = run_plan(c, plan)
ck("660201 → 6602 管理费用", rows[1][4:] == ("6602", "管理费用"), rows[1])

print("=== 3. 分隔符 / 空格编码也能切 ===")
c, cols = mk([(1, "1002.01.01", "a", "1"), (2, "1122-01", "b", "2"),
              (3, "1002 银行存款", "c", "3")])
plan = clean._b_derive_l1({"col": "code", "length": 4, "_conn": c, "_std": STD}, cols, "t")
got = [r[4] for r in run_plan(c, plan)]
ck("点分 / 横线 / 空格都取前导数字", got == ["1002", "1122", "1002"], got)

print("=== 4. 名称优先级：标准表 > 本表名称 ===")
c, cols = mk([(1, "6602", "管理费用—旧", "1"), (2, "660201", "办公费", "2")])
plan = clean._b_derive_l1({"col": "code", "name_col": "name", "length": 4,
                           "_conn": c, "_std": STD}, cols, "t")
rows = run_plan(c, plan)
ck("标准表名称压过本表名称", rows[0][5] == "管理费用", rows[0])

print("=== 5. 标准表没有 → 用本表同前缀那一行 ===")
c, cols = mk([(1, "8101", "营业外收入", "1"), (2, "810101", "罚款收入", "2")])
plan = clean._b_derive_l1({"col": "code", "name_col": "name", "length": 4,
                           "_conn": c, "_std": STD}, cols, "t")
rows = run_plan(c, plan)
ck("数据兜底取到名称", rows[1][5] == "营业外收入", rows[1])

print("=== 6. 两级都没有 → 只写编码，并在 note 里报数 ===")
c, cols = mk([(1, "810201", "", "1"), (2, "999901", "", "2")])
plan = clean._b_derive_l1({"col": "code", "name_col": "name", "length": 4,
                           "_conn": c, "_std": STD}, cols, "t")
rows = run_plan(c, plan)
ck("名称回落成编码", rows[0][5] == "8102", rows[0])
ck("note 报出兜底行数", "按编码兜底 2 行" in plan.get("note", ""), plan.get("note"))
ck("note 给出样例编码", "8102" in plan.get("note", ""), plan.get("note"))

print("=== 7. 三种 noop 原因 ===")
c, cols = mk([(1, "1001", "库存现金", "1")])
ck("未选列 → noop", clean._b_derive_l1({"col": "", "_conn": c, "_std": STD}, cols, "t")["kind"] == "noop")
ck("列不在表里 → noop", clean._b_derive_l1({"col": "nope", "_conn": c, "_std": STD}, cols, "t")["kind"] == "noop")
c2, cols2 = mk([(1, "合计", "", "1"), (2, "小计", "", "2")])
p2 = clean._b_derive_l1({"col": "code", "length": 4, "_conn": c2, "_std": STD}, cols2, "t")
ck("没有数字编码 → noop 且有说明", p2["kind"] == "noop" and "没有可用的科目编码" in p2["reason"], p2.get("reason"))
ck("推不出位数 → noop 且有说明",
   clean._b_derive_l1({"col": "code", "length": 0, "_conn": c, "_std": STD}, cols, "t")["kind"] == "noop")

print("=== 8. 列序、行数、原值都不动 ===")
c, cols = mk([(1, "10020101", "工行", "11.5"), (2, "10020201", "中行", "22.5")])
plan = clean._b_derive_l1({"col": "code", "name_col": "name", "length": 4,
                           "_conn": c, "_std": STD}, cols, "t")
rows = run_plan(c, plan)
ck("仍是 2 行", len(rows) == 2, len(rows))
ck("原始四列原样", [r[:4] for r in rows] == [(1, "10020101", "工行", "11.5"),
                                            (2, "10020201", "中行", "22.5")], rows)
ck("新列在最后两列", [r[4] for r in rows] == ["1002", "1002"], rows)

print("=== 9. 新列重名自动让位 ===")
c, cols = mk([(1, "1001", "库存现金", "1")], ["_src_row", "code", "name", "amt", "一级科目"])
plan = clean._b_derive_l1({"col": "code", "name_col": "name", "length": 4,
                           "_conn": c, "_std": STD}, cols, "t")
ck("已有一级科目 → 产出 一级科目2", '"一级科目2"' in plan["sql"], plan["sql"][-80:])

print("=== 10. 走完整 clean.run() 流水线 ===")
c, cols = mk([(1, "10020101", "工行", "1"), (2, "合计", "", "2")])
out = clean.run(c, "sidl1", "t", [{"type": "derive_l1",
                                   "params": {"col": "code", "name_col": "name",
                                              "length": 4, "out_code": "一级代码",
                                              "out_name": "一级科目"}}])
ck("流水线无 error", not out.get("error"), out.get("error"))
ck("步骤 note 带位数", "按前 4 位" in (out["steps"][0].get("note") or ""), out["steps"][0].get("note"))
vals = c.execute('SELECT "一级科目" FROM stg_sidl1 ORDER BY _src_row').fetchall()
ck("落库值正确", [v[0] for v in vals] == ["银行存款", None], vals)

print("=== 11. 客户覆盖文件 data/std_accounts.json ===")
tmpd = tempfile.mkdtemp(prefix="stdacc_")
with open(os.path.join(tmpd, "std_accounts.json"), "w", encoding="utf-8") as f:
    json.dump({"accounts": {"1002": "银行存款(改)", "9999": "测试科目"}}, f, ensure_ascii=False)
m, info = std_accounts.load(tmpd, with_info=True)
ck("同码被覆盖", m.get("1002") == "银行存款(改)", m.get("1002"))
ck("新码被追加", m.get("9999") == "测试科目", m.get("9999"))
ck("未涉及的键保留", m.get("1001") == "库存现金", m.get("1001"))
ck("info 计数正确", info["override"] == 2 and info["total"] == info["builtin"] + 1, info)

print("=== 12. 位数推断的三种结局 ===")
c, _ = mk([(1, "1002", "银行存款", "1"), (2, "10020101", "工行", "2")])
ck("长短码并存 → 4", clean._auto_l1_len(c, "t", "code", STD)[0] == 4)
c2, _ = mk([(1, "99990101", "a", "1"), (2, "88880101", "b", "2")])
n2, why2 = clean._auto_l1_len(c2, "t", "code", {})
ck("统一长度 + 空标准表 → 0 并给原因", n2 == 0 and "手填" in why2, (n2, why2))

print("\n%s  %d 项通过，失败 %d" % (
    "ALL PASS" if not fails else "FAILED", _checks[0] - len(fails), len(fails)))
sys.exit(1 if fails else 0)
