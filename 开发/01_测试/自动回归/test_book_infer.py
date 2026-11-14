# -*- coding: utf-8 -*-
"""主体对照：账簿名推断 + 第零组「主体完整性」检查（2026-11-11）。

两个回归点：
  A. 账簿对照表是完整主体清单；账簿名对不上时，按名称推断到最可能的公司，
     并留痕（方法 / 分数 / 中文依据 / 是否需确认）。只有「归一后完全相同」能自动认定，
     门店前缀、包含、相似度一律要人工看一眼——「（南海店）」和「（东湖店）」去掉前缀
     是同一串字，但那是两家店。
  B. 第零组检查要看三件事：缺余额表、缺序时账、超范围（表里有、对照表没有）。
"""
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
while not os.path.isdir(os.path.join(ROOT, "程序")):
    parent = os.path.dirname(ROOT)
    if parent == ROOT:
        break
    ROOT = parent
sys.path.insert(0, os.path.join(ROOT, "程序"))

import duckdb
from core.db import _book_key, _pick_book
from core import recon

fails = []
n = [0]


def ck(name, cond, got=None):
    n[0] += 1
    print(("  ok   " if cond else "  FAIL ") + name + ("" if cond else " | %r" % (got,)))
    if not cond:
        fails.append(name)


# 候选池格式：(保守键, 激进键, 公司名, 对照表原账簿名)
def cands_of(pairs):
    out = []
    for bn, cn in pairs:
        for s in {bn, cn}:
            out.append((_book_key(s), _book_key(s, True), cn, bn))
    return out


C = cands_of([
    ("上海万联大药房有限公司(43701)", "上海万联大药房有限公司"),
    ("（万嘉店）上海固生堂万嘉中医门诊有限公司", "上海固生堂万嘉中医门诊有限公司"),
    ("广州蓝海医药有限公司", "广州蓝海医药有限公司"),
])

print("=== A. 账簿名推断 ===")
sc, m, cn, ev, need = _pick_book("上海万联大药房有限公司(43705)", C)
ck("只差尾码 → norm 自动认定", m == "norm" and need == 0 and cn == "上海万联大药房有限公司", (m, cn, need))

sc, m, cn, ev, need = _pick_book("（东湖店）上海固生堂万嘉中医门诊有限公司", C)
ck("同公司、门店不同 → norm_prefix 但要确认", m == "norm_prefix" and need == 1, (m, need))
ck("依据里点名了差异在哪", "门店不同" in ev, ev)

sc, m, cn, ev, need = _pick_book("广州蓝海医药", C)
ck("少「有限公司」→ 归一后相同", m == "norm" and cn == "广州蓝海医药有限公司", (m, cn))

sc, m, cn, ev, need = _pick_book("上海万联大药房有限公司(43701)", C)
ck("原样命中", m == "norm" and cn == "上海万联大药房有限公司", (m, cn))

sc, m, cn, ev, need = _pick_book("深圳前海某某贸易有限公司", C)
ck("完全无关 → none 且要人工指定", m == "none" and cn is None and need == 1, (m, cn, need))

sc, m, cn, ev, need = _pick_book("广州蓝海医药有限公司佛山分公司", C)
ck("长名包含短名 → contain 且要确认", m == "contain" and cn == "广州蓝海医药有限公司" and need == 1,
   (m, cn, need, sc))

ck("归一去掉公司类型后缀", _book_key("广州蓝海医药有限公司") == _book_key("广州蓝海医药股份有限公司"),
   (_book_key("广州蓝海医药有限公司"), _book_key("广州蓝海医药股份有限公司")))
ck("归一去掉尾部编码", _book_key("上海万联大药房有限公司(43701)") == "上海万联大药房",
   _book_key("上海万联大药房有限公司(43701)"))
ck("保守归一保留门店前缀", "万嘉" in _book_key("（万嘉店）上海固生堂万嘉中医门诊有限公司"),
   _book_key("（万嘉店）上海固生堂万嘉中医门诊有限公司"))
ck("激进归一去掉门店前缀", not _book_key("（万嘉店）上海固生堂万嘉中医门诊有限公司", True).startswith("万嘉"),
   _book_key("（万嘉店）上海固生堂万嘉中医门诊有限公司", True))

print("=== B. 第零组 · 主体完整性 ===")
c = duckdb.connect()
c.execute("CREATE TABLE tb(company_name VARCHAR, ytd_debit DOUBLE, ytd_credit DOUBLE)")
c.execute("CREATE TABLE journal(company_name VARCHAR, acct_code VARCHAR, acct_l1 VARCHAR, "
          "vdate VARCHAR, debit DOUBLE, credit DOUBLE)")
c.execute("CREATE TABLE _books(book_name VARCHAR, company_name VARCHAR, short_name VARCHAR, "
          "src VARCHAR, imported_at TIMESTAMP)")
for cn_, amt in (("A公司", 100.0), ("B公司", 50.0), ("D公司", 10.0)):
    c.execute("INSERT INTO tb VALUES (?,?,?)", [cn_, amt, amt])
for cn_ in ("A公司", "C公司", "D公司"):
    c.execute("INSERT INTO journal VALUES (?,?,?,?,?,?)", [cn_, "1001", "库存现金", "2026-01-01", 1.0, 0.0])
for cn_ in ("A公司", "B公司", "C公司"):
    c.execute("INSERT INTO _books VALUES (?,?,?,?,?)", [cn_, cn_, "", "test", None])

r = recon._check_entity_scope(c)
parts = {p["key"]: p for p in r["parts"]}
ck("缺余额表 = C 公司", [x[0] for x in parts["no_tb"]["rows"]] == ["C公司"], parts["no_tb"]["rows"])
ck("缺序时账 = B 公司", [x[0] for x in parts["no_jr"]["rows"]] == ["B公司"], parts["no_jr"]["rows"])
ck("超范围 = D 公司（两张表都有）", [x[0] for x in parts["out_scope"]["rows"]] == ["D公司"],
   parts["out_scope"]["rows"])
ck("超范围说明了它出现在哪", "余额表" in str(parts["out_scope"]["rows"][0][1])
   and "序时账" in str(parts["out_scope"]["rows"][0][1]), parts["out_scope"]["rows"][0])
ck("总体不通过（3 项）", r["ok"] is False and r["bad"] == 3, (r["ok"], r["bad"]))
ck("三张结果表都落了库", all(c.execute(
    "SELECT count(*) FROM information_schema.tables WHERE table_name=?", [t]).fetchone()[0] == 1
    for t in ("_recon_scope_no_tb", "_recon_scope_no_jr", "_recon_scope_out")))

# 没有对照表时：只能查「缺对方表」，超范围不报，且要有警告
c2 = duckdb.connect()
c2.execute("CREATE TABLE tb(company_name VARCHAR, ytd_debit DOUBLE, ytd_credit DOUBLE)")
c2.execute("CREATE TABLE journal(company_name VARCHAR, acct_code VARCHAR, acct_l1 VARCHAR, "
           "vdate VARCHAR, debit DOUBLE, credit DOUBLE)")
c2.execute("CREATE TABLE _books(book_name VARCHAR, company_name VARCHAR, short_name VARCHAR, "
           "src VARCHAR, imported_at TIMESTAMP)")
c2.execute("INSERT INTO tb VALUES ('X公司',1,1)")
c2.execute("INSERT INTO journal VALUES ('Y公司','1001','现金','2026-01-01',1,0)")
r2 = recon._check_entity_scope(c2)
p2 = {p["key"]: p for p in r2["parts"]}
ck("无对照表 → 有警告提醒先挂表", any("对照表" in w for w in r2.get("warnings") or []), r2.get("warnings"))
ck("无对照表 → 不报超范围", p2["out_scope"]["rows"] == [], p2["out_scope"]["rows"])
ck("无对照表 → 缺表仍能查（X 缺序时账）", [x[0] for x in p2["no_jr"]["rows"]] == ["X公司"],
   p2["no_jr"]["rows"])


print("=== C. 推断链路（精确匹配失败 → 推断补上 → 写留痕） ===")
import json as _json
from core.db import Ledger

led = Ledger.__new__(Ledger)          # 只借方法，不跑 __init__
led.conn = duckdb.connect()
led.conn.execute("CREATE TABLE _books(book_name VARCHAR, company_name VARCHAR, "
                 "short_name VARCHAR, src VARCHAR, imported_at TIMESTAMP)")
led.conn.execute("CREATE TABLE _book_match(book_name VARCHAR PRIMARY KEY, company_name VARCHAR, "
                 "method VARCHAR, score DOUBLE, evidence VARCHAR, need_confirm INTEGER, "
                 "src VARCHAR, updated_at TIMESTAMP)")
led.conn.execute("CREATE TABLE tb(book_name VARCHAR, company_name VARCHAR)")
led.conn.execute("INSERT INTO _books VALUES ('佛山固生堂中医门诊部有限公司','佛山固生堂中医门诊部有限公司','','t',NULL)")
led.conn.execute("INSERT INTO tb VALUES ('（东湖店）佛山固生堂中医门诊部有限公司(99999)', NULL)")

stat = led.infer_book_matches("tb")
rows = led.conn.execute("SELECT book_name, company_name, method, need_confirm FROM _book_match").fetchall()
ck("推断写进了留痕表", len(rows) == 1 and rows[0][1] == "佛山固生堂中医门诊部有限公司", rows)
ck("门店不同 → norm_prefix 且标需确认", rows[0][2] == "norm_prefix" and rows[0][3] == 1, rows[0])
ck("统计里记了 1 条需确认", stat.get("need_confirm") == 1, stat)

# 对照表命中的（这条 book_name 与对照表完全一致）也要留痕，且不需确认
led.conn.execute("INSERT INTO tb VALUES ('佛山固生堂中医门诊部有限公司', NULL)")
led.infer_book_matches("tb")
rows = led.conn.execute(
    "SELECT book_name, method, need_confirm FROM _book_match ORDER BY book_name").fetchall()
ck("对照表直接命中也留痕、不标确认",
   any(r[0] == "佛山固生堂中医门诊部有限公司" and r[1] == "norm" and r[2] == 0 for r in rows), rows)

print("\n%s  共 %d 项，失败 %d" % ("ALL PASS" if not fails else "FAILED", n[0], len(fails)))
if fails:
    for f in fails:
        print("   -", f)
sys.exit(1 if fails else 0)
