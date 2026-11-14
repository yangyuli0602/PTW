# -*- coding: utf-8 -*-
"""验证第五组「辅助核算可用性」的四种判定（内存库，不碰真实 data）"""
import sys, json
sys.path.insert(0, r"D:\桌面\FDE工作\01_PTW框架\程序")
import duckdb
from core import recon


def make(with_aux=True, raw_cols="aux", raw_filled=True, mapping_aux=True, dropped=0):
    con = duckdb.connect()
    con.execute("""CREATE TABLE tb(company_name VARCHAR, acct_code VARCHAR, acct_name VARCHAR,
        acct_l1 VARCHAR, currency VARCHAR, aux1 VARCHAR, aux2 VARCHAR, aux3 VARCHAR, aux_text VARCHAR,
        bank_account VARCHAR, bank_name VARCHAR)""")
    a1 = "账号010" if with_aux else None
    a2 = "H公司" if with_aux else None
    con.executemany("INSERT INTO tb VALUES (?,?,?,?,?,?,?,?,?,?,?)", [
        ("1A公司", "1001", "库存现金", "库存现金", "人民币", None, None, None, None, None, None),
        ("1A公司", "1002", "银行存款", "银行存款", "人民币", a1, a2, None, "账号010|H公司" if with_aux else None, None, None),
        ("2B公司", "1002", "银行存款", "银行存款", "美元", a1, a2, None, "账号010|H公司" if with_aux else None, None, None),
        ("2B公司", "1012", "其他货币资金", "其他货币资金", "人民币", None, None, None, None, None, None),
    ])
    con.execute("""CREATE TABLE journal(book_name VARCHAR, acct_code VARCHAR, acct_name VARCHAR,
        acct_l1 VARCHAR, summary VARCHAR, aux1 VARCHAR, aux2 VARCHAR, aux3 VARCHAR, aux_text VARCHAR)""")
    con.executemany("INSERT INTO journal VALUES (?,?,?,?,?,?,?,?,?)", [
        ("1A公司", "1002.01", "银行存款_工商银行", "银行存款", "付款", None, None, None, None),
        ("1A公司", "1002.01", "银行存款_工商银行", "银行存款", "收款", None, None, None, None),
        ("2B公司", "1002", "银行存款", "银行存款", "结汇", None, None, None, None),
    ])
    mp = {"acct_code": "科目编码", "acct_name": "科目名称"}
    if mapping_aux:
        mp.update({"aux1": "核算维度编码", "aux2": "核算维度名称"})
    con.execute("CREATE TABLE _sheets(sid VARCHAR, file_id VARCHAR, file_name VARCHAR, sheet_name VARCHAR,"
                " header_row INT, header_rows INT, table_name VARCHAR, col_names VARCHAR,"
                " n_rows BIGINT, target VARCHAR, mapping VARCHAR, imported_at TIMESTAMP,"
                " clean VARCHAR, stg_table VARCHAR)")
    con.execute("INSERT INTO _sheets VALUES ('s1','f1','tb.xlsx','源表',2,1,'raw_s1',NULL,4,'tb',?,NULL,NULL,NULL)",
                [json.dumps(mp, ensure_ascii=False)])
    cols = {"plain": ["方案层级", "科目编码", "科目名称", "币别"]}
    cols["aux"] = cols["plain"] + ["核算维度编码", "核算维度名称"]
    con.execute('CREATE TABLE raw_s1(%s)' % ", ".join('"%s" VARCHAR' % c for c in cols[raw_cols]))
    if raw_cols == "aux":
        v = ("账号010", "H公司") if raw_filled else (None, None)
        con.execute('INSERT INTO raw_s1 VALUES (?,?,?,?,?,?)',
                    ["1A公司(100)", "1001", "库存现金", "人民币", None, None])
        con.execute('INSERT INTO raw_s1 VALUES (?,?,?,?,?,?)',
                    ["1A公司(100)", "1002", "银行存款", "人民币", v[0], v[1]])
    if dropped:
        con.execute("CREATE TABLE _dropped_s1(step INT, step_name VARCHAR, reason VARCHAR, src_row INT)")
        for i in range(dropped):
            con.execute("INSERT INTO _dropped_s1 VALUES (1,'删空行','全列为空',?)", [i + 50])
    return con


CASES = [
    ("正常：有辅助核算", dict(), "present"),
    ("原始表就没有该列", dict(with_aux=False, raw_cols="plain"), "never_present"),
    ("原始表有列但全空", dict(with_aux=False, raw_filled=False), "never_filled"),
    ("映射漏配", dict(with_aux=False, mapping_aux=False), "not_mapped"),
    ("映射配了但结果空（疑清洗丢）", dict(with_aux=False, dropped=12), "lost"),
]

for name, kw, want in CASES:
    con = make(**kw)
    r = recon._check_aux(con)
    got = r["verdict"]
    flag = "✓" if got == want else "✗"
    print("%s %-32s 判定=%-14s 期望=%-14s ok=%s bad=%s" % (flag, name, got, want, r["ok"], r["bad"]))
    for n in r["notes"][:2]:
        print("      ·", n[:100])
    con.close()
