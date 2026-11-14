# -*- coding: utf-8 -*-
"""账户清单 + 银行流水配对。

银行流水本身不知道自己属于哪个账户、对方那笔是不是自己人划转。
这两件事都要靠「账户清单」——它来自科目余额表的货币资金明细（账号/户名/开户行），
补上上年 WP 回笼的账号和人工补录，成为一张可复用的主数据。

设计上的三条硬规矩：
1. 清单是**表**不是变量（`_accounts`），能导出、能人工回填、能反复使用。
2. 配对**可重跑**：清单补全后重新 link 一次即可，不用重导流水。
3. 判不出就留空，绝不猜。`is_internal` 三态（1/0/空），空必须能单独被看见。
"""
import re

from .db import _acct_norm, qi, ql

# 货币资金科目：现金没有账号，但也要进清单（否则 BKD 会漏行）
CASH_L1 = ("库存现金", "银行存款", "其他货币资金")
CASH_CODE = ("1001", "1002", "1012")

BANK_WORDS = re.compile(r"银行|支行|分行|信用社|营业部|信合|农信|邮储|信银")
ACCT_NO_RX = re.compile(r"[0-9][0-9\-]{5,}")
SUM_WORDS = ("合计", "小计", "总计", "合计数")
# 多账簿导出里的汇总层，不是主体。留着会让账户清单多出一堆假公司
FAKE_COMPANY = re.compile(r"多账簿|合并方案|汇总|全部账簿|^小计|^合计")
TAIL_CODE = re.compile(r"\s*[（(]\s*\d{2,4}\s*[)）]\s*$")


def _t(v):
    return "" if v is None else str(v).strip()


def _pick_acct(row):
    """从一行数据里找账号。优先级：专用列 → 辅助核算槽（取最长的那串数字）。

    账号长度差别大（对公 12~19 位，外币户可能带连字符），所以取最长的候选，
    短串多半是行号、部门号之类的误命中。
    """
    cands = []
    if _t(row.get("bank_account")):
        cands.append(_t(row["bank_account"]))
    for slot in ("aux1", "aux2", "aux3", "aux_text"):
        v = _t(row.get(slot))
        if v:
            cands.extend(ACCT_NO_RX.findall(v))
    if not cands:
        return ""
    return max(cands, key=len).strip("-")


def _pick_bank(row, acct_name):
    """开户行：先看辅助核算里带"银行"字样的，再从科目名称拆。

    拆的时候要排除「银行存款」本身——它含"银行"但不是机构名。
    """
    for slot in ("aux1", "aux2", "aux3", "aux_text"):
        v = _t(row.get(slot))
        if v and BANK_WORDS.search(v) and v not in CASH_L1:
            return v
    parts = [p.strip() for p in re.split(r"[-－—–_/|]+", _t(acct_name))]
    parts = [p for p in parts
             if p and p not in CASH_L1 and not re.fullmatch(r"[0-9]+", p)
             and not re.fullmatch(r"[借贷]", p)]
    return parts[0] if parts else ""


def _norm_company(name):
    """主体名归一去尾码：同一家公司可能被导成「A公司(100)」「A公司（101）」。

    不归一的话同一家会拆成两家，账户清单里出现重复账户。
    """
    s = _t(name)
    if not s:
        return ""
    return TAIL_CODE.sub("", s).strip()


def build(conn):
    """从 tb 生成账户清单，写 `_accounts`。返回统计 dict。

    只读 tb，不改任何账套数据。重复执行结果一致（整表重建）。
    """
    try:
        have = {r[0] for r in conn.execute("DESCRIBE tb").fetchall()}
    except Exception as e:
        return {"error": "读不到科目余额表：%s" % e, "total": 0}
    if not have:
        return {"error": "还没有合并科目余额表，请先导入并合并 tb", "total": 0}

    def c(name):
        return name if name in have else "NULL"

    sql = """
        SELECT {company} AS company_name,
               {acct_name} AS acct_name,
               {acct_l1} AS acct_l1,
               {acct_code} AS acct_code,
               {currency} AS currency,
               {bank} AS bank_account,
               {aux1} AS aux1, {aux2} AS aux2, {aux3} AS aux3, {aux_text} AS aux_text
        FROM tb
        WHERE coalesce({acct_l1}, '') IN ('库存现金','银行存款','其他货币资金')
           OR starts_with(trim(coalesce({acct_code}, '')), '1001')
           OR starts_with(trim(coalesce({acct_code}, '')), '1002')
           OR starts_with(trim(coalesce({acct_code}, '')), '1012')
    """.format(company=c("company_name"), acct_name=c("acct_name"), acct_l1=c("acct_l1"),
               acct_code=c("acct_code"), currency=c("currency"), bank=c("bank_account"),
               aux1=c("aux1"), aux2=c("aux2"), aux3=c("aux3"), aux_text=c("aux_text"))
    try:
        rows = conn.execute(sql).fetchall()
    except Exception as e:
        return {"error": "查询货币资金科目失败：%s" % e, "total": 0}

    accounts = {}
    skipped_sum = 0
    for (company, acct_name, l1, code, currency, bank_col,
         aux1, aux2, aux3, aux_text) in rows:
        row = {"bank_account": bank_col, "aux1": aux1, "aux2": aux2, "aux3": aux3,
               "aux_text": aux_text}
        # 多账簿导出里每科目都有一条「合计」维度行，和明细并存会重复计账户
        if any(_t(v) in SUM_WORDS for v in (aux1, aux2, aux3, aux_text)):
            skipped_sum += 1
            continue
        co = _norm_company(company)
        if not co or FAKE_COMPANY.search(co):
            skipped_sum += 1
            continue
        subj = _t(acct_name) or _t(l1)
        subj_sub = _t(acct_name)
        if "库存现金" in (_t(l1) + subj_sub) or _t(code) == "1001":
            # 现金没有账号，按「主体 + 库存现金」单独成一个账户
            key = (co, "现金")
            accounts.setdefault(key, {
                "acct_key": "", "acct_no": "", "acct_name": subj or "库存现金",
                "bank_name": "", "company_name": co, "currency": _t(currency),
                "source": "tb", "status": "non_bank", "note": "现金类账户，无账号",
            })
            continue
        no = _pick_acct(row)
        acct_key = _acct_norm_py(no)
        bank = _pick_bank(row, subj)
        key = (co, acct_key or ("%s|%s" % (subj, bank)))
        cur = accounts.get(key)
        item = {
            "acct_key": acct_key, "acct_no": no, "acct_name": subj,
            "bank_name": bank, "company_name": co,
            "currency": _t(currency), "source": "tb",
            "status": "ok" if acct_key else "no_acct",
            "note": "" if acct_key else "账面没有账号（辅助核算缺失或未填），无法与流水配对",
        }
        if cur is None:
            accounts[key] = item
        else:
            # 同一账户多币种行：账号相同即视为一个账户，币种合并展示
            if item["currency"] and item["currency"] not in (cur["currency"] or ""):
                cur["currency"] = "/".join(x for x in (cur["currency"], item["currency"]) if x)
            cur["bank_name"] = cur["bank_name"] or item["bank_name"]
            cur["acct_name"] = cur["acct_name"] or item["acct_name"]

    conn.execute("DROP TABLE IF EXISTS _accounts")
    conn.execute("""CREATE TABLE _accounts(
        acct_key VARCHAR, acct_no VARCHAR, acct_name VARCHAR, bank_name VARCHAR,
        company_name VARCHAR, currency VARCHAR, source VARCHAR,
        status VARCHAR, note VARCHAR)""")
    if accounts:
        conn.executemany(
            "INSERT INTO _accounts VALUES (?,?,?,?,?,?,?,?,?)",
            [(a["acct_key"], a["acct_no"], a["acct_name"], a["bank_name"],
              a["company_name"], a["currency"], a["source"], a["status"], a["note"])
             for a in accounts.values()])

    n_ok = sum(1 for a in accounts.values() if a["status"] == "ok")
    n_noacct = sum(1 for a in accounts.values() if a["status"] == "no_acct")
    return {"total": len(accounts), "ok": n_ok, "no_acct": n_noacct,
            "non_bank": len(accounts) - n_ok - n_noacct,
            "skippedSum": skipped_sum, "companies": len({a["company_name"] for a in accounts.values()})}


def _acct_norm_py(s):
    """Python 侧的同一条归一规则，必须和 SQL 里的 _acct_norm 结果一致。"""
    if s is None:
        return ""
    t = str(s)
    t = t.translate(str.maketrans("０１２３４５６７８９　－—−", "0123456789 ---"))
    t = re.sub(r"[ 　—−\-]", "", t.strip())
    t = re.sub(r"^0+", "", t)
    return t


def list_accounts(conn, limit=500):
    """账户清单，页面和导出用。有账号的排前面，没账号的排后面——先看能用的。"""
    try:
        rows = conn.execute("""
            SELECT acct_key, acct_no, acct_name, bank_name, company_name, currency,
                   source, status, note
            FROM _accounts
            ORDER BY CASE status WHEN 'ok' THEN 0 WHEN 'non_bank' THEN 1 ELSE 2 END,
                     company_name, bank_name
            LIMIT ?
        """, [limit]).fetchall()
    except Exception:
        return []
    keys = ("acctKey", "acctNo", "acctName", "bank", "company", "currency",
            "source", "status", "note")
    return [dict(zip(keys, [_t(v) for v in r])) for r in rows]


def has_list(conn):
    try:
        return conn.execute("SELECT count(*) FROM _accounts").fetchone()[0] > 0
    except Exception:
        return False


def link(conn, only_missing=True):
    """把 flow 和账户清单对上。

    only_missing=True 时只处理还没判定的行（第一次 link 用）；
    False 时全表重算（清单补全后重跑用）。
    """
    if not has_list(conn):
        return {"error": "账户清单为空。请先在「WP初稿生成」里生成 BKD，"
                         "或导入科目余额表后点「重建账户清单」。"}
    try:
        conn.execute("DESCRIBE flow")
    except Exception:
        return {"error": "还没有合并银行流水，请先导入并合并 flow"}

    cond = "AND coalesce(f.is_internal, '') = ''" if only_missing else ""
    # 1) 本方账号 → 主体 / 户名 / 开户行（流水自带的不覆盖，以流水原文为准）
    conn.execute("""
        UPDATE flow AS f SET
            company_name = coalesce(nullif(trim(coalesce(f.company_name,'')), ''), a.company_name),
            acct_name    = coalesce(nullif(trim(coalesce(f.acct_name,'')), ''),   a.acct_name),
            bank_name    = coalesce(nullif(trim(coalesce(f.bank_name,'')), ''),   a.bank_name),
            acct_source  = coalesce(nullif(trim(coalesce(f.acct_source,'')), ''),
                                    CASE WHEN a.acct_key IS NULL THEN NULL ELSE '账面账户清单' END)
        FROM _accounts AS a
        WHERE f.acct_key = a.acct_key AND f.acct_key IS NOT NULL {cond}
    """.format(cond=cond))

    # 2) 对方账号命中清单 → 内部；账号有值但没命中 → 外部
    conn.execute("""
        UPDATE flow AS f SET
            cp_company = a.company_name,
            is_internal = '1',
            internal_reason = '对方账号在账面账户清单内'
        FROM _accounts AS a
        WHERE f.cp_key = a.acct_key AND f.cp_key IS NOT NULL
          AND f.acct_key IS NOT NULL          -- 本方都不知道是哪个账户，结论不成立
          AND a.status IN ('ok', 'non_bank') {cond}
    """.format(cond=cond))

    conn.execute("""
        UPDATE flow AS f SET
            is_internal = '0',
            internal_reason = '对方账号不在账面账户清单内（外部交易）'
        WHERE coalesce(f.is_internal, '') = ''
          AND f.acct_key IS NOT NULL
          AND f.cp_key IS NOT NULL AND f.cp_key <> ''
    """)

    # 3) 只有户名没有账号：按户名精确匹配（不做模糊，模糊会误判成本公司）
    conn.execute("""
        UPDATE flow AS f SET
            cp_company = a.company_name,
            is_internal = '1',
            internal_reason = '对方户名与账户清单户名一致（无对方账号，按户名判定）'
        FROM _accounts AS a
        WHERE coalesce(f.is_internal, '') = ''
          AND (f.cp_key IS NULL OR f.cp_key = '')
          AND f.acct_key IS NOT NULL
          AND nullif(trim(coalesce(f.counterparty_name, '')), '') IS NOT NULL
          AND trim(f.counterparty_name) = trim(a.acct_name)
    """)

    # 4) 剩下的：连对方是谁都不知道
    conn.execute("""
        UPDATE flow AS f SET
            is_internal = '',
            internal_reason = '缺对方账号/户名，无法判定；或本方账号未配对到账户清单'
        WHERE coalesce(f.is_internal, '') = ''
    """)

    stat = conn.execute("""
        SELECT coalesce(is_internal,'') AS k, count(*) FROM flow GROUP BY 1 ORDER BY 1
    """).fetchall()
    stat = {str(k): n for k, n in stat}
    rows_linked = conn.execute(
        "SELECT count(*) FROM flow WHERE coalesce(acct_source,'') LIKE '%清单%'").fetchone()[0]
    return {"internal": stat.get("1", 0), "external": stat.get("0", 0),
            "unknown": stat.get("", 0), "rows": sum(stat.values()),
            "acctLinked": rows_linked,
            "accounts": conn.execute("SELECT count(*) FROM _accounts").fetchone()[0],
            "okAccounts": conn.execute(
                "SELECT count(*) FROM _accounts WHERE status = 'ok'").fetchone()[0]}


def gap_list(conn, limit=200):
    """配对不上的清单，给页面和 Excel 检查表用。

    分三类，因为处理动作完全不同：
      - 本方账号没配上 → 账户清单缺这个账户，去补清单
      - 对方有账号但判成外部 → 可能清单缺对方账户，也可能真是外部，要人看
      - 对方什么都没有 → 流水导出不全，只能找银行补
    """
    out = []
    try:
        rows = conn.execute("""
            SELECT acct_no, acct_key, acct_name, bank_name, counterparty_acct,
                   counterparty_name, is_internal, internal_reason, count(*) AS n,
                   sum(amount) AS amt
            FROM flow
            WHERE coalesce(is_internal, '') = ''
               OR (coalesce(acct_source, '') = '' AND coalesce(acct_no,'') <> '')
            GROUP BY 1,2,3,4,5,6,7,8
            ORDER BY n DESC LIMIT ?
        """, [limit]).fetchall()
    except Exception:
        return out
    for (acct_no, acct_key, acct_name, bank, cp_acct, cp_name,
         internal, reason, n, amt) in rows:
        out.append({
            "acct": _t(acct_no), "acctName": _t(acct_name), "bank": _t(bank),
            "cpAcct": _t(cp_acct), "cpName": _t(cp_name),
            "kind": ("本方账户未入清单" if not _t(acct_key)
                     else ("对方无账号无户名" if not _t(cp_acct) and not _t(cp_name)
                           else ("对方只有户名，未匹配到清单" if not _t(cp_acct)
                                 else "对方账号未命中清单"))),
            "reason": _t(reason), "rows": int(n or 0),
            "amount": round(float(amt or 0), 2),
        })
    return out
