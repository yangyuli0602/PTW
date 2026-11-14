# -*- coding: utf-8 -*-
"""数据检查：四组平衡核对。

一、科目余额表 · 借贷发生额与余额平衡（按公司名称）
    本期借/贷发生额合计是否相等；本年累计借/贷是否相等；
    期初净额 + 本期借 − 本期贷 = 期末净额（余额勾稽）。
    这三条都是「同一主体全表合计」层面的恒等式，明细行单看某一科目本来就不平。

二、序时账 · 借贷平衡（按公司名称，三层）
    2.1 总额平衡   Σ借 = Σ贷
    2.2 月平衡     每个会计期间 Σ借 = Σ贷
    2.3 凭证平衡   每张凭证 Σ借 = Σ贷
    凭证的唯一性靠「唯一凭证号」：优先用字段映射里的 voucher_key 列
    （数据清洗「生成唯一凭证号」产出的那一列）；没映射就现场按
    公司名称 + 期间 + 凭证号 拼一个，并在提示里说明。

三、科目余额表 × 序时账 · 按「公司名称 + 一级科目」平衡
    口径同旧版：只取「一级科目」列（acct_l1），两边都必须有值。

四、科目余额表 × 序时账 · 按「公司名称 + 科目全称」平衡
    键优先取 acct_full（清洗「生成一级科目」产出的「科目全称」列），
    没有就退到 acct_name（科目名称）。退化会在提示里明确写出来。

为什么主体取「公司名称 → 账簿名称」
    多公司（集团合并、代账多户）混在一套账里时，只按科目合计看不出是哪家对不上。
    主体优先公司名称，没有公司名称就用账簿名称；两者都没有归到「(未标注主体)」，
    单主体账套下退化成全表合计，行为不变。

返回结构（给前端 / 导出共用）
    {
      "checks": [
        {"key","no","title","desc","ok","total","bad","error","notes","warnings",
         "parts":[{"key","title","columns","rows","total","bad","table","note"}]}
      ],
      "summary": {...}, "tables": [(sheet_name, table_name), ...]
    }
每个 part 都会落成一张 _recon_* 表，导出时直接读表，不重算。
"""
import json
import re as _re

from .fields import to_cn


def qi(name):
    return '"' + str(name).replace('"', '""') + '"'


def _q(s):
    return "'" + str(s).replace("'", "''") + "'"


def _scalar(conn, sql, default=None):
    try:
        r = conn.execute(sql).fetchone()
        return r[0] if r else default
    except Exception:
        return default


def _has_table(conn, name):
    try:
        return bool(_scalar(conn, "SELECT count(*) FROM information_schema.tables "
                                  "WHERE table_name=%s" % _q(name), 0))
    except Exception:
        return False


def _cols(conn, table):
    try:
        cur = conn.execute('SELECT * FROM "%s" LIMIT 0' % table)
        return [d[0] for d in cur.description]
    except Exception:
        return []


def _exec(conn, sql):
    """执行并返回 (columns, rows)，Decimal 统一转成 float 方便前端/Excel。"""
    cur = conn.execute(sql)
    cols = [d[0] for d in cur.description]
    return cols, [[_num(v) for v in r] for r in cur.fetchall()]


def _num(v):
    try:
        from decimal import Decimal
        if isinstance(v, Decimal):
            return float(v)
    except Exception:
        pass
    return v


def _save(conn, table, sql):
    """把结果落成表，便于导出和二次查询。空结果也能落（CREATE TABLE AS 0 行）。"""
    conn.execute('DROP TABLE IF EXISTS "%s"' % table)
    conn.execute('CREATE TABLE "%s" AS %s' % (table, sql))


# ---------------------------------------------------------------- 主体
def entity_expr(cols, alias=""):
    """主体表达式：公司名称 → 账簿名称 → (未标注主体)。"""
    p = (alias + ".") if alias else ""
    parts = []
    if "company_name" in cols:
        parts.append("nullif(trim(%scompany_name), '')" % p)
    if "book_name" in cols:
        parts.append("nullif(trim(%sbook_name), '')" % p)
    if not parts:
        return "'(未标注主体)'"
    return "coalesce(%s, '(未标注主体)')" % ", ".join(parts)


# ---------------------------------------------------------------- 币别
# 「本位币视图」的几种写法。多币别长表里汇总层一般写「综合本位币」，
# 单币种系统往往只写「人民币」，或者干脆不写。三种都要能认。
BASE_CUR_WORDS = ("综合本位币", "本位币", "记账本位币", "本币", "人民币本位币")
BASE_CUR_LABEL = "本位币（合计）"


def cur_base_pred(col="currency"):
    """行级：这一行是不是「本位币视图」的行（空币别也算）。"""
    return ("(nullif(trim(%s), '') IS NULL OR trim(%s) IN (%s))"
            % (col, col, ", ".join(_q(w) for w in BASE_CUR_WORDS)))


def cur_shape_sql(ent, table="tb"):
    """主体级币别形态（一段 SELECT，供 CTE 或子查询包）。

    币别不能只看单行来分组：同一个「人民币」，在这个主体身上是本位币，
    在另一个主体身上可能是外币。所以先按主体算两件事：

      has_base  该主体有没有「综合本位币」这类汇总行 → 有，其它币别才是真外币
      n_other   除本位币标记外的 distinct 币别数 → 0/1 说明是单币种主体

    单币种主体必须并到本位币组：余额表按「人民币」分组、序时账全归本位币组，
    两边永远对不上，会成批报「仅序时账有 / 仅余额表有」（V1.2.04 飞书 3.3）。
    """
    b = cur_base_pred("currency")
    return ("SELECT {ent} AS entity, "
            "max(CASE WHEN {b} THEN 1 ELSE 0 END) AS has_base, "
            "count(DISTINCT CASE WHEN NOT {b} THEN trim(currency) END) AS n_other "
            'FROM "{t}" GROUP BY 1').format(ent=ent, b=b, t=table)


def cur_group_preds(alias="ce", t_alias="t"):
    """返回 (本位币组条件, 外币组条件)，用于在主体形态已知后给每行归组。

    本位币组 = 标记行本身；该主体压根没有标记行时，整个主体都算本位币组。
    外币组   = 标记行之外的币别，且该主体确实有外币形态
               （有标记行，或者币别多于一种——后者没有标记行，本币组取本币列合计，
                原币组仍要保留，否则真多币种的客户会丢掉原币核对）。
    """
    b = cur_base_pred("%s.currency" % t_alias)
    base = "(%s OR %s.has_base = 0)" % (b, alias)
    fx = "(NOT %s AND (%s.has_base = 1 OR %s.n_other >= 2))" % (b, alias, alias)
    return base, fx


def _period_expr(cols, alias=""):
    """会计期间表达式：period 列优先，没有就用 vdate 推出 YYYY-MM。"""
    p = (alias + ".") if alias else ""
    if "period" in cols:
        return "coalesce(nullif(trim(%speriod), ''), '(无期间)')" % p
    if "vdate" in cols:
        return "coalesce(strftime(%svdate, '%%Y-%%m'), '(无期间)')" % p
    return None


def _period_fmt(conn):
    """判断序时账期间列的格式，返回 (fmt, year)。

    fmt = 'm'  纯月份（1..12）
        = 'ym' 带年份（2026-01 / 2026.01 / 2026/1 / 2026年1期 / 202601 …）
        = None 无法判断
    year 仅在 ym 模式有值，用于把「用户只填月份」补成年份。
    """
    try:
        rows = conn.execute(
            'SELECT DISTINCT trim(coalesce(period, \'\')) FROM "journal" '
            'WHERE nullif(trim(coalesce(period, \'\')), \'\') IS NOT NULL LIMIT 40'
        ).fetchall()
    except Exception:
        return None, None
    vals = [str(r[0]).strip() for r in rows if str(r[0] or "").strip()]
    if not vals:
        return None, None
    year, has_year, all_month = None, False, True
    for v in vals:
        m = _re.fullmatch(r"(\d{4})[-年./](\d{1,2})(?:期)?", v)
        if m:
            has_year, year = True, (year or m.group(1))
            continue
        if _re.fullmatch(r"\d{6}", v):
            has_year, year = True, (year or v[:4])
            continue
        if not _re.fullmatch(r"\d{1,2}", v):
            all_month = False
    if has_year:
        return "ym", year
    if all_month:
        return "m", None
    return None, None


def _period_key(expr, fmt="ym", year=None):
    """期间比较键：把各种格式归一成可排序的形式。

    fmt='m'  纯数字补零到 2 位（"1" -> "01"）。
    fmt='ym' 统一成 "YYYY-MM"：识别 2026-01 / 2026.01 / 2026/1 /
             2026年1期 / 2026年1 / 202601；识别不了的原样返回（字典序兜底）。

    **纯月号在 ym 模式下要补年份**（2026-11-11 修）：明细表常出现「同列既有 2026-01
    又有 01」的混写——前者来自系统导出，后者来自明细表。原先把 01 原样留着，
    和用户传入的 2026-01 按字符串比，"01" < "2026-01" 恒成立，
    这批行会被期间过滤静默丢掉（实测香港 3,253 行凭空消失，组三误报 24 条"仅余额表有"）。
    """
    p = "trim(coalesce(%s, ''))" % expr
    if fmt == "m":
        return ("CASE WHEN regexp_full_match(%s, '\\d+') THEN lpad(%s, 2, '0') "
                "ELSE %s END" % (p, p, p))
    pure = "lpad(%s, 2, '0')" % p
    if year:
        pure = "'%s-' || lpad(%s, 2, '0')" % (year, p)
    return ("CASE "
            "WHEN regexp_full_match(%(p)s, '\\d{4}[-年./]\\d{1,2}(期)?') "
            "THEN left(%(p)s, 4) || '-' || lpad(regexp_extract(%(p)s, '[-年./](\\d{1,2})', 1), 2, '0') "
            "WHEN regexp_full_match(%(p)s, '\\d{6}') "
            "THEN left(%(p)s, 4) || '-' || right(%(p)s, 2) "
            "WHEN regexp_full_match(%(p)s, '\\d{1,2}') "
            "THEN %(pure)s "
            "ELSE %(p)s END") % {"p": p, "pure": pure}


def _pval(v, fmt="ym", year=None):
    """用户输入的期间按「数据实际格式」归一，保证和 _period_key 落在同一套键上。"""
    s = str(v or "").strip()
    if fmt == "m":
        m = _re.fullmatch(r"(\d{4})[-年./](\d{1,2})(?:期)?", s)
        if m:
            return m.group(2).zfill(2)
        if _re.fullmatch(r"\d{6}", s):
            return s[4:].zfill(2)
        if _re.fullmatch(r"\d{1,2}", s):
            return s.zfill(2)
        return s
    m = _re.fullmatch(r"(\d{4})[-年./](\d{1,2})(?:期)?", s)
    if m:
        return "%s-%s" % (m.group(1), m.group(2).zfill(2))
    if _re.fullmatch(r"\d{6}", s):
        return "%s-%s" % (s[:4], s[4:])
    if _re.fullmatch(r"\d{1,2}", s):
        return ("%s-%s" % (year, s.zfill(2))) if year else s.zfill(2)
    return s


def _period_cond(conn, frm, to, cols):
    """序时账期间过滤。返回不带 AND 的条件串（空串=不过滤）。"""
    if not frm and not to:
        return ""
    if "period" in cols:
        fmt, year = _period_fmt(conn)
        fmt = fmt or "ym"
        pk = _period_key("period", fmt, year)
        cond = []
        if frm:
            cond.append("%s >= %s" % (pk, _q(_pval(frm, fmt, year))))
        if to:
            cond.append("%s <= %s" % (pk, _q(_pval(to, fmt, year))))
        return " AND ".join(cond)
    if "vdate" in cols:
        cond = []
        if frm:
            cond.append("vdate >= _to_date(%s)" % _q(frm))
        if to:
            cond.append("vdate <= _to_date(%s)" % _q(to))
        return " AND ".join(cond)
    return ""


def _n_rows(conn, table, col=None):
    """表行数；给了 col 则统计该列非空行数。"""
    if col:
        return _scalar(conn, 'SELECT count(*) FROM "%s" WHERE nullif(trim(cast("%s" AS VARCHAR)), \'\') '
                             'IS NOT NULL' % (table, col), 0) or 0
    return _scalar(conn, 'SELECT count(*) FROM "%s"' % table, 0) or 0


# ---------------------------------------------------------------- 一、余额表自平衡
TB_BASES = [
    ("ytd", "本年累计(本币)"),
    ("cur", "本期发生(本币)"),
    ("open", "期初(本币)"),
    ("end", "期末(本币)"),
]


def _base_label(base):
    for k, lb in TB_BASES:
        if k == base:
            return lb
    return base


def _check_tb_self(conn, tol):
    """余额表：本年借方 vs 本年贷方，相减是否为零（2026-11-11 按用户要求精简）。

    一个公司一行。期初、本期发生、期末、原币都不再进这张表——它们属于别的核对口径，
    混在一起会让"不平"到底卡在哪一步说不清。要看那些，去查询页自己查。
    """
    res = {"key": "tb_self", "no": "一",
           "title": "科目余额表 · 本年借贷发生额平衡",
           "desc": "按公司汇总「本年累计借方」与「本年累计贷方」，两者相减为零即平衡。"
                   "不再核本期发生额、期初/期末勾稽与外币。",
           "error": None, "notes": [], "warnings": [], "parts": []}
    if not _has_table(conn, "tb"):
        res["error"] = "还没有科目余额表，先在第 3 步「合并与检查 → 合并成表」生成"
        return res
    cols = _cols(conn, "tb")
    if not cols:
        res["error"] = "科目余额表读取失败"
        return res
    if "ytd_debit" not in cols or "ytd_credit" not in cols:
        res["error"] = ("余额表没有「本年累计借方 / 本年累计贷方」列，无法核对。"
                        "到「字段映射」把 ytd_debit、ytd_credit 配上"
                        "（源表里通常叫「本年累计借方发生额」）。")
        return res

    ent = entity_expr(cols)
    sql = """
    SELECT {ent} AS "公司名称",
           count(*) AS "行数",
           round(sum(coalesce(ytd_debit, 0)), 2)  AS "本年借方",
           round(sum(coalesce(ytd_credit, 0)), 2) AS "本年贷方",
           round(sum(coalesce(ytd_debit, 0)) - sum(coalesce(ytd_credit, 0)), 2) AS "差额",
           CASE WHEN abs(sum(coalesce(ytd_debit, 0)) - sum(coalesce(ytd_credit, 0))) <= {tol}
                THEN '平衡' ELSE '不平' END AS "状态"
    FROM "tb" GROUP BY 1 ORDER BY 1
    """.format(ent=ent, tol=float(tol))

    columns, rows = _exec(conn, sql)
    _save(conn, "_recon_tb_self", sql)
    total = len(rows)
    si = columns.index("状态")
    bad = sum(1 for r in rows if r[si] != "平衡")

    res["notes"].append("按公司汇总，共 %d 个主体；只核「本年借方 − 本年贷方 = 0」" % total)
    if not bad:
        res["notes"].append("全部主体本年借贷相等")
    res["parts"].append({
        "key": "tb_self", "title": "各公司本年借贷发生额",
        "columns": columns, "rows": rows, "total": total, "bad": bad,
        "table": "_recon_tb_self", "status_col": "状态", "diff_cols": ["差额"],
        "view_cols": columns,
    })
    res["ok"] = bad == 0 and total > 0
    res["total"] = total
    res["bad"] = bad
    return res


# ---------------------------------------------------------------- 二、序时账自平衡
def _check_jr_self(conn, tol, period_from=None, period_to=None):
    res = {"key": "jr_self", "no": "二",
           "title": "序时账 · 借贷平衡（总额 / 月 / 唯一凭证号）",
           "desc": "按公司名称分组，检查全期借贷是否相等、每个月是否相等、每张凭证是否相等；"
                   "本币判平衡；原币仅展示",
           "error": None, "notes": [], "warnings": [], "parts": []}
    if not _has_table(conn, "journal"):
        res["error"] = "还没有序时账，先在第 3 步「合并与检查 → 合并成表」生成"
        return res
    cols = _cols(conn, "journal")
    if not cols:
        res["error"] = "序时账读取失败"
        return res
    if "debit" not in cols and "credit" not in cols:
        res["error"] = "序时账没有借方/贷方金额列，无法核对。请先在字段映射里映射金额列。"
        return res

    ent = entity_expr(cols)
    pexpr = _period_expr(cols)
    total_rows = _n_rows(conn, "journal")

    # 2026-11-11 按用户要求改口径：**会计期间严格按截止日执行**。
    # 清洗阶段保留全部月份（客户导出常带跨期凭证，删了会丢数据），但核对只看所选期间——
    # 否则 7 月以后的凭证会把差额带进 2.1/2.2/2.3，看起来像"1-6 月不平"。
    pcond = None
    if (period_from or period_to) and "period" in cols:
        try:
            pcond = _period_cond(conn, period_from, period_to, cols)
        except Exception:
            pcond = None
    jr_where = (" AND (%s)" % pcond) if pcond else ""
    jr_from = 'FROM "journal" WHERE 1=1' + jr_where
    if pcond:
        try:
            n_in = _scalar(conn, 'SELECT count(*) FROM "journal" WHERE (%s)' % pcond, 0) or 0
            n_out = _scalar(conn, 'SELECT count(*) FROM "journal" WHERE NOT (%s)' % pcond, 0) or 0
        except Exception:
            n_in = n_out = None
        if n_in is not None:
            res["notes"].append(
                "已按所选期间过滤：期间内 %d 行参与核对，期间外 %d 行不参与（清洗仍保留全部月份）"
                % (n_in, n_out))
    has_fx = bool(("debit_fx" in cols and _n_rows(conn, "journal", "debit_fx")) or
                  ("credit_fx" in cols and _n_rows(conn, "journal", "credit_fx")))
    d_fx = "sum(coalesce(debit_fx, 0))" if "debit_fx" in cols else "0"
    c_fx = "sum(coalesce(credit_fx, 0))" if "credit_fx" in cols else "0"
    fx_sel = (", {d} AS jr_debit_fx, {c} AS jr_credit_fx, {d} - {c} AS jr_diff_fx"
              .format(d=d_fx, c=c_fx)) if has_fx else \
             ", NULL AS jr_debit_fx, NULL AS jr_credit_fx, NULL AS jr_diff_fx"
    # 原币不参与"借贷平衡"判定：一张凭证可以借美元、贷人民币，原币层天然不等
    # （本位币折算后才是平的）。原币列只做展示，判定一律按本币。
    fx_ok = ""
    d_cny = "coalesce(sum(coalesce(debit, 0)), 0)"
    c_cny = "coalesce(sum(coalesce(credit, 0)), 0)"

    # ---- 2.1 总额平衡
    sql = """
    SELECT {ent} AS entity,
           count(*) AS n_jr_rows,
           {d}  AS jr_debit_sum,
           {c} AS jr_credit_sum,
           {d} - {c} AS jr_diff{fxsel},
           CASE WHEN abs({d} - {c}) > {tol}{fxok} THEN '不平' ELSE '平衡' END AS status
    {frm} GROUP BY 1 ORDER BY 1
    """.format(ent=ent, d=d_cny, c=c_cny, fxsel=fx_sel, tol=float(tol), fxok=fx_ok,
               frm=jr_from)
    c1, r1 = _exec(conn, sql)
    _save(conn, "_recon_jr_total", sql)
    b1 = sum(1 for r in r1 if r[c1.index("status")] == "不平")
    res["parts"].append({
        "key": "total", "title": "2.1 总额平衡", "columns": c1, "rows": r1,
        "total": len(r1), "bad": b1, "table": "_recon_jr_total",
        "status_col": "status", "diff_cols": ["jr_diff"],
        "note": ("所选期间内的借贷发生额合计" if pcond else "全部期间的借贷发生额合计")
                + "；本币判平衡，原币仅展示（跨币种凭证原币不可相加）",
        "empty_note": "%d 个主体全部借贷平衡，无不平" % len(r1),
    })

    # ---- 2.2 月平衡
    if pexpr:
        sql = """
        SELECT {ent} AS entity, {pe} AS period,
               count(*) AS n_jr_rows,
               {d}  AS jr_debit_sum,
               {c} AS jr_credit_sum,
               {d} - {c} AS jr_diff{fxsel},
               CASE WHEN abs({d} - {c}) > {tol}{fxok} THEN '不平' ELSE '平衡' END AS status
        {frm} GROUP BY 1, 2 ORDER BY 1, 2
        """.format(ent=ent, pe=pexpr, d=d_cny, c=c_cny, fxsel=fx_sel,
                   tol=float(tol), fxok=fx_ok, frm=jr_from)
        c2, r2 = _exec(conn, sql)
        _save(conn, "_recon_jr_month", sql)
        b2 = sum(1 for r in r2 if r[c2.index("status")] == "不平")
        res["parts"].append({
            "key": "month", "title": "2.2 月平衡", "columns": c2, "rows": r2,
            "total": len(r2), "bad": b2, "table": "_recon_jr_month",
            "status_col": "status", "diff_cols": ["jr_diff"],
            "note": "每个会计期间一行",
            "empty_note": "%d 个主体×期间全部借贷平衡，无不平" % len(r2),
        })
    else:
        res["warnings"].append("序时账没有「会计期间」也没有「日期」列，月平衡无法检查。")

    # ---- 2.3 唯一凭证号平衡
    vkey = None
    if "voucher_key" in cols and _n_rows(conn, "journal", "voucher_key"):
        vkey = "nullif(trim(voucher_key), '')"
        res["notes"].append("凭证按映射的「唯一凭证号」列归集")
    elif "vno" in cols and _n_rows(conn, "journal", "vno"):
        bits = [ent]
        if pexpr:
            bits.append(pexpr)
        bits.append("nullif(trim(vno), '')")
        vkey = "concat_ws('-', %s)" % ", ".join(bits)
        res["warnings"].append(
            "序时账没有映射「唯一凭证号」列，本次现场按 %s 拼接凭证主键。"
            "要固定下来，到第 2 步「清洗与映射 → 数据清洗」加一步「生成唯一凭证号」，"
            "再映射到「唯一凭证号」。" % "公司名称+期间+凭证号")
    if vkey:
        fx_base = (", {d} AS jr_debit_fx, {c} AS jr_credit_fx".format(d=d_fx, c=c_fx)
                   if has_fx else "")
        base = """
        SELECT {ent} AS entity, {vk} AS voucher_key, {pe} AS period,
               min(nullif(trim(vno), '')) AS vno,
               count(*) AS n_jr_rows,
               {d}  AS jr_debit_sum,
               {c} AS jr_credit_sum{fxbase}
        FROM "journal" WHERE {vk} IS NOT NULL{jw} GROUP BY 1, 2, 3
        """.format(ent=ent, vk=vkey, pe=(pexpr or "'(无期间)'"),
                   d=d_cny, c=c_cny, fxbase=fx_base, jw=jr_where)
        cond = "abs(jr_debit_sum - jr_credit_sum) > %f" % float(tol)
        fx_out = ("jr_debit_fx, jr_credit_fx, jr_debit_fx - jr_credit_fx AS jr_diff_fx,"
                  if has_fx else
                  "NULL AS jr_debit_fx, NULL AS jr_credit_fx, NULL AS jr_diff_fx,")
        sql = """
        SELECT entity, voucher_key, period, vno, n_jr_rows,
               jr_debit_sum, jr_credit_sum,
               jr_debit_sum - jr_credit_sum AS jr_diff,
               %s
               '不平' AS status
        FROM (%s) g WHERE %s
        ORDER BY abs(jr_debit_sum - jr_credit_sum) DESC, entity, voucher_key
        LIMIT 5000
        """ % (fx_out, base, cond)
        c3, r3 = _exec(conn, sql)
        _save(conn, "_recon_jr_voucher", sql)
        n_v = _scalar(conn, "SELECT count(*) FROM (%s) v" % base, 0) or 0
        n_bad = _scalar(conn, "SELECT count(*) FROM (%s) v WHERE %s" % (base, cond), 0) or 0
        note = "%d 张凭证，%s" % (n_v, ("其中不平 %d 张" % n_bad) if n_bad else "全部借贷平衡")
        if n_bad > len(r3):
            note += "（页面只列前 %d 张，导出为完整清单）" % len(r3)
        res["parts"].append({
            "key": "voucher", "title": "2.3 唯一凭证号平衡", "columns": c3, "rows": r3,
            "total": int(n_v), "bad": int(n_bad), "table": "_recon_jr_voucher",
            "status_col": "status", "diff_cols": ["jr_diff"],
            "note": note + ("；表格只显示借贷不平的凭证" if n_bad else ""),
            "empty_note": ("%d 张凭证全部借贷平衡，没有借贷不平的凭证" % n_v) if n_bad == 0 else "",
        })
    else:
        res["warnings"].append("序时账没有「凭证号」列，凭证级平衡无法检查。"
                               "凭证平衡能定位到具体哪张凭证做错了。")

    res["notes"].insert(0, "序时账 %d 行，按公司名称分组" % total_rows)
    if has_fx:
        res["notes"].append("本币判平衡；原币取 debit_fx/credit_fx 仅展示（跨币种凭证原币不可相加）")
    bad_total = sum(p["bad"] for p in res["parts"])
    res["total"] = sum(p["total"] for p in res["parts"])
    res["bad"] = bad_total
    res["ok"] = bad_total == 0 and res["total"] > 0
    return res

def _pair_check(conn, tol, tb_base, period_from, period_to, key_kind):
    """余额表 × 序时账：按「主体 + 科目」核对发生额（本币 vs 本币）。

    2026-11-11 按用户要求改口径：**不再按币别分组**。
    余额表一侧取本币金额列（本期或本年），序时账一侧取借贷金额——两边本来都是本位币，
    逐科目相加后直接比借贷差即可。原币不再参与本组判定。
    """
    if key_kind == "l1":
        res = {"key": "pair_l1", "no": "三",
               "title": "科目余额表 × 序时账 · 按「公司名称 + 一级科目」核对发生额",
               "desc": "逐公司、逐一级科目比对借贷发生额（均为本位币），不再分币别",
               "error": None, "notes": [], "warnings": [], "parts": []}
    else:
        res = {"key": "pair_full", "no": "四",
               "title": "科目余额表 × 序时账 · 按「公司名称 + 科目全称」核对发生额",
               "desc": "逐公司、逐科目比对借贷发生额（均为本位币），比一级科目更细",
               "error": None, "notes": [], "warnings": [], "parts": []}

    if not _has_table(conn, "tb"):
        res["error"] = "还没有科目余额表，先在第 3 步「合并与检查 → 合并成表」生成"
        return res
    if not _has_table(conn, "journal"):
        res["error"] = "还没有序时账，先在第 3 步「合并与检查 → 合并成表」生成"
        return res

    tb_cols = _cols(conn, "tb")
    jr_cols = _cols(conn, "journal")
    if not tb_cols or not jr_cols:
        res["error"] = "标准表读取失败"
        return res

    # ---- 科目键
    if key_kind == "l1":
        if "acct_l1" not in tb_cols or "acct_l1" not in jr_cols:
            res["error"] = ("核对需要「一级科目」列，但余额表和序时账并未都映射上。请先在第 2 步"
                            "「清洗与映射 → 数据清洗」用「提取一级科目」「生成一级科目」或"
                            "「按编码派生一级科目」（定长编码如 10020101）拆出一级科目列，"
                            "再到同一步的「字段映射」里把它映射到「一级科目」——两张表都要。")
            res["need"] = "l1_mapping"
            return res
        n_tb = _n_rows(conn, "tb", "acct_l1")
        n_jr = _n_rows(conn, "journal", "acct_l1")
        if not (n_tb and n_jr):
            res["error"] = ("余额表有 %d 行、序时账有 %d 行填了一级科目，两边都必须有值才能核对。"
                            "请检查两张表的「提取一级科目 / 生成一级科目 / 按编码派生一级科目」"
                            "是否都生效、映射是否都指到了「一级科目」。"
                            % (n_tb, n_jr))
            res["need"] = "l1_mapping"
            return res
        key_expr = "nullif(trim(acct_l1), '')"
        key_tb, key_jr = key_expr, key_expr
        label_key_cn = "一级科目"
        res["notes"].append("按「一级科目」归集（余额表 %d 行 / 序时账 %d 行）" % (n_tb, n_jr))
    else:
        tb_full = _n_rows(conn, "tb", "acct_full") if "acct_full" in tb_cols else 0
        jr_full = _n_rows(conn, "journal", "acct_full") if "acct_full" in jr_cols else 0
        tb_name = _n_rows(conn, "tb", "acct_name") if "acct_name" in tb_cols else 0
        jr_name = _n_rows(conn, "journal", "acct_name") if "acct_name" in jr_cols else 0
        if not tb_full and not tb_name:
            res["error"] = "余额表没有「科目全称」也没有「科目名称」，无法按科目核对。"
            return res
        if not jr_full and not jr_name:
            res["error"] = "序时账没有「科目全称」也没有「科目名称」，无法按科目核对。"
            return res

        def _key(full_n):
            if full_n:
                return ("coalesce(nullif(trim(acct_full), ''), nullif(trim(acct_name), ''))")
            return "nullif(trim(acct_name), '')"

        key_tb, key_jr = _key(tb_full), _key(jr_full)
        label_key_cn = "科目全称"
        if tb_full and jr_full:
            res["notes"].append("两边都用「科目全称」列（余额表 %d 行 / 序时账 %d 行）" % (tb_full, jr_full))
        else:
            res["warnings"].append(
                "两边「科目全称」没对上：余额表%s、序时账%s，本次按「科目名称」核对。"
                % ("有" if tb_full else "没有", "有" if jr_full else "没有"))

    # ---- 金额列：只取本币
    dcol, ccol = ("ytd_debit", "ytd_credit") if tb_base == "ytd" else ("cur_debit", "cur_credit")
    if dcol not in tb_cols or ccol not in tb_cols:
        res["error"] = ("余额表没有「%s」，无法核对。到「字段映射」把它配上。"
                        % ("本年累计借方/贷方" if tb_base == "ytd" else "本期借方/贷方"))
        return res
    d_, c_ = qi(dcol), qi(ccol)

    pcond = _period_cond(conn, period_from, period_to, jr_cols)
    pw = (" AND " + pcond) if pcond else ""
    if pcond and _n_rows(conn, "journal"):
        n_jr_hit = _scalar(conn, 'SELECT count(*) FROM "journal" WHERE 1=1 %s'
                           % (" AND " + pcond), 0) or 0
        if n_jr_hit == 0 and _n_rows(conn, "tb"):
            res["error"] = ("期间过滤后序时账 0 行。序时账的期间格式和你填的期间不一致——"
                            "数据里可能是「1」~「12」或「2026年1期」，而你填的是「2026-01」（或反之）。")
            return res
        if n_jr_hit:
            res["notes"].append("序时账已按期间过滤：%d / %d 行命中" % (n_jr_hit, _n_rows(conn, "journal")))

    ent_tb = entity_expr(tb_cols)
    ent_jr = entity_expr(jr_cols)
    BL = '"%s"' % label_key_cn

    sql = """
    WITH t AS (
      SELECT {ent_tb} AS company, {key_tb} AS acct,
             sum(coalesce({d}, 0)) AS td, sum(coalesce({c}, 0)) AS tc
      FROM "tb" WHERE {key_tb} IS NOT NULL GROUP BY 1, 2
    ),
    j AS (
      SELECT {ent_jr} AS company, {key_jr} AS acct,
             sum(coalesce(debit, 0)) AS jd, sum(coalesce(credit, 0)) AS jc
      FROM "journal" WHERE {key_jr} IS NOT NULL {p} GROUP BY 1, 2
    ),
    k AS (SELECT company, acct FROM t UNION SELECT company, acct FROM j)
    SELECT coalesce(k.company, '(未标注主体)') AS "公司名称",
           k.acct AS {BL},
           round(coalesce(t.td, 0), 2) AS "余额表借方",
           round(coalesce(t.tc, 0), 2) AS "余额表贷方",
           round(coalesce(j.jd, 0), 2) AS "序时账借方",
           round(coalesce(j.jc, 0), 2) AS "序时账贷方",
           round(coalesce(t.td, 0) - coalesce(j.jd, 0), 2) AS "差异借方",
           round(coalesce(t.tc, 0) - coalesce(j.jc, 0), 2) AS "差异贷方",
           round((coalesce(t.td, 0) - coalesce(t.tc, 0))
                 - (coalesce(j.jd, 0) - coalesce(j.jc, 0)), 2) AS "差异净额",
           CASE
             WHEN t.acct IS NULL AND abs(coalesce(j.jd,0) - coalesce(j.jc,0)) <= {tol} THEN '一致'
             WHEN t.acct IS NULL THEN '仅序时账有'
             WHEN j.acct IS NULL AND abs(coalesce(t.td,0) - coalesce(t.tc,0)) <= {tol} THEN '一致'
             WHEN j.acct IS NULL THEN '仅余额表有'
             WHEN abs(coalesce(t.td,0) - coalesce(j.jd,0)) <= {tol}
              AND abs(coalesce(t.tc,0) - coalesce(j.jc,0)) <= {tol} THEN '一致'
             ELSE '不符' END AS "状态"
    FROM k
    LEFT JOIN t ON t.company IS NOT DISTINCT FROM k.company AND t.acct = k.acct
    LEFT JOIN j ON j.company IS NOT DISTINCT FROM k.company AND j.acct = k.acct
    ORDER BY 1, 2
    """.format(ent_tb=ent_tb, ent_jr=ent_jr, key_tb=key_tb, key_jr=key_jr,
               d=d_, c=c_, p=pw, tol=float(tol), BL=BL)

    columns, rows = _exec(conn, sql)
    tbl_name = "_recon_pair_l1" if key_kind == "l1" else "_recon_pair_full"
    _save(conn, tbl_name, sql)

    si = columns.index("状态")
    di = columns.index("差异净额")
    ok_n = sum(1 for r in rows if r[si] == "一致")
    bad_n = sum(1 for r in rows if r[si] == "不符")
    only_tb = sum(1 for r in rows if r[si] == "仅余额表有")
    only_jr = sum(1 for r in rows if r[si] == "仅序时账有")

    ents = {}
    for r in rows:
        e = r[0] or "(未标注主体)"
        g = ents.setdefault(e, {"entity": e, "n": 0, "ok": 0, "bad": 0,
                                "only_tb": 0, "only_jr": 0, "diff_net": 0.0})
        g["n"] += 1
        st = r[si]
        if st == "一致":
            g["ok"] += 1
        elif st == "不符":
            g["bad"] += 1
        elif st == "仅余额表有":
            g["only_tb"] += 1
        else:
            g["only_jr"] += 1
        g["diff_net"] += (r[di] or 0)
    ent_rows = sorted(ents.values(),
                      key=lambda x: (-(x["bad"] + x["only_tb"] + x["only_jr"]), x["entity"]))
    for g in ent_rows:
        g["diff_net"] = round(g["diff_net"], 2)

    res["notes"].append("口径：余额表「%s」vs 序时账发生额，均为本位币，不分币别"
                        % _base_label(tb_base))
    res["notes"].append("主体 %d 个 · 科目组 %d 个：一致 %d / 不符 %d / 仅余额表有 %d / 仅序时账有 %d"
                        % (len(ent_rows), len(rows), ok_n, bad_n, only_tb, only_jr))
    if len(ent_rows) == 1 and ent_rows[0]["entity"] == "(未标注主体)":
        res["warnings"].append("两张表都没有公司名称/账簿名称，本次按全表合计核对。")

    try:
        tb_net = _scalar(conn, 'SELECT coalesce(sum(coalesce(%s,0)),0) - '
                               'coalesce(sum(coalesce(%s,0)),0) FROM "tb"' % (d_, c_))
        jr_net = _scalar(conn, 'SELECT coalesce(sum(coalesce(debit,0)),0) - '
                               'coalesce(sum(coalesce(credit,0)),0) FROM "journal" '
                               'WHERE 1=1 %s' % ((" AND " + pcond) if pcond else ""))
        if tb_net is not None and jr_net is not None:
            res["notes"].append("口径试算：余额表「%s」净额 %.2f vs 序时账净额 %.2f（差 %.2f）"
                                % (_base_label(tb_base), tb_net, jr_net, jr_net - tb_net))
    except Exception:
        pass

    res["parts"].append({
        "key": key_kind, "title": label_key_cn,
        "columns": columns, "rows": rows, "total": len(rows),
        "bad": bad_n + only_tb + only_jr, "table": tbl_name,
        "status_col": "状态",
        "diff_cols": ["差异借方", "差异贷方", "差异净额"],
        "view_cols": columns,
        "stats": {"ok": ok_n, "bad": bad_n, "only_tb": only_tb, "only_jr": only_jr},
        "entities": ent_rows,
        "empty_note": "没有可核对的科目" if not rows else "",
    })
    res["tb_base"] = tb_base
    res["tb_label"] = _base_label(tb_base)
    res["ok"] = (bad_n + only_tb + only_jr) == 0 and len(rows) > 0
    res["total"] = len(rows)
    res["bad"] = bad_n + only_tb + only_jr
    return res


# ---------------------------------------------------------------- 主流程
# ============================================================ 五、货币资金辅助核算可用性
# 这一组不是"两边对不对"，而是"数据够不够用"：BKD 要的是「主体 × 银行 × 账号」，
# 拿不到账户粒度就退化成一行"银行存款"，函证、对账单、截止性测试全都做不了。
AUX_COL_RX = _re.compile(r"核算维度|辅助核算|辅助项|银行账号|账号", _re.I)
BANK_TXT_RX = r"银行|支行|分行|信用社|营业部|财务公司"
ACCTNO_TXT_RX = r"[0-9]{6,}"


def _money_where(alias=""):
    p = (alias + ".") if alias else ""
    return ("(coalesce(trim(%sacct_l1),'') IN ('库存现金','银行存款','其他货币资金') "
            "OR trim(coalesce(%sacct_code,'')) IN ('1001','1002','1012'))" % (p, p))


def _need_acct_where(alias=""):
    """银行存款 / 其他货币资金才需要账号和开户行；库存现金本来就没有。"""
    p = (alias + ".") if alias else ""
    return ("(coalesce(trim(%sacct_l1),'') IN ('银行存款','其他货币资金') "
            "OR trim(coalesce(%sacct_code,'')) IN ('1002','1012'))" % (p, p))


def _has_aux_where(alias=""):
    """能不能定位到「具体那个账户 / 客商」，不限于辅助核算槽。

    2026-11-11 补 bank_account / bank_name：账套把「核算维度_银行账户」这类列
    直接映射到**专用字段**（bank_account）而不是 aux 槽时，账户信息其实是在的，
    只认 aux1-3 会把 40 行「有账号」误判成「无法定位」。
    实测：10 个账簿表 40 行银行账户全在 bank_account，值为「012-9006-939,宝中堂DBS账户」。
    """
    p = (alias + ".") if alias else ""
    return ("(coalesce(trim(%saux1),'')<>'' OR coalesce(trim(%saux2),'')<>'' "
            "OR coalesce(trim(%saux3),'')<>'' OR coalesce(trim(%saux_text),'')<>'' "
            "OR coalesce(trim(%sbank_account),'')<>'' OR coalesce(trim(%sbank_name),'')<>'')"
            % (p, p, p, p, p, p))


def _name_hint_where(alias=""):
    """科目名称里能否看出**具体**银行或账号。

    注意不能只匹配「银行」二字——「银行存款」这个科目名本身就有「银行」，
    那样每一行都会被算成"能看出对方是谁"。要求满足其一：
      1) 有 6 位以上连续数字（账号）；
      2) 含银行/支行/分行等字样，且名称不是光秃秃的「银行存款」；
      3) **一级科目名之后还有名字** —— 2026-11-11 补：其他货币资金_微信 / _支付宝 /
         _富途保证金户 / _数字人民币 这类第三方支付、证券账户本来就没有银行账号，
         科目名称本身就是账户标识。实测金斯耐 157 行全是 1012.xx（微信 59、支付宝 24、
         通联 22、京东钱包 11、数字人民币 10、富途 20、Pay Now 9…），按旧口径全被判「无法定位」。
    """
    p = (alias + ".") if alias else ""
    name = f"coalesce({p}acct_name,'')"
    full = f"({name}||'|'||coalesce({p}aux_text,''))"
    rest = (f"regexp_replace(trim({name}), "
            f"'^(银行存款|其他货币资金|库存现金)[[:space:]_,.-]*', '')")
    return (f"(regexp_matches({full}, '{ACCTNO_TXT_RX}') "
            f" OR (regexp_matches({full}, '{BANK_TXT_RX}') "
            f"     AND NOT regexp_full_match(trim({name}), '(银行存款|其他货币资金|库存现金)'))"
            f" OR length({rest}) >= 2)")


def _tb_sheets(conn, target="tb"):
    """_sheets 里某个 target 的原始表信息（含映射 JSON），用于判断数据是丢了还是本来没有。"""
    out = []
    if not _has_table(conn, "_sheets"):
        return out
    try:
        rows = conn.execute(
            "SELECT sid, sheet_name, table_name, col_names, mapping, clean FROM _sheets "
            "WHERE target=%s" % _q(target)).fetchall()
    except Exception:
        return out
    for r in rows:
        out.append({"sid": r[0], "sheet": r[1], "table": r[2],
                    "colNames": r[3] or "", "mapping": r[4] or "", "clean": r[5] or ""})
    return out


def _aux_origin(conn, tb_stat, sheets):
    """判断"辅助核算为什么没有"：结果里已经有 / 原始就没有 / 客户没填 / 映射漏配 / 清洗丢了。

    这是本组检查的核心——只看结果数据永远分不清这几种，必须查原始表、映射和剔除留痕。
    """
    evidence = []
    verdict = "unknown"

    # 结果表里已经拿到辅助核算，就不用再找原因了
    if tb_stat and tb_stat.get("with_aux", 0) > 0:
        return "present", [{
            "sheet": "tb", "stage": "辅助核算已生效",
            "detail": "合并后的 tb 里有 %d 行带辅助核算，可直接生成账户级 BKD。"
                      % tb_stat["with_aux"]}]

    for sh in sheets:
        sid = sh["sid"]
        raw = "raw_" + str(sid)
        cols = _cols(conn, raw) if _has_table(conn, raw) else []
        aux_cols = [c for c in cols if AUX_COL_RX.search(str(c))]
        mp_aux = []
        try:
            _mp = json.loads(sh["mapping"] or "{}")
            mp_aux = [k for k in ("aux1", "aux2", "aux3", "bank_account")
                      if _mp.get(k)]
        except Exception:
            mp_aux = []

        dropped = 0
        drop_tbl = "_dropped_" + str(sid)
        if _has_table(conn, drop_tbl):
            dropped = _scalar(conn, 'SELECT count(*) FROM "%s"' % drop_tbl, 0) or 0

        if not cols:
            evidence.append({"sheet": sh["sheet"], "stage": "原始表已不存在",
                             "detail": "无法回溯，只能看当前结果"})
            continue

        if not aux_cols:
            verdict = "never_present"
            evidence.append({
                "sheet": sh["sheet"], "stage": "原始表就没有辅助核算列",
                "detail": "原始表头里没有「核算维度 / 辅助核算 / 银行账号」类列。"
                          "这不是清洗造成的：客户导出时没勾选辅助核算，"
                          "或系统未启用。需客户重新导出，或改用科目名称反解。"})
            continue

        filled = 0
        code_col = next((c for c in cols if "科目编码" in str(c) or str(c).lower() == "acct_code"), None)
        a_col = aux_cols[0]
        if code_col:
            filled = _scalar(conn, 'SELECT count(*) FROM "%s" WHERE CAST(%s AS VARCHAR) LIKE \'100%%\' '
                                   'AND coalesce(trim(CAST(%s AS VARCHAR)),\'\')<>\'\'' % (
                                       raw, qi(code_col), qi(a_col)), 0) or 0
        if not filled:
            verdict = "never_filled"
            evidence.append({
                "sheet": sh["sheet"], "stage": "原始表有该列但全空",
                "detail": "有「%s」列，但货币资金行里一个值都没有。属于客户没填，不是清洗删的。" % a_col})
        elif not mp_aux:
            verdict = "not_mapped"
            evidence.append({
                "sheet": sh["sheet"], "stage": "映射没配辅助核算",
                "detail": "原始表里该列有 %d 行有值，但字段映射没有把任何辅助核算列配到 aux 槽。"
                          "去「字段映射」页补上即可，不用重新导入。" % filled})
        else:
            verdict = "lost"
            evidence.append({
                "sheet": sh["sheet"], "stage": "映射配了但结果为空",
                "detail": "原始表有 %d 行有值、映射也配了（%s），合并后却是空的。"
                          "该表累计剔除 %d 行，请查清洗留痕是不是误删了行或覆盖了列。"
                          % (filled, "、".join(mp_aux), dropped)})
    return verdict, evidence


def _check_aux(conn, tol=0.01):
    """第五组：货币资金辅助核算可用性（含成因判定）。"""
    notes, warnings, parts = [], [], []

    # ---------- A 科目余额表 ----------
    has_tb = _has_table(conn, "tb")
    n_all = need = with_aux = name_ok = need_ok = 0
    if has_tb:
        base = "FROM tb WHERE " + _money_where()
        n_all = _scalar(conn, "SELECT count(*) " + base, 0) or 0
        need = _scalar(conn, "SELECT count(*) " + base + " AND " + _need_acct_where(), 0) or 0
        with_aux = _scalar(conn, "SELECT count(*) " + base + " AND " + _has_aux_where(), 0) or 0
        name_ok = _scalar(conn, "SELECT count(*) " + base + " AND " + _name_hint_where(), 0) or 0
        need_ok = _scalar(conn, "SELECT count(*) " + base + " AND " + _need_acct_where()
                          + " AND (" + _has_aux_where() + " OR " + _name_hint_where() + ")", 0) or 0

    bad_tb = max(0, need - need_ok)
    if not has_tb:
        v_tb = "none"
        notes.append("还没有合并后的科目余额表（tb），无法判断。")
    elif need == 0:
        v_tb = "ok"
        notes.append("科目余额表里没有「银行存款 / 其他货币资金」行，本组不适用。")
    elif need_ok == need:
        v_tb = "ok"
    elif need_ok > 0:
        v_tb = "partial"
    else:
        v_tb = "none"

    parts.append({
        "key": "tb_aux", "title": "科目余额表 · 货币资金账户粒度",
        "columns": ["item", "value", "note"],
        "headers": ["检查项", "数值", "说明"],
        "view_cols": ["item", "value", "note"],
        "rows": [
            ["货币资金账户行数", n_all, "含库存现金 / 银行存款 / 其他货币资金"],
            ["需要账户粒度的行", need, "仅银行存款 + 其他货币资金；现金不需要账号"],
            ["带辅助核算", with_aux, "aux1/aux2/aux3 任一有值"],
            ["科目名称里能看出银行/账号", name_ok, "正则匹配 银行|支行|分行|信用社 或 6 位以上数字"],
            ["能定位到具体账户", need_ok, "辅助核算 或 名称线索，二者之一即可"],
            ["无法定位（会退化成一行）", bad_tb, "这些行在 BKD 里只能合成一行「银行存款」"],
        ],
        "total": need, "bad": bad_tb, "table": None,
        "note": {"ok": "可生成账户级 BKD", "partial": "部分账户无法定位",
                 "none": "只能到科目层"}.get(v_tb, ""),
    })

    # ---------- B 序时账 ----------
    j_rows = j_detail = j_aux = j_sum = 0
    if _has_table(conn, "journal"):
        jw = ("(coalesce(trim(acct_l1),'') IN ('库存现金','银行存款','其他货币资金') "
              "OR trim(coalesce(acct_code,'')) LIKE '100%' "
              "OR coalesce(acct_name,'') LIKE '%银行存款%' "
              "OR coalesce(acct_name,'') LIKE '%库存现金%' "
              "OR coalesce(acct_name,'') LIKE '%其他货币资金%')")
        j_rows = _scalar(conn, "SELECT count(*) FROM journal WHERE " + jw, 0) or 0
        j_detail = _scalar(conn, "SELECT count(*) FROM journal WHERE " + jw
                           + " AND (trim(coalesce(acct_code,'')) LIKE '1001.%' "
                                   "OR trim(coalesce(acct_code,'')) LIKE '1002.%' "
                                   "OR trim(coalesce(acct_code,'')) LIKE '1012.%')", 0) or 0
        j_aux = _scalar(conn, "SELECT count(*) FROM journal WHERE " + jw + " AND " + _has_aux_where(), 0) or 0
        # 注意：这段不要用 % 格式化——上面的 jw 里含 LIKE '100%'，% 会被当成占位符
        j_sum = _scalar(conn, "SELECT count(*) FROM journal WHERE " + jw
                        + " AND regexp_matches(coalesce(summary,''), " + _q(BANK_TXT_RX) + ")", 0) or 0

    if j_rows == 0:
        v_jr = "ok"
        jnote = "序时账里没有货币资金行，本组不适用。"
    elif j_detail > 0 or j_aux > 0:
        v_jr = "ok"
        jnote = "可按账户核对"
    elif j_sum > 0:
        v_jr = "partial"
        jnote = "只能靠摘要里的银行名反推，不精确"
    else:
        v_jr = "none"
        jnote = "只能到科目层，无法按账户核对"

    parts.append({
        "key": "jr_aux", "title": "序时账 · 货币资金明细程度",
        "columns": ["item", "value", "note"],
        "headers": ["检查项", "数值", "说明"],
        "view_cols": ["item", "value", "note"],
        "rows": [
            ["货币资金相关行数", j_rows, "含一级科目与明细科目"],
            ["带明细码（如 1002.01）", j_detail, "有明细码就能按账户核对"],
            ["带辅助核算", j_aux, ""],
            ["摘要里出现银行/支行", j_sum, "最后的兜底线索，不精确"],
        ],
        "total": j_rows, "bad": 0 if v_jr in ("ok",) else j_rows, "table": None,
        "note": jnote,
    })

    # ---------- C 成因判定 ----------
    verdict, evidence = _aux_origin(conn, {"with_aux": with_aux}, _tb_sheets(conn, "tb"))
    if evidence and verdict != "present":
        for e in evidence:
            notes.append("【成因】%s：%s" % (e["stage"], e["detail"]))
    vmap = {"present": "辅助核算可用，无需处理",
            "never_present": "原始数据就没有辅助核算",
            "never_filled": "原始表有列但客户没填",
            "not_mapped": "字段映射漏配（可立即修复，不用重导）",
            "lost": "映射配了但结果为空，疑清洗误删",
            "unknown": "无法判断（缺原始表）"}
    if verdict != "present":
        notes.append("成因判定：%s" % vmap.get(verdict, verdict))

    if v_tb != "ok":
        warnings.append(
            "银行存款 BKD 会退化成一行「银行存款」：%d 行里有 %d 行拿不到银行/账号，"
            "函证、银行对账单、截止性测试都无法按账户执行。" % (need, bad_tb))
    if v_jr == "none" and j_rows > 0:
        warnings.append("序时账只有一级科目，无法按账户核对货币资金流水。")

    ok = (v_tb in ("ok",)) and (v_jr in ("ok", "partial"))
    return {
        "key": "aux", "no": "五", "title": "货币资金辅助核算可用性",
        "ok": ok, "total": need, "bad": bad_tb,
        "parts": parts, "notes": notes, "warnings": warnings,
        "verdict": verdict, "evidence": evidence,
        "auxStat": {"tb": v_tb, "journal": v_jr, "need": need, "ok": need_ok},
    }


def _headers(columns, override=None):
    """核对结果表的导出表头：内部列名走 RECON_LABEL 翻中文。"""
    out = [to_cn(c, "recon") for c in columns]
    if override:
        for i, c in enumerate(columns):
            if c in override:
                out[i] = override[c]
    return out


def _save_rows(conn, name, columns, rows):
    """把结果落成一张表，供页面/导出复用（列全按文本存，避免类型推断出错）。"""
    conn.execute('DROP TABLE IF EXISTS "%s"' % name)
    conn.execute('CREATE TABLE "%s"(%s)' % (
        name, ", ".join('"%s" VARCHAR' % c for c in columns)))
    if rows:
        conn.executemany('INSERT INTO "%s" VALUES (%s)' % (name, ",".join("?" * len(columns))),
                         [[None if v is None else str(v) for v in r] for r in rows])


def _check_entity_scope(conn, tol=0.01):
    """第零组：主体完整性。

    以**挂接的账簿对照表**（_books）为完整主体清单，看三件事：
      1. 哪些主体没有科目余额表
      2. 哪些主体没有序时账（顺带标出「本年发生额为 0」，那是本期没业务，不算错）
      3. 哪些主体只出现在两张标准表里、不在对照表里（超范围）

    放在最前面：主体没对上，后面五组的「按公司+科目核对」结论都不成立。
    """
    res = {"key": "entity_scope", "no": "零",
           "title": "主体完整性 · 对照表 vs 两张标准表",
           "desc": "以账簿对照表为完整主体清单，查谁缺余额表、谁缺序时账、谁不在清单里",
           "error": None, "notes": [], "warnings": [], "parts": []}
    has_tb = _has_table(conn, "tb")
    has_jr = _has_table(conn, "journal")
    if not has_tb and not has_jr:
        res["error"] = "还没有合并后的标准表，先在第 3 步「合并与检查 → 合并成表」生成"
        res["ok"], res["total"], res["bad"] = True, 0, 0
        return res

    def ents(t):
        if not _has_table(conn, t):
            return set()
        return {r[0] for r in conn.execute(
            'SELECT DISTINCT trim(company_name) FROM "%s" '
            "WHERE nullif(trim(coalesce(company_name,'')),'') IS NOT NULL" % t).fetchall()}

    tb_set, jr_set = ents("tb"), ents("journal")
    try:
        book_set = {r[0] for r in conn.execute(
            "SELECT DISTINCT trim(company_name) FROM _books "
            "WHERE nullif(trim(coalesce(company_name,'')),'') IS NOT NULL").fetchall()}
    except Exception:
        book_set = set()

    if book_set:
        universe = book_set
        res["notes"].append("完整主体清单：账簿对照表 %d 个主体" % len(book_set))
    else:
        universe = tb_set | jr_set
        res["warnings"].append("没有挂接账簿对照表——本次以两张标准表出现过的 %d 个主体为全集，"
                               "只能查「缺对方表」，查不出「超范围」" % len(universe))
        res["notes"].append("建议先上传「账簿→公司」对照表，进项才算有基准")

    if not has_tb or not has_jr:
        miss = "科目余额表" if not has_tb else "序时账"
        res["warnings"].append("还没有合并%s，缺表判断不完整" % miss)

    # ---- 1) 缺科目余额表
    rows1 = []
    if has_tb:
        for e in sorted(universe - tb_set):
            jr_n = _scalar(conn, 'SELECT count(*) FROM "journal" WHERE trim(coalesce(company_name,\'\')) = '
                                 + _q(e), 0) if has_jr else 0
            rows1.append([e, jr_n, "只有序时账，没有余额表"])
    # ---- 2) 缺序时账（分两类：真缺 vs 本期本来就没业务）
    rows2 = []
    n_miss_jr_real = 0
    if has_jr:
        for e in sorted(universe - jr_set):
            amt = _scalar(conn, 'SELECT round(coalesce(sum(coalesce(ytd_debit,0) '
                                '+ coalesce(ytd_credit,0)),0),2) '
                                'FROM "tb" WHERE trim(coalesce(company_name,\'\')) = ' + _q(e), 0) if has_tb else None
            if (amt or 0) == 0:
                note = "本期无发生额（新设/停业/空壳），不算错"
            else:
                note = "有余额表发生额，但序时账里没有它的凭证"
                n_miss_jr_real += 1
            rows2.append([e, amt if amt is not None else "", note])
    # ---- 3) 超范围（只在有对照表时有意义）
    rows3 = []
    if book_set:
        for e in sorted((tb_set | jr_set) - book_set):
            where = []
            if e in tb_set:
                where.append("余额表")
            if e in jr_set:
                where.append("序时账")
            rows3.append([e, " + ".join(where), "对照表里没有这个主体名"])

    def part(key, title, cols, rows, tname, empty_note):
        _save_rows(conn, tname, cols, rows)
        return {"key": key, "title": title, "columns": cols, "rows": rows,
                "total": len(rows), "bad": len(rows), "table": tname,
                "status_col": None, "diff_cols": [], "view_cols": cols,
                "headers": cols, "empty_note": empty_note}

    res["parts"].append(part(
        "no_tb", "缺科目余额表的主体",
        ["公司名称", "序时账行数", "说明"], rows1, "_recon_scope_no_tb",
        "清单里每个主体都有余额表"))
    res["parts"].append(part(
        "no_jr", "缺序时账的主体",
        ["公司名称", "余额表本年发生额", "说明"], rows2, "_recon_scope_no_jr",
        "清单里每个主体都有序时账"))
    res["parts"].append(part(
        "out_scope", "不在对照表里的主体",
        ["公司名称", "出现在", "说明"], rows3, "_recon_scope_out",
        "两张表里的主体都在对照表内" if book_set else "未挂接对照表，本项不适用"))

    # bad 只算「真缺」：缺序时账里那些本期发生额为 0 的主体（新设/停业/空壳）照列不误，
    # 但不计为问题——它们本来就没有凭证可导，把它们算错会让这一组永远不可能通过。
    n_bad = len(rows1) + n_miss_jr_real + len(rows3)
    res["notes"].append(
        "缺余额表 %d 个 / 缺序时账 %d 个（其中真缺 %d、本期无业务 %d）/ 超范围 %d 个"
        % (len(rows1), len(rows2), n_miss_jr_real, len(rows2) - n_miss_jr_real, len(rows3)))
    if rows3:
        res["warnings"].append("有 %d 个主体不在对照表里——多半是账簿名没对上，"
                               "到「账簿对照表」补一行，或在推断结果里确认" % len(rows3))
    res["total"] = len(universe)
    res["bad"] = n_bad
    res["ok"] = n_bad == 0
    return res


def run(conn, period_from=None, period_to=None, tol=0.01, tb_base="ytd"):
    """跑完四组检查。任一组的错误都只影响自己，不会让整页报错。"""
    tol = float(tol or 0.01)
    checks = []
    # 第零组：主体完整性。放最前面——主体都没对上，后面五组的「按公司核对」都不成立。
    try:
        checks.append(_check_entity_scope(conn, tol))
    except Exception as e:
        checks.append({"key": "entity_scope", "no": "零", "title": "主体完整性",
                       "error": "检查执行失败：%s" % str(e)[:200],
                       "notes": [], "warnings": [], "parts": [],
                       "ok": False, "total": 0, "bad": 0})
    for fn in (_check_tb_self, _check_jr_self):
        try:
            checks.append(fn(conn, tol, period_from, period_to) if fn is _check_jr_self
                          else fn(conn, tol))
        except Exception as e:
            checks.append({"key": fn.__name__, "no": "?", "title": fn.__name__,
                           "error": "检查执行失败：%s" % str(e)[:200],
                           "notes": [], "warnings": [], "parts": [],
                           "ok": False, "total": 0, "bad": 0})
    for kind in ("l1", "full"):
        try:
            checks.append(_pair_check(conn, tol, tb_base, period_from, period_to, kind))
        except Exception as e:
            checks.append({"key": "pair_" + kind, "no": "?" ,
                           "title": "两表对账",
                           "error": "检查执行失败：%s" % str(e)[:200],
                           "notes": [], "warnings": [], "parts": [],
                           "ok": False, "total": 0, "bad": 0})

    # 第五组：不是勾稽，是「数据够不够用」——货币资金有没有账户粒度
    try:
        checks.append(_check_aux(conn, tol))
    except Exception as e:
        checks.append({"key": "aux", "no": "五", "title": "货币资金辅助核算可用性",
                       "error": "检查执行失败：%s" % str(e)[:200],
                       "notes": [], "warnings": [], "parts": [],
                       "ok": False, "total": 0, "bad": 0})

    # 导出用：(sheet 名, 表名)
    tables = []
    items = []
    for c in checks:
        for p in c.get("parts") or []:
            if not p.get("headers"):
                p["headers"] = _headers(p.get("columns") or [])
            if p.get("table"):
                nm = ("%s-%s" % (c["no"], p["title"])).replace("/", "·")[:31]
                tables.append({"sheet": nm, "table": p["table"],
                               "headers": p.get("headers") or [],
                               "empty_note": p.get("empty_note") or ""})
        items.append({
            "key": c.get("key"), "no": c.get("no"), "title": c.get("title"),
            "ok": bool(c.get("ok")), "total": c.get("total", 0),
            "bad": c.get("bad", 0), "error": c.get("error"),
        })

    failed = sum(1 for i in items if not i["ok"])
    summary = {
        "total": len(items), "passed": len(items) - failed, "failed": failed,
        "bad_items": sum(i["bad"] for i in items),
        "items": items,
    }
    return {"tol": tol, "period": {"from": period_from or "", "to": period_to or ""},
            "tb_base": tb_base or "ytd", "tb_label": _base_label(tb_base or "ytd"),
            "checks": checks, "summary": summary, "tables": tables,
            "warnings": [], "notes": []}
