# -*- coding: utf-8 -*-
"""数据清洗：把「原始表 raw」按可配置的步骤流水线变成「清洗表 stg」。

设计原则
  1. raw 表只读。任何一步都作用在临时表上，最后才落成 stg_<sid>，随时能退回原样。
  2. 每步是一个原子动作，可单独关掉。参数由人工配置，不猜。
  3. 剔除的行不消失：写进 _dropped_<sid>，记录第几步、什么原因，可以随时查回来。
  4. 每步都报影响行数（删了几行、填充了几行），这是审计留痕。

只依赖 duckdb 的 SQL，10 个动作全部是单条 SQL，千万行也能跑。
"""
import re

from . import std_accounts

# ---------------------------------------------------------------- 工具
def qi(name):
    return '"' + str(name).replace('"', '""') + '"'


def ql(s):
    return "'" + str(s).replace("'", "''") + "'"


def _kw_list(text):
    parts = re.split(r"[,，;；\n]+", str(text or ""))
    return [p.strip() for p in parts if p.strip()]


def _like_regex(words):
    return "|".join(re.escape(w) for w in words)


def _data_cols(cols):
    """去掉内部行号列，剩下的才是业务列。"""
    return [c for c in cols if c != "_src_row"]


# ---------------------------------------------------------------- 动作定义
def _b_drop_empty(p, cols, t):
    """按「指定的 N 列」判空。

    mode=all（默认）：这 N 列**全部**为空才删（真正的空行判断，
                      不会因为「摘要空但金额有值」这种正常行被误删）
    mode=any        ：这 N 列**只要有一列**为空就删（用于「科目编码必须填」这类硬性要求）
    """
    use = [c for c in (p.get("cols") or []) if c in cols]
    if not use:
        use = _data_cols(cols)                       # 未选列 = 所有业务列
    if not use:
        return {"kind": "noop", "reason": "表里没有业务列可判空"}
    mode = str(p.get("mode") or "all")
    ne = ["nullif(trim(%s), '') IS NOT NULL" % qi(c) for c in use]
    if mode == "any":
        return {"kind": "filter", "cond": " AND ".join(ne),
                "reason": "所选 %d 列中有空值" % len(use)}
    return {"kind": "filter", "cond": " OR ".join(ne),
            "reason": "所选 %d 列全为空" % len(use)}


def _b_drop_keyword(p, cols, t):
    mode = p.get("mode") or "包含"
    kws = _kw_list(p.get("kw"))
    if not kws:
        return {"kind": "noop", "reason": "未填关键词"}
    col = p.get("col") or ""
    note = ""
    # 填了列名但表里没有这一列：**绝不退化成「全列搜索」**。
    # 2026-11-11 事故：把「记账日期」误写成「日期」，动作不报错，转而把所有列拼起来
    # 搜「2026-10」，任何列（含摘要）出现该字样的行全被删，一次误删 206,244 行。
    # 「不填列名」仍然保留全列搜索（预设里删页脚就是这么用的），
    # 但「填了却找不到」必须停下来告诉人。
    if col and col not in cols:
        cands = [c for c in cols if c != "_src_row" and (col in str(c) or str(c) in col)]
        if len(cands) == 1:
            note = "（列名「%s」按包含匹配到「%s」）" % (col, cands[0])
            col = cands[0]
        elif cands:
            return {"kind": "noop",
                    "reason": "列名「%s」有 %d 个候选（%s），请改选准确的列名"
                              % (col, len(cands), "、".join(str(x) for x in cands[:6]))}
        else:
            return {"kind": "noop",
                    "reason": "列名「%s」不在这张表里，本步未执行。可用列：%s"
                              % (col, "、".join(str(c) for c in cols if c != "_src_row")[:160])}
    pool = _data_cols(cols)
    if not col and not pool:
        return {"kind": "noop", "reason": "表里没有业务列可查"}
    target = qi(col) if col else "concat_ws('|', %s)" % ", ".join(qi(c) for c in pool)
    if mode == "全等":
        cond = "trim(coalesce(%s, '')) NOT IN (%s)" % (target, ", ".join(ql(k) for k in kws))
    else:
        pat = str(p.get("kw")) if mode == "正则" else _like_regex(kws)
        if mode == "正则":
            # 用户手填的正则可能是 [ / (abc / * 这种未闭合写法。不预检的话
            # DuckDB 会在执行期抛错，run() 直接中断，后面所有步骤都不跑了。
            try:
                re.compile(pat)
            except re.error as e:
                return {"kind": "noop", "reason": "正则表达式无效（%s）：%s" % (e, pat[:40])}
        cond = "NOT regexp_matches(coalesce(%s, ''), %s)" % (target, ql(pat))
    return {"kind": "filter", "cond": cond,
            "reason": (p.get("reason") or ("命中 " + "/".join(kws))) + note}


def _b_trim_text(p, cols, t):
    use = p.get("cols") or _data_cols(cols)
    use = [c for c in use if c in cols] or _data_cols(cols)
    exprs = []
    for c in cols:
        if c in use:
            exprs.append(
                "nullif(trim(regexp_replace(translate(coalesce(%s, ''), '　', ' '), '\\s+', ' ', 'g')), '') AS %s"
                % (qi(c), qi(c)))
        else:
            exprs.append(qi(c))
    return {"kind": "select", "sql": "SELECT %s FROM {t}" % ", ".join(exprs)}


def _b_pad_code(p, cols, t):
    col = p.get("col") or ""
    if col not in cols:
        return {"kind": "noop",
                "reason": ("列「%s」不在当前表里" % col) if col else "未选列"}
    n = int(p.get("length") or 0)
    if n <= 0:
        n = int(p.get("_auto_len") or 0)
    if n <= 0:
        return {"kind": "noop", "reason": "未能判断目标长度"}
    exprs = []
    for c in cols:
        if c == col:
            exprs.append(
                "CASE WHEN nullif(trim(%s), '') IS NULL THEN %s "
                "WHEN regexp_full_match(trim(%s), '\\d+') AND length(trim(%s)) < %d "
                "THEN lpad(trim(%s), %d, '0') ELSE trim(%s) END AS %s"
                % (qi(c), qi(c), qi(c), qi(c), n, qi(c), n, qi(c), qi(c)))
        else:
            exprs.append(qi(c))
    return {"kind": "select", "sql": "SELECT %s FROM {t}" % ", ".join(exprs)}


def _b_negate_col(p, cols, t):
    """整列取反：只翻正负号，不新增列、不动文本行。

    用在「流水支出记负、账簿支出记正」方向相反，或红字冲销要转成正数的时候。
    判数字走 try_cast（先去掉千分位），所以「合计」「备注」这类文本行原样留着，
    不会被拼成「-合计」；0 取反没意义，也保持原样。
    """
    col = p.get("col") or ""
    if col not in cols:
        return {"kind": "noop",
                "reason": ("列「%s」不在当前表里" % col) if col else "未选列"}
    mode = p.get("mode") or "flip"
    c = qi(col)
    tc = "trim(%s)" % c
    num = "try_cast(replace(%s, ',', '') AS DOUBLE)" % tc
    # 文本层面翻符号，保留原来的千分位/小数位写法（1,234.00 → -1,234.00）
    body = ("CASE WHEN starts_with({t}, '-') "
            "THEN CASE WHEN substr({t}, 2, 1) = '+' THEN substr({t}, 3) ELSE substr({t}, 2) END "
            "WHEN starts_with({t}, '+') THEN '-' || substr({t}, 2) "
            "ELSE '-' || {t} END").format(t=tc)
    if mode == "to_negative":
        cond = "NOT starts_with(%s, '-')" % tc
    elif mode == "to_positive":
        cond = "starts_with(%s, '-')" % tc
    else:
        cond = "TRUE"
    # 0 取反没意义（会写出 -0），空值和文本一律保持原样
    expr = ("CASE WHEN {n} IS NULL OR {n} = 0 THEN {c} "
            "WHEN {cond} THEN {body} ELSE {c} END AS {c}").format(
            n=num, c=c, cond=cond, body=body)
    exprs = [expr if x == col else qi(x) for x in cols]
    return {"kind": "select", "sql": "SELECT %s FROM {t}" % ", ".join(exprs)}


def _b_fill_down(p, cols, t):
    col = p.get("col") or ""
    if col not in cols:
        return {"kind": "noop", "reason": "未选列"}
    part = p.get("partition") or ""
    # 分区：多主体/多账簿混在一张表时，按分区列分别向下填，避免上一个主体的末值串到下一个主体
    if part and part in cols:
        over_part = "PARTITION BY %s " % qi(part)
        fill_part = "PARTITION BY %s, _g" % qi(part)
    else:
        over_part = ""
        fill_part = "PARTITION BY _g"
    # 保持列的原始顺序：被填充的那一列仍在原位
    outer = ", ".join(("_filled AS %s" % qi(c)) if c == col else qi(c) for c in cols)
    sql = (
        "WITH b AS (SELECT * FROM {{t}}), "
        "d AS (SELECT *, count(nullif(trim({c}), '')) OVER ({op}ORDER BY _src_row "
        "ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) AS _g FROM b), "
        "e AS (SELECT *, nullif(max(nullif(trim({c}), '')) OVER ({fp}), '') AS _filled "
        "FROM d) "
        "SELECT {outer} FROM e ORDER BY _src_row"
    ).format(c=qi(col), outer=outer, op=over_part, fp=fill_part)
    return {"kind": "select", "sql": sql}


def _b_drop_empty_key(p, cols, t):
    """按「关键列」判空删行：账簿/公司/科目编码/科目名称/期间/凭证号 全空才删。

    专门抓页脚行：金额落在某列（如「科目全名」），但关键列全空，普通「删除空行」
    （所有列都空才删）删不掉。这里只看关键列。

    2026-11-11 改：原来只认 10 个写死的中文列名，英文/繁体表头（`总帐科目`、
    `Cl. Bal.`）一个都不命中时会**静默退化成「所有列都空才删」**——看起来跑了，
    实际一行没删。现在改成包含匹配（补繁体与英文写法），仍认不出就明确报 noop。
    """
    _KEY = ("账簿", "公司名称", "账簿名称", "科目编码", "科目代码", "科目名称",
            "总账科目", "总帐科目", "期间", "会计期间", "凭证号", "凭证字号",
            "accountcode", "account code", "acct", "voucher", "book name", "period")
    use = []
    for c in cols:
        if c == "_src_row" or c in use:
            continue
        n = str(c).replace(" ", "").replace("　", "").lower()
        if any(k.replace(" ", "") in n for k in _KEY):
            use.append(c)
    if not use:
        return {"kind": "noop",
                "reason": "没认出关键列（账簿 / 科目编码 / 凭证号 / 期间 / 日期），本步未执行。"
                          "可选：改用「删除空行」并手动勾选列，或「按关键词删行」指定列"}
    ne = ["nullif(trim(%s), '') IS NOT NULL" % qi(c) for c in use]
    return {"kind": "filter", "cond": " OR ".join(ne),
            "reason": "关键列（%s）全为空" % "、".join(use[:4])}


def _b_drop_subtotal(p, cols, t):
    """删小计/合计/总计行。

    判定收得比「格子里出现关键词」严：关键词必须**在格子开头或结尾**。
    只看子串会误删「管理费用—其他合计性支出」「结转本年合计支出」这类真实明细，
    而删错一行比漏删一行危险得多（漏删能在留痕里看见，误删是静默丢数）。
    「本月合计」「本页合计」「合计：2,405,000」都能命中，中间夹字的不算。
    """
    use = [c for c in (p.get("cols") or []) if c in cols]
    if not use:
        use = [c for c in cols if c != "_src_row"]
    if not use:
        return {"kind": "noop", "reason": "表里没有业务列可查"}
    _hit = (r"(合计|小计|总计)\s*$|^\s*(合计|小计|总计)")
    conds = ["regexp_matches(trim(coalesce(%s, '')), '%s')" % (qi(c), _hit) for c in use]
    # cond 在 run() 里是「保留条件」（CREATE TABLE AS SELECT ... WHERE cond），
    # 所有 filter 类动作都必须返回「留下的行」。少这个 NOT 会变成只留合计行、删光明细。
    return {"kind": "filter", "cond": "NOT (%s)" % " OR ".join(conds),
            "reason": "命中合计/小计/总计关键词"}


def _b_drop_base_cur(p, cols, t):
    """删「综合本位币」类汇总行。

    **为什么必须删**：多币别账套导出的余额表里，「综合本位币」行是各外币折算后的
    合计视图，与「人民币」「美元」这些具体币别行**并存**。两者加起来就是把同一笔钱
    算两遍。瑞云实测：本币合计 14.23 亿，若把汇总行也算进去会变成 28.46 亿。

    而且这一行本身「看不出具体币别」—— 外币金额在它里面已经被折平了，留着它对
    原币核对没有任何用处，只会让每一处下游都得再判一次「这行算不算数」。
    清洗阶段删掉，下游就干净了。

    **保护**：如果整张表都是综合本位币行（客户只导出了汇总视图、没有具体币别行），
    删完就一行不剩。这种情况保留不删 —— 宁可让人在表里看见，也不能把表清空。
    """
    col = p.get("col") or _find(cols, "币别", "币种", "货币", "currency", "Currency")
    if not col or col not in cols:
        return {"kind": "noop", "reason": "表里没有币别列，不需要处理"}
    words = ("综合本位币", "本位币", "记账本位币", "本币", "人民币本位币")
    is_base = "trim(coalesce(%s, '')) IN (%s)" % (qi(col), ", ".join(ql(w) for w in words))
    conn = p.get("_conn")
    if conn is not None:
        left = conn.execute("SELECT count(*) FROM %s WHERE NOT (%s)"
                            % (qi(t), is_base)).fetchone()[0]
        if left == 0:
            return {"kind": "noop",
                    "reason": "整表都是综合本位币行，删掉就没有数据了，保留"}
    return {"kind": "filter", "cond": "NOT (%s)" % is_base,
            "reason": "综合本位币汇总行（与具体币别行重复，且看不出具体币别）"}


def _b_keep_base_cur(p, cols, t):
    """只保留「综合本位币」行（与「删除综合本位币行」互斥的另一半选择）。

    什么时候用它：客户只导出了汇总视图、没有分币别明细行——此时按明细上数根本
    取不到账户；或者项目组明确决定按汇总层上数（底稿将没有原币/折算列）。

    为什么必须由清洗阶段显式选，而不是底稿后台自动降级：
    综本行与原币明细行是"汇总/明细"关系，取哪一层决定底稿是"有原币"还是"只有
    折算值"，这是口径选择，必须留痕在清洗配方里，不能藏在取数代码里替人决定。

    保护：整表都没有综本行时 noop —— 否则一执行就把表清空了。
    """
    col = p.get("col") or _find(cols, "币别", "币种", "货币", "currency", "Currency")
    if not col or col not in cols:
        return {"kind": "noop", "reason": "表里没有币别列，不需要处理"}
    words = ("综合本位币", "本位币", "记账本位币", "本币", "人民币本位币")
    is_base = "trim(coalesce(%s, '')) IN (%s)" % (qi(col), ", ".join(ql(w) for w in words))
    conn = p.get("_conn")
    if conn is not None:
        n_base = conn.execute("SELECT count(*) FROM %s WHERE %s"
                             % (qi(t), is_base)).fetchone()[0]
        if n_base == 0:
            return {"kind": "noop",
                    "reason": "整表没有综合本位币行，此步不执行（避免把表清空）"}
    return {"kind": "filter", "cond": is_base,
            "reason": "只保留综合本位币层（按汇总层上数，底稿无原币明细列）"}


def _b_fill_const(p, cols, t):
    """新增一列，整列填同一个固定值。

    用途：余额表里「公司名称」「币种」常常不是一列数据，而是整份表共用一个值
    （抬头里写着「XX公司 2026年10月」、全是人民币）。用这步把它们补成标准列，
    后面映射就能直接对上。
    """
    name = str(p.get("name") or "").strip()
    if not name:
        return {"kind": "noop", "reason": "未填列名"}
    val = p.get("value")
    val = "" if val is None else str(val)
    if not val.strip():
        return {"kind": "noop", "reason": "未填取值"}
    if name in cols:
        exprs = ["%s AS %s" % (ql(val), qi(c)) if c == name else qi(c) for c in cols]
    else:
        exprs = [qi(c) for c in cols] + ["%s AS %s" % (ql(val), qi(name))]
    return {"kind": "select", "sql": "SELECT %s FROM {t}" % ", ".join(exprs)}


def _b_derive_level(p, cols, t):
    col = p.get("col") or ""
    if col not in cols:
        return {"kind": "noop", "reason": "未选列"}
    new = p.get("name") or "科目级次"
    if new in cols:
        new = new + "_推断"
    sql = (
        "WITH b AS (SELECT * FROM {{t}}), "
        "lv AS (SELECT L, row_number() OVER (ORDER BY L) AS lv FROM "
        "(SELECT DISTINCT length(trim({c})) AS L FROM b WHERE nullif(trim({c}), '') IS NOT NULL)) "
        "SELECT b.*, lv.lv::VARCHAR AS {n} FROM b LEFT JOIN lv ON lv.L = length(trim(b.{c})) "
        "ORDER BY b._src_row"
    ).format(c=qi(col), n=qi(new))
    return {"kind": "select", "sql": sql, "adds": [new]}


def _b_add_l1(p, cols, t):
    """从一列里按分隔符取第一段，作为「一级科目」。

    余额表常把编码和名称挤在一列：「1001-库存现金」「1001 库存现金」
    「管理费用/办公费」「1001·库存现金」。取第一个分隔符之前那一段，
    就得到 1001 或 管理费用——正是核对时要按它归集的一级科目。
    """
    col = p.get("col") or ""
    if col not in cols:
        return {"kind": "noop", "reason": "未选列"}
    new = str(p.get("name") or "").strip() or "一级科目"
    if new in cols:
        new = new + "_提取"
    seps = str(p.get("sep") or "") or "-_/ .:：|、（）()　·"
    # 放进正则字符类前，把有特殊含义的字符转义
    chars = "".join("\\" + ch if ch in "\\^-][]" else ch for ch in seps)
    pat = "^([^%s]*)" % chars
    e = "nullif(trim(regexp_extract(trim(%s), %s, 1)), '')" % (qi(col), ql(pat))
    exprs = [qi(c) for c in cols] + ["%s AS %s" % (e, qi(new))]
    return {"kind": "select", "sql": "SELECT %s FROM {t}" % ", ".join(exprs),
            "adds": [new]}


def _uniq_col(base, cols):
    """新列名防重：与 add_acct_path 的 `_提取` 不同，这里统一用 2/3/4 后缀，
    前端 availCols 必须用同一套规则，否则下拉里会出现后端根本产出不了的列名。"""
    base = str(base or "").strip() or "一级科目"
    if base not in cols:
        return base
    k = 2
    while "%s%d" % (base, k) in cols:
        k += 1
    return "%s%d" % (base, k)


def _b_derive_l1(p, cols, t):
    """按科目编码前 N 位派生「一级代码 / 一级科目」两列（只派生，不删任何行）。

    为什么单开一个动作而不改 add_l1：add_l1 的语义是「按分隔符取第一段」，已经写进
    方案库和一堆老配方；定长编码（10020101）是另一套规则，改它会动到老账套的输出。

    名称来源三级兜底：内置标准科目表 → 本表同前缀那一行的名称 → 编码本身。
    """
    col = p.get("col") or ""
    if not col:
        return {"kind": "noop", "reason": "未选科目编码列"}
    if col not in cols:
        return {"kind": "noop", "reason": "列「%s」不在当前表里" % col}

    try:
        n = int(p.get("length") or 0)
    except (TypeError, ValueError):
        n = 0
    if n <= 0:
        n = int(p.get("_auto_l1_len") or 0)
    if n <= 0:
        return {"kind": "noop",
                "reason": p.get("_auto_l1_reason")
                or "未能推断一级编码位数，请手动填写（8 位明细码通常填 4 位）"}

    name_col = p.get("name_col") or ""
    if name_col and name_col not in cols:
        name_col = ""
    out_code = _uniq_col(p.get("out_code") or "一级代码", cols)
    out_name = _uniq_col(p.get("out_name") or "一级科目", cols)

    std = p.get("_std")
    if std is None:                     # 单测直接调 builder 时的兜底
        std = std_accounts.load(p.get("_data_dir"))

    conn = p.get("_conn")
    sub = std
    if conn is not None and len(std) > 500:
        # 覆盖表可能有几千条，只把「数据里真的出现的前缀」编进 SQL，别把整张表塞进去
        try:
            need = {r[0] for r in conn.execute(
                "SELECT DISTINCT substr(regexp_extract(trim(coalesce(%s, '')), '[0-9]+', 0), 1, %d) "
                "FROM %s WHERE %s <> ''" % (qi(col), n, qi(t), qi(col))).fetchall()}
            sub = {k: v for k, v in std.items() if k in need}
        except Exception:
            sub = std

    # 前缀：先取「第一段连续数字」（兼容 1002.01.01 / 1002 银行存款），再截前 N 位；
    # 本身就不足 N 位的整段保留（1002 在 8 位表里照样是 1002）
    dig = "regexp_extract(trim(coalesce({c}, '')), '[0-9]+', 0)".format(c=qi(col))
    pref = ("CASE WHEN {d} = '' THEN NULL "
            "WHEN length({d}) <= {n} THEN {d} ELSE substr({d}, 1, {n}) END").format(d=dig, n=n)

    nq = ("nullif(trim(coalesce(%s, '')), '')" % qi(name_col)) if name_col else None
    has_row = "_src_row" in cols
    cte = ("b AS (SELECT * FROM {t}), "
           "sa AS (SELECT %s AS m), "
           "d AS (SELECT *, %s AS _dg, %s AS _l1c FROM b)" % (std_accounts.sql_map(sub), dig, pref))
    if nq:
        # 同前缀取名：只认「编码正好等于这个前缀」的那一行——那才是真的一级科目行。
        # 拿 22110101「基本工资」给 2211 当名字会写出「基本工资」，比留空更难发现。
        cte += (", nm AS (SELECT _l1c, arg_min(%s, [%s length(%s)]) AS _dname "
                "FROM d WHERE _l1c IS NOT NULL AND _dg = _l1c AND %s IS NOT NULL GROUP BY 1)"
                % (nq, "_src_row," if has_row else "", qi(col), nq))
    else:
        cte += ", nm AS (SELECT CAST(NULL AS VARCHAR) AS _l1c, CAST(NULL AS VARCHAR) AS _dname)"
    join = " FROM d CROSS JOIN sa LEFT JOIN nm ON nm._l1c = d._l1c"

    note = "按前 %d 位派生" % n
    if conn is not None:
        try:
            r = conn.execute(
                ("WITH " + cte + " SELECT count(*), count(d._l1c), count(DISTINCT d._l1c), "
                 "count(*) FILTER (WHERE d._l1c IS NOT NULL AND sa.m[d._l1c] IS NULL "
                 "AND nm._dname IS NULL), "
                 "count(*) FILTER (WHERE d._l1c IS NOT NULL AND sa.m[d._l1c] IS NULL "
                 "AND nm._dname IS NOT NULL)" + join).replace("{t}", qi(t))).fetchone()
        except Exception:
            r = None
        if r:
            n_all, n_code, n_l1, n_fb, n_dt = [int(x or 0) for x in r]
            if n_code == 0:
                return {"kind": "noop",
                        "reason": "所选列里没有可用的科目编码（要求以数字开头，如 10020101、1002.01）；"
                                  "请确认选中的是科目编码列，且没有被前面的步骤清空"}
            note = ("按前 %d 位派生：%d 个一级编码；标准科目表命中 %d 行，数据同前缀补名 %d 行，"
                    "按编码兜底 %d 行" % (n, n_l1, n_code - n_fb - n_dt, n_dt, n_fb))
            if n_all > n_code:
                note += "；另有 %d 行没有数字编码（如合计行），两列留空" % (n_all - n_code)
            if n_fb:
                try:
                    samp = [x for (x,) in conn.execute(
                        ("WITH " + cte + " SELECT DISTINCT d._l1c" + join +
                         " WHERE d._l1c IS NOT NULL AND sa.m[d._l1c] IS NULL "
                         "AND nm._dname IS NULL LIMIT 3").replace("{t}", qi(t))).fetchall()]
                    if samp:
                        note += "（如 %s，可补进 data/std_accounts.json）" % "、".join(samp)
                except Exception:
                    pass

    # 原列一律加 d. 前缀：CTE 里多了 sa/nm，不带前缀会 Ambiguous reference
    orig = ", ".join("d.%s" % qi(c) for c in cols)
    order = " ORDER BY d._src_row" if has_row else ""
    sql = ("WITH " + cte +
           " SELECT %s, d._l1c AS %s, coalesce(sa.m[d._l1c], nm._dname, d._l1c) AS %s"
           % (orig, qi(out_code), qi(out_name)) + join + order)
    return {"kind": "select", "sql": sql, "adds": [out_code, out_name], "note": note}


def _pure_parent_ids(conn, t, ent_expr, code_expr, amt_expr, sep, tol):
    """找出「纯汇总父行」的 _src_row 集合（自下而上递归判定）。

    为什么不能只看「编码是别人的前缀」：父行可能带着自己的钱。
      560406_BM00003  1,941.13   ← 自己的残值
      560406_BM00003_SP00002  134,505.32
    这两行里父行金额 ≠ 子行之和，父行必须保留，否则 1,941.13 凭空消失。

    也不能简单比「所有后代之和」：三层结构会把中间层重复加一遍
      （560406 已含孙辈，再加孙辈就多算），结果真正的纯汇总父行反被判成有残值。

    正确做法是自下而上算「这一行贡献给上级多少」：
      叶子        → 贡献 = 自身金额
      中间层      → 先看自身金额是否 ≈ 各子行贡献之和
                    相等 = 纯汇总，贡献 = 自身金额（= 子行之和），该行删
                    不等 = 有残值，保留，贡献 = 自身金额 + 子行贡献之和
    子行比父行编码长，所以按编码长度倒序处理即可保证先算子后算父。
    """
    if not conn or not sep:
        return []
    try:
        rows = conn.execute(
            "SELECT %s AS e, %s AS c, %s AS a, _src_row FROM %s" % (
                ent_expr, code_expr, amt_expr, qi(t))).fetchall()
    except Exception:
        return []

    per = {}          # (entity, code) -> [金额合计, [src_row]]
    for e, c, a, rid in rows:
        if c is None or c == "":
            continue
        cur = per.setdefault((e, c), [0.0, []])
        try:
            cur[0] += float(a or 0)
        except Exception:
            pass
        cur[1].append(rid)

    by_ent = {}
    for (e, c) in per:
        by_ent.setdefault(e, set()).add(c)

    def anc_candidates(code):
        """可能成为父编码的候选：按分隔符逐级前缀 + 纯数字前缀（用友/金蝶常见）。"""
        out = []
        parts = str(code).split(sep)
        for i in range(1, len(parts)):
            out.append(sep.join(parts[:i]))
        m = re.match(r"^(\d+)", str(code))
        if m:
            digits = m.group(1)
            for i in range(1, len(digits)):
                out.append(digits[:i])
        return out

    contrib, pure = {}, []
    for e, codes in by_ent.items():
        kids = {}
        for c in codes:
            best = None
            for a in anc_candidates(c):
                if a in codes and a != c and (best is None or len(a) > len(best)):
                    best = a
            if best:
                kids.setdefault(best, []).append(c)
        for c in sorted(codes, key=lambda x: len(str(x)), reverse=True):
            amt_c = per[(e, c)][0]
            ch = kids.get(c) or []
            if not ch:
                contrib[(e, c)] = amt_c
                continue
            s = sum(contrib.get((e, k), per[(e, k)][0]) for k in ch)
            if abs(amt_c - s) <= tol:
                contrib[(e, c)] = amt_c        # 纯汇总：金额就是子行之和
                pure.append((e, c))
            else:
                contrib[(e, c)] = amt_c + s    # 带残值：父行自己的钱 + 子行
    out = []
    for key in pure:
        out.extend(per[key][1])
    return out


def _b_add_acct_path(p, cols, t):
    """生成一级科目 + 科目全称，并按主体删非末级行。

    一次完成四件事（顺序固定，且都以「主体」为单位独立进行）：
      1. 分列：科目编码按分隔符取第一段 → 新列「一级编码」
      2. 填名称：每个「主体 + 一级编码」最靠上那行的名称 → 新列「一级科目」
      3. 拼全称：逐级拼接各级标准名（拼接符可设，默认 _）→ 新列「科目全称」
      4. 删非末级：编码被同主体更长编码当作前缀的，是非末级，删掉。

    多账簿余额表的「汇总行 vs 维度明细行」另有一套判断（参数 aux + detail_first）：
    科目编码/名称只写在「所有账簿」层，向下填充后同一科目下会有
      ① 账簿科目汇总行（核算维度为空，金额 = 该科目明细合计）
      ② 若干维度明细行（核算维度 = 银行账号 / 客商，金额 = 单个维度）
    这两类编码相同，只看编码分不出来。开了 detail_first 后按组判断：
      - 组内有维度明细 → 只留明细，删掉汇总行（否则同科目重复计一次）
      - 组内没有明细 → 留 1 行科目行（否则整个科目消失，金额缺口）
    分组键默认「主体 + 币别」，避免「美元有明细」把「欧元只有汇总」的行误删。

    不传 aux / detail_first 时，行为与之前完全一致（老配方不受影响）。
    """
    code_col = p.get("code") or ""
    name_col = p.get("name") or ""
    entity_col = p.get("entity") or ""
    if code_col not in cols:
        return {"kind": "noop", "reason": "未选科目编码列"}
    if name_col not in cols:
        return {"kind": "noop", "reason": "未选科目名称列"}
    # 主体列可以不选：单主体账套把整张表当成一个主体，其余逻辑完全不变。
    single_entity = entity_col not in cols
    sep = p.get("sep") or "."
    join = p.get("join") or "_"
    amount_col = p.get("amount") or ""
    has_amount = bool(amount_col and amount_col in cols)

    # ---- 核算维度明细优先（多账簿 TB 的关键，不传则走老逻辑）
    aux_raw = p.get("aux") or []
    if isinstance(aux_raw, str):
        aux_raw = [aux_raw]
    aux_cols = [c for c in aux_raw if c in cols]
    # 参数在 UI 上是「关/开」下拉，取值是字符串 "0"/"1"，不能直接 bool("0")
    _df = p.get("detail_first")
    detail_first = str(_df).strip().lower() in ("1", "true", "on", "yes") and bool(aux_cols)

    grp_raw = p.get("group_cols") or []
    if isinstance(grp_raw, str):
        grp_raw = [grp_raw]
    group_cols = [c for c in grp_raw
                  if c in cols and c not in (code_col, name_col, entity_col)]
    if not group_cols:
        # 没指定就自动找币别列：多币种表若不按币别分组，
        # 「美元有明细」会把「欧元只有汇总」的行一起删掉
        for c in cols:
            if ("币别" in c) or ("币种" in c) or ("货币" in c) or c.lower() == "currency":
                group_cols = [c]
                break

    code = "nullif(trim(coalesce(%s, '')), '')" % qi(code_col)
    name = "nullif(trim(coalesce(%s, '')), '')" % qi(name_col)
    if single_entity:
        ent = ql("(全部)")
    else:
        ent = "nullif(trim(coalesce(%s, '')), '')" % qi(entity_col)
    # 核算维度：任一列非空即视为明细行（多选时按 coalesce 取第一个非空）
    aux = ("coalesce(%s)" % ", ".join("nullif(trim(coalesce(%s, '')), '')" % qi(c)
                                      for c in aux_cols)) if aux_cols else "NULL::VARCHAR"
    # 分组键：主体 + 编码 + 币别等额外列；没有额外列时用空串占位，保证 JOIN 成立
    gkey = ("concat_ws('|', %s)" % ", ".join("nullif(trim(coalesce(%s, '')), '')" % qi(c)
                                             for c in group_cols)) if group_cols else ql("")

    l1_code = "一级编码"
    l1_name = "一级科目"
    full = "科目全称"
    n = 2
    for nm in (l1_code, l1_name, full):
        while nm in cols:
            nm = nm + str(n)
            n += 1

    # 末级判断分两步：
    #   ① 有更长编码以本编码为前缀 → 是本行的「后代」，本行不是叶子
    #   ② 但「不是叶子」不等于「该删」：父行可能带自身残值。
    #      例：560406 研发费用_测试费 136,446.45 = 1,941.13 + 134,505.32（纯汇总，该删）
    #          560406_BM00003 行政部 1,941.13 ≠ 子行 134,505.32（父行 + 自身金额，不能删）
    #      所以只有「父行金额 ≈ 所有叶子后代金额之和」时才判为可删的纯汇总行。
    tol = p.get("parent_tol")
    try:
        tol = abs(float(tol)) if str(tol).strip() != "" else 0.01
    except Exception:
        tol = 0.01
    amt = ("_num(%s)" % qi(amount_col)) if has_amount else "0"

    # 保留规则三选一：
    #   明细优先 → 有明细只留明细，没明细留一行科目行
    #   老配方+有金额 → 同码保留金额绝对值最大那行（与之前一致）
    #   老配方+无金额 → 同码全留（与之前一致，不自动删）
    if detail_first:
        keep_pred = ("AND ((_has_detail = 1 AND _aux IS NOT NULL) "
                     "OR (_has_detail = 0 AND _aux IS NULL AND _rn = 1))")
    elif has_amount:
        keep_pred = "AND _rn = 1"
    else:
        keep_pred = ""

    # 行集分析部分（base → grp → ranked），主查询和剔除记录共用
    # 父行该不该删，先自下而上算好（见 _pure_parent_ids 的说明），再拼进 SQL。
    # 纯 SQL 里做这个递归要按层迭代，容易写错；行数就几千，Python 里算更快也更清楚。
    _conn = p.get("_conn")
    if _conn is None:
        return {"kind": "noop",
                "reason": "内部错误：缺少数据库连接，无法判断末级科目（本步未执行）"}
    del_ids = _pure_parent_ids(_conn, t, ent, code, amt, sep, float(tol))
    nonleaf_expr = (("CASE WHEN b._src_row IN (%s) THEN 1 ELSE 0 END"
                     % ",".join(str(int(x)) for x in del_ids)) if del_ids else "0")

    bcte = (
        "base AS (SELECT *, %s AS _code, %s AS _name, %s AS _entity, "
        "%s AS _aux, %s AS _gkey, %s AS _amt FROM {t} "
        "WHERE %s IS NOT NULL AND %s IS NOT NULL), "
        "grp AS (SELECT _entity, _code, _gkey, "
        "  max(CASE WHEN _aux IS NOT NULL THEN 1 ELSE 0 END) AS _has_detail "
        "  FROM base GROUP BY 1, 2, 3), "
        "ranked AS (SELECT b.*, gr._has_detail, "
        "  %s AS _nonleaf, "
        "  row_number() OVER (PARTITION BY b._entity, b._code, b._gkey "
        "    ORDER BY %s abs(coalesce(%s, 0)) DESC, b._src_row) AS _rn "
        "  FROM base b JOIN grp gr ON gr._entity = b._entity "
        "    AND gr._code = b._code AND gr._gkey = b._gkey)"
        % (code, name, ent, aux, gkey, amt, code, name, nonleaf_expr,
           ("CASE WHEN b._aux IS NOT NULL THEN 0 ELSE 1 END, " if detail_first else ""),
           amt))

    # 全称部分：各级标准名拼接，相邻重复段跳过（1002_1002 不该拼成「银行存款_银行存款」）
    fcte = (
        "named AS (SELECT _entity, _code, min_by(_name, _src_row) AS _seg_name "
        "  FROM base GROUP BY _entity, _code), "
        "l1 AS (SELECT _src_row, first_value(_name) OVER (PARTITION BY _entity, "
        "  list_extract(string_split(_code, %s), 1) ORDER BY _src_row) AS _l1_name FROM base), "
        "exp AS (SELECT _src_row, _entity, _code, _name, g.k, "
        "  len(string_split(_code, %s)) AS _maxk, "
        "  array_to_string(list_slice(string_split(_code, %s), 1, g.k), %s) AS _pref "
        "  FROM base, LATERAL generate_series(1, len(string_split(_code, %s))) g(k)), "
        "j AS (SELECT exp._src_row, exp.k, nm._seg_name "
        "  FROM exp LEFT JOIN named nm "
        "    ON nm._entity = exp._entity AND nm._code = exp._pref), "
        "segd AS (SELECT _src_row, k, _seg_name, "
        "  lag(_seg_name) OVER (PARTITION BY _src_row ORDER BY k) AS _prev FROM j), "
        "fulln AS (SELECT base._src_row, "
        "  CASE WHEN nullif(trim(base._name), '') IS NULL THEN agg._full "
        # 科目名称本身就是完整路径（含分节符）时直接用它：
        # 「其他业务收入_其他」不该再前置父名变成「其他业务收入_其他业务收入_其他」
        "       WHEN strpos(coalesce(nm._seg_name, ''), %s) > 0 "
        "         OR strpos(coalesce(nm._seg_name, ''), %s) > 0 THEN nm._seg_name "
        "       WHEN nm._seg_name IS NULL "
        "         OR nullif(trim(base._name), '') = nm._seg_name THEN agg._full "
        "       ELSE concat(agg._full, %s, base._name) END AS _full "
        "  FROM base "
        "  JOIN (SELECT _src_row, string_agg(CASE WHEN _seg_name IS NULL "
        "          OR _seg_name = _prev THEN NULL ELSE _seg_name END, %s ORDER BY k) AS _full "
        "        FROM segd GROUP BY _src_row) agg ON agg._src_row = base._src_row "
        "  LEFT JOIN named nm ON nm._entity = base._entity AND nm._code = base._code)"
        % (ql(sep), ql(sep), ql(sep), ql(sep), ql(sep),
           ql(join), ql(sep), ql(join), ql(join)))

    orig = ["ranked._src_row" if c == "_src_row" else qi(c) for c in cols]
    sql = ("WITH " + bcte + ", " + fcte + " "
           "SELECT %s, list_extract(string_split(_code, %s), 1) AS %s, "
           "       l1._l1_name AS %s, fulln._full AS %s "
           "FROM ranked "
           "JOIN l1 ON l1._src_row = ranked._src_row "
           "LEFT JOIN fulln ON fulln._src_row = ranked._src_row "
           "WHERE _nonleaf = 0 %s "
           "ORDER BY ranked._src_row"
           % (", ".join(orig), ql(sep), qi(l1_code), qi(l1_name), qi(full), keep_pred))

    # 剔除记录分三类，便于用户一眼看出"删的是汇总行还是真实明细"
    drops = [("非末级科目行",
              "WITH " + bcte + " SELECT _src_row FROM ranked WHERE _nonleaf = 1")]
    if detail_first:
        drops.append(("有明细的汇总行",
                      "WITH " + bcte + " SELECT _src_row FROM ranked "
                      "WHERE _nonleaf = 0 AND _has_detail = 1 AND _aux IS NULL"))
        drops.append(("无明细科目的同码重复行",
                      "WITH " + bcte + " SELECT _src_row FROM ranked "
                      "WHERE _nonleaf = 0 AND _has_detail = 0 AND _aux IS NULL AND _rn > 1"))
    elif has_amount:
        drops.append(("同码重复行",
                      "WITH " + bcte + " SELECT _src_row FROM ranked "
                      "WHERE _nonleaf = 0 AND _rn > 1"))

    notes = []
    if single_entity:
        notes.append("未选主体列，整张表按一个主体判断末级")
    if detail_first:
        notes.append("已开启核算维度明细优先：有明细的科目只留明细，无明细的科目保留一行")
        notes.append("若某科目的明细合计与原汇总行不一致，会体现在后面的余额平衡检查里，"
                     "不会自动调平")
        if not group_cols:
            if grp_raw:
                notes.append("分组列与主体列重复，已按主体分组")
            else:
                notes.append("未选分组列且表里没有币别列，多币种表请手动把币别列入分组")
    elif not has_amount:
        notes.append("未选金额列，同码多行全部保留（如需去重请选金额列，或用「按关键词删行」删合计/占位行）")
    elif aux_cols and not detail_first:
        notes.append("已选核算维度列但未开启「明细优先」，仍按金额最大保留一行")
    reason = ("非末级科目行或同码重复行" if not detail_first else
              "非末级科目行 / 有明细的汇总行 / 同码重复行")
    return {"kind": "select", "sql": sql, "drops": drops,
            "adds": [l1_code, l1_name, full],
            "reason": reason, "note": "；".join(notes)}


def _b_split_aux(p, cols, t):
    col = p.get("col") or ""
    if col not in cols:
        return {"kind": "noop", "reason": "未选列"}
    sep = p.get("sep") or "/"
    names = []
    for i in (1, 2, 3):
        nm = (p.get("name%d" % i) or ("辅助核算%d" % i))
        if nm in cols:
            nm = nm + "_拆分"
        names.append(nm)
    strip = (p.get("strip") == "1")
    base = "trim(coalesce(%s, ''))" % qi(col)
    if strip:
        base = "trim(regexp_replace(%s, '^[^:：]{1,10}[:：]\\s*', ''))" % base
    exprs = [qi(c) for c in cols]
    for i, nm in enumerate(names, 1):
        if sep:
            e = "nullif(trim(list_extract(string_split(%s, %s), %d)), '')" % (base, ql(sep), i)
        else:
            e = "nullif(%s, '')" % base if i == 1 else "NULL"
        exprs.append("%s AS %s" % (e, qi(nm)))
    return {"kind": "select", "sql": "SELECT %s FROM {t}" % ", ".join(exprs), "adds": names}


def _b_drop_dup(p, cols, t):
    use = p.get("cols") or [c for c in cols if c != "_src_row"]
    use = [c for c in use if c in cols] or cols
    part = ", ".join(qi(c) for c in use)
    inner = "SELECT *, row_number() OVER (PARTITION BY %s ORDER BY _src_row) AS _rn FROM {t}" % part
    return {
        "kind": "select",
        "sql": "SELECT * EXCLUDE (_rn) FROM (%s) WHERE _rn = 1 ORDER BY _src_row" % inner.replace("{t}", "{t}"),
        "drop_sql": "SELECT _src_row FROM (%s) WHERE _rn > 1" % inner.replace("{t}", "{t}"),
        "reason": "整行重复",
    }


def _b_add_voucher_key(p, cols, t):
    """把「公司名称 + 会计期间 + 凭证号」拼成一列唯一凭证号。

    为什么需要
        多公司/多账套合并后，各家的凭证号都从「记-001」重新编号，
        直接按凭证号分组核对借贷平衡，会把两家公司的 记-001 当成同一张凭证。
        拼上公司名称和月份，才是真正意义上的「一张凭证」。

    取哪几列
        公司名称列 / 会计期间列 / 凭证号列，三列都选最准。
        少选也可以（比如单公司账套只选期间+凭证号），拼出来的段数相应变少；
        某一行的某段为空时自动跳过该段，不会拼出「A--001」这种空档。
    """
    ent = p.get("entity") or ""
    per = p.get("period") or ""
    vno = p.get("vno") or ""
    used = [c for c in (ent, per, vno) if c in cols]
    if not used:
        return {"kind": "noop", "reason": "三列都没选或都不在表里"}
    sep = p.get("sep")
    sep = "-" if sep is None or str(sep) == "" else str(sep)
    new = str(p.get("name") or "").strip() or "唯一凭证号"
    if new in cols:
        new = new + "_生成"
    # concat_ws 会跳过 NULL，所以空段不会留下连续分隔符
    args = ", ".join("nullif(trim(cast(%s AS VARCHAR)), '')" % qi(c) for c in used)
    e = "concat_ws(%s, %s)" % (ql(sep), args)
    exprs = [qi(c) for c in cols] + ["%s AS %s" % (e, qi(new))]
    miss = []
    if ent not in cols:
        miss.append("公司名称")
    if per not in cols:
        miss.append("会计期间")
    if vno not in cols:
        miss.append("凭证号")
    note = ""
    if miss:
        note = "未选%s列，唯一凭证号可能重复" % "、".join(miss)
    return {"kind": "select", "sql": "SELECT %s FROM {t}" % ", ".join(exprs),
            "adds": [new], "note": note}


def _b_check_balance(p, cols, t):
    return {"kind": "check", "params": p}


STEPS = [
    {
        "type": "drop_empty", "name": "删除空行", "group": "选行", "build": _b_drop_empty,
        "desc": "按你勾选的 N 列判空：勾中的列全部为空才算空行（也可切成「任一列为空就删」）。",
        "params": [
            {"key": "cols", "label": "判断列", "kind": "cols", "default": [],
             "hint": "可多选；不勾 = 所有列都算。勾 2 列以上就是「这几列都空才删」"},
            {"key": "mode", "label": "判空方式", "kind": "select",
             "options": ["all:所选列全部为空才删", "any:所选列任一为空就删"],
             "default": "all"},
        ],
    },
    {
        "type": "drop_empty_key", "name": "按关键列判空删行", "group": "选行", "build": _b_drop_empty_key,
        "desc": "账簿/公司/科目编码/科目名称/期间/凭证号 这几列全空才删。专门抓页脚行："
                "金额落在某列、但关键列全空，普通「删除空行」删不掉。",
        "params": [],
    },
    {
        "type": "drop_subtotal", "name": "删除小计行", "group": "选行", "build": _b_drop_subtotal,
        "desc": "「合计/小计/总计」出现在格子开头或结尾才删（本月合计、本页合计、合计：2,405,000 都算），"
                "中间夹字的不删——避免误删「其他合计性支出」这类真实科目。",
        "params": [
            {"key": "cols", "label": "检查列（可多选，留空=全部列）", "kind": "cols", "default": []},
        ],
    },
    {
        "type": "drop_keyword", "name": "按关键词删行", "group": "选行", "build": _b_drop_keyword,
        "desc": "命中关键词的行删掉：合计行 / 小计行 / 重复表头 / 页脚（制表人、打印日期）。",
        "params": [
            {"key": "col", "label": "检查列", "kind": "col", "default": "",
             "hint": "留空 = 任一列命中即算；建议选「科目名称」或「科目编码」"},
            {"key": "kw", "label": "关键词", "kind": "text", "default": "合计,小计,总计,本页合计",
             "hint": "逗号分隔"},
            {"key": "mode", "label": "匹配方式", "kind": "select",
             "options": ["包含", "全等", "正则"], "default": "包含"},
            {"key": "reason", "label": "剔除原因（留痕用）", "kind": "text", "default": "合计/小计行"},
        ],
    },
    {
        "type": "drop_base_cur", "name": "删除综合本位币行", "group": "选行",
        "build": _b_drop_base_cur,
        "desc": "删掉币别列写着「综合本位币/本位币/记账本位币」的汇总行。多币别账套里"
                "这一行与「人民币」「美元」等具体币别行并存，两份都算会翻倍；"
                "而且它看不出具体币别，留着对原币核对没用。整表都是汇总行时不删。",
        "params": [
            {"key": "col", "label": "币别列", "kind": "col", "default": "",
             "hint": "留空自动找「币别/币种/currency」列"},
        ],
    },
    {
        "type": "keep_base_cur", "name": "只保留综合本位币行", "group": "选行",
        "build": _b_keep_base_cur,
        "desc": "与「删除综合本位币行」互斥的另一半：只留下「综合本位币」汇总行，删掉所有"
                "分币别明细行。用于客户只导了汇总视图（按明细取不到账户），或项目组决定"
                "按汇总层上数（此时底稿无原币明细列）。整表没有汇总行时不执行。",
        "params": [
            {"key": "col", "label": "币别列", "kind": "col", "default": "",
             "hint": "留空自动找「币别/币种/currency」列"},
        ],
    },
    {
        "type": "drop_dup", "name": "整行去重", "group": "选行", "build": _b_drop_dup,
        "desc": "完全相同的行只留第一条，其余进剔除记录。",
        "params": [
            {"key": "cols", "label": "参与比较的列", "kind": "cols", "default": [],
             "hint": "留空 = 所有列都参与"},
        ],
    },
    {
        "type": "trim_text", "name": "文本归一", "group": "改值", "build": _b_trim_text,
        "desc": "去掉首尾空白、全角空格缩进，连续空白压成一个空格。",
        "params": [
            {"key": "cols", "label": "处理列", "kind": "cols", "default": [],
             "hint": "留空 = 所有列；一般只选名称、摘要类列"},
        ],
    },
    {
        "type": "pad_code", "name": "编码补零", "group": "改值", "build": _b_pad_code,
        "desc": "科目编码被 Excel 吃掉前导零时补回来（1001 → 001001）。",
        "params": [
            {"key": "col", "label": "编码列", "kind": "col", "default": ""},
            {"key": "length", "label": "补齐到几位", "kind": "number", "default": 0,
             "hint": "填 0 = 自动取该列最常见的长度"},
        ],
    },
    {
        "type": "negate_col", "name": "整列取反", "group": "改值", "build": _b_negate_col,
        "desc": "把某一列的正负号翻过来：银行流水「支出记负」和账簿方向相反、红字冲销要转成正数时用。"
                "只翻数字，文本行（合计、备注）原样保留；就地改，不新增列。",
        "params": [
            {"key": "col", "label": "目标列", "kind": "col", "default": "",
             "hint": "一般选借方/贷方/发生额这类金额列"},
            {"key": "mode", "label": "改法", "kind": "select",
             "options": ["flip:整列取反（正变负、负变正）",
                         "to_negative:统一成负数（正数取反，负数不动）",
                         "to_positive:统一成正数（负数取反，正数不动）"],
             "default": "flip"},
        ],
    },
    {
        "type": "fill_down", "name": "向下填充", "group": "填值", "build": _b_fill_down,
        "desc": "辅助核算明细成行时，科目编码/名称只在第一行有值，向下填满同一个科目的所有明细行。",
        "params": [
            {"key": "col", "label": "要填充的列", "kind": "col", "default": ""},
            {"key": "partition", "label": "分区列（可选）", "kind": "col", "default": "",
             "hint": "多主体/多账簿混在一张表时按此列分区，各自向下填，避免上一主体的末值串到下一主体"},
        ],
    },
    {
        "type": "fill_const", "name": "填充固定值", "group": "派生", "build": _b_fill_const,
        "desc": "新增一列并整列填同一个值。公司名称、币种这类「整张表共用一个值」的字段用它补成标准列。",
        "params": [
            {"key": "name", "label": "新列名", "kind": "text", "default": "公司名称",
             "hint": "写成标准字段名最好认：公司名称 / 币种"},
            {"key": "value", "label": "取值", "kind": "text", "default": "",
             "hint": "例如 某某有限公司、人民币"},
        ],
    },
    {
        "type": "derive_level", "name": "科目级次推断", "group": "派生", "build": _b_derive_level,
        "desc": "按科目编码长度推断级次：1001 为 1 级、100101 为 2 级……新增一列，不改原数据。",
        "params": [
            {"key": "col", "label": "科目编码列", "kind": "col", "default": ""},
            {"key": "name", "label": "新列名", "kind": "text", "default": "科目级次"},
        ],
    },
    {
        "type": "add_l1", "name": "提取一级科目", "group": "派生", "build": _b_add_l1,
        "desc": "从某一列里按分隔符取第一段，作为一级科目。适用于「1001-库存现金」「管理费用/办公费」"
                "这种编码和名称挤在一列的余额表——拆出来的这一段就是核对时要按它归集的一级科目。",
        "params": [
            {"key": "col", "label": "源列", "kind": "col", "default": "",
             "hint": "一般是科目编码列，或「编码-名称」挤在一起的那一列"},
            {"key": "sep", "label": "分隔符", "kind": "text", "default": "-_/ .:|、",
             "hint": "留空 = 默认整套（- _ / 空格 . : | 、括号 ·）；科目名称常用 _ 分节，"
                     "从名称列拆一级科目时填 _ 即可"},
            {"key": "name", "label": "新列名", "kind": "text", "default": "一级科目",
             "hint": "建成后到「字段映射」页把它映射到标准字段「一级科目」"},
        ],
    },
    {
        "type": "derive_l1", "name": "按编码派生一级科目", "group": "派生", "build": _b_derive_l1,
        "desc": "定长科目编码（如 10020101）按前 N 位切出一级编码，再用内置标准科目表补一级科目名称，"
                "一次生成「一级代码」「一级科目」两列；只派生、不删任何行。名称查不到时先用本表里"
                "同前缀那一行的名称，再查不到就只写编码，兜底行数写进这一步的说明里。",
        "params": [
            {"key": "col", "label": "科目编码列", "kind": "col", "default": "",
             "hint": "定长数字编码那一列（10020101）；带分隔符（1002.01.01）、编码与名称挤一起"
                     "（1002 银行存款）也能切，只取前导数字"},
            {"key": "name_col", "label": "科目名称列（可选）", "kind": "col", "default": "",
             "hint": "标准科目表查不到该编码时，用本表里同前缀那一行的名称兜底；不选则查不到就只写编码"},
            {"key": "length", "label": "一级编码位数", "kind": "number", "default": 0,
             "hint": "填 0 = 自动推断（按标准科目表的键投票，8 位明细码也能推成 4 位）；推不出会提示手填"},
            {"key": "out_code", "label": "一级代码列名", "kind": "text", "default": "一级代码"},
            {"key": "out_name", "label": "一级科目列名", "kind": "text", "default": "一级科目",
             "hint": "生成后到「字段映射」页把它映射到标准字段「一级科目」"},
        ],
    },
    {
        "type": "add_acct_path", "name": "生成一级科目", "group": "派生", "build": _b_add_acct_path,
        "desc": "把多级科目编码展开：分出一级编码、向下填一级科目名称、按级拼接出科目全称，"
                "最后按「主体」删掉非末级行（只留末级叶子）。多账簿/多公司混在一张表时，"
                "每个主体独立判断，互不误删。不选主体列 = 整张表当成一个主体，照常清洗。",
        "params": [
            {"key": "code", "label": "科目编码列", "kind": "col", "default": "",
             "hint": "含多级点分编码的那一列，如 1002.01.01"},
            {"key": "name", "label": "科目名称列", "kind": "col", "default": "",
             "hint": "用来填一级科目名、拼科目全称"},
            {"key": "entity", "label": "主体列", "kind": "col", "default": "",
             "hint": "账簿名称/公司名称那一列；按它分组，每个主体独立判断末级。留空 = 整表一个主体"},
            {"key": "amount", "label": "金额列（可选）", "kind": "col", "default": "",
             "hint": "同一编码出现多行时，保留金额绝对值最大那一行（真实金额行）。不选则同码多行全部保留，不自动删。"},
            {"key": "parent_tol", "label": "父行残值容差（元）", "kind": "text", "default": "0.01",
             "hint": "父行金额与子行合计相差超过这个数，就认定父行还带着自己的金额，不删它。"
                     "默认 0.01 元；数据本身有尾差时可放宽"},
            {"key": "aux", "label": "核算维度列（可选）", "kind": "cols", "default": [],
             "hint": "用来判断「这个科目有没有明细」的列，多账簿余额表选「核算维度名称」。"
                     "任一列非空即视为明细行；不选则不启用明细优先判断"},
            {"key": "detail_first", "label": "有明细时只保留明细行", "kind": "select",
             "options": ["0:关（保持原行为）", "1:开"], "default": "0",
             "hint": "开启后：同一科目有核算维度明细时，删掉科目汇总行、保留全部明细；"
                     "没有明细的科目仍保留一行。多账簿 TB 必须开，否则会丢科目或留重复汇总"},
            {"key": "group_cols", "label": "分组列（可选）", "kind": "cols", "default": [],
             "hint": "按这些列分组独立判断，留空则自动带上币别列。多主体多币种建议把币别选上，"
                     "否则「美元有明细」会把「欧元只有汇总」的行删掉"},
            {"key": "sep", "label": "编码分隔符", "kind": "text", "default": ".",
             "hint": "科目编码层级间的分隔符，默认点号"},
            {"key": "join", "label": "全称拼接符", "kind": "text", "default": "_",
             "hint": "科目全称里各级名称之间的连接符，默认下划线，可换成 - / 等"},
        ],
    },
    {
        "type": "split_aux", "name": "辅助核算拆分", "group": "派生", "build": _b_split_aux,
        "desc": "把「客户:阿里巴巴/部门:销售部」这种一格多值的辅助核算拆成辅助核算1/2/3 三列。",
        "params": [
            {"key": "col", "label": "辅助核算列", "kind": "col", "default": ""},
            {"key": "sep", "label": "分隔符", "kind": "text", "default": "/",
             "hint": "留空 = 不拆，整列当辅助核算1"},
            {"key": "strip", "label": "去掉「客户:」这类前缀", "kind": "select",
             "options": ["0:保留前缀", "1:去掉前缀"], "default": "0"},
        ],
    },
    {
        "type": "add_voucher_key", "name": "生成唯一凭证号", "group": "派生",
        "build": _b_add_voucher_key,
        "desc": "公司名称 + 会计期间 + 凭证号 拼成一列「唯一凭证号」，用于按整张凭证核对借贷平衡。"
                "多公司账套里各家凭证号都从「记-001」重新编号，不加公司名称就会把两家的同一号凭证"
                "当成一张。生成后到「字段映射」把它映射到「唯一凭证号」。",
        "params": [
            {"key": "entity", "label": "公司名称列", "kind": "col", "default": "",
             "hint": "多公司/多账套必选；单公司账套可留空"},
            {"key": "period", "label": "会计期间列", "kind": "col", "default": "",
             "hint": "月份/所属期间那一列，如 2026-01"},
            {"key": "vno", "label": "凭证号列", "kind": "col", "default": "",
             "hint": "凭证字号/凭证编号那一列"},
            {"key": "sep", "label": "分隔符号", "kind": "text", "default": "-",
             "hint": "默认 -，拼出来是 公司-2026-01-记001"},
            {"key": "name", "label": "新列名", "kind": "text", "default": "唯一凭证号"},
        ],
    },
    {
        "type": "check_balance", "name": "勾稽校验", "group": "校验", "build": _b_check_balance,
        "desc": "期初 + 本期借 − 本期贷 = 期末。只出报告，不改数据。",
        "params": [
            {"key": "open", "label": "期初列", "kind": "col", "default": ""},
            {"key": "dr", "label": "本期借方列", "kind": "col", "default": ""},
            {"key": "cr", "label": "本期贷方列", "kind": "col", "default": ""},
            {"key": "end", "label": "期末列", "kind": "col", "default": ""},
            {"key": "tol", "label": "允许差额", "kind": "number", "default": 0.01},
        ],
    },
]

BY_TYPE = {s["type"]: s for s in STEPS}


def types_for_ui():
    """给前端用的定义（去掉 build 函数）。"""
    out = []
    for s in STEPS:
        out.append({k: v for k, v in s.items() if k != "build"})
    return out


# ---------------------------------------------------------------- 配方
def _find(cols, *words, **kw):
    excl = kw.get("exclude") or ()
    for c in cols:
        n = str(c)
        if any(w and w in n for w in words) and not any(e in n for e in excl):
            return c
    return ""


def _preset_flow():
    """银行流水的默认清洗：只做最保守的四件事。

    不套用科目表那套「合计/小计」规则——流水里「转账」「合计」可能出现在摘要里，
    按关键词删会把真实交易删掉。账号去空格不在这里做（`_acct_norm` 在合并时统一处理），
    因为流水里带空格的账号是银行原样格式，删了反而不好回溯。
    """
    steps = [
        {"type": "drop_empty", "params": {"cols": []}},
        {"type": "drop_keyword", "params": {
            "col": "", "kw": "制表人,制表单位,打印日期,打印时间,复核人,财务负责人",
            "mode": "包含", "reason": "页脚"}},
        {"type": "drop_keyword", "params": {
            "col": "",
            "kw": "交易日期,记账日期,交易金额,收入金额,支出金额,对方账号,对方户名,账户余额",
            "mode": "包含", "reason": "重复表头（多页导出常见）"}},
        {"type": "trim_text", "params": {"cols": []}},
    ]
    return [{"id": "p%d" % i, "type": s["type"], "enabled": True, "params": s["params"]}
            for i, s in enumerate(steps, 1)]


def _preset_fa(cols):
    """固定资产台账的默认清洗：去空行 → 删合计/小计行 → 去空格 → 整行去重。

    为什么不套用科目表那套：台账没有「科目名称」列，`preset` 里按科目列删合计会
    拿到空列名，一条也删不掉。台账的合计行写法是「分类」列填"合计"、金额照填，
    不删就会多出一个叫"合计"的分类、金额翻倍——实测过。
    """
    cat = _find(cols, "分类", "类别", "资产分类", "Asset")
    steps = [
        {"type": "drop_empty", "params": {"cols": []}},
        {"type": "drop_keyword", "params": {
            "col": cat, "kw": "合计,小计,总计,Total,Grand Total", "mode": "包含",
            "reason": "合计/小计行"}},
        {"type": "drop_keyword", "params": {
            "col": cat, "kw": "分类,类别", "mode": "全等", "reason": "重复表头"}},
        {"type": "trim_text", "params": {"cols": [c for c in (cat,) if c]}},
        {"type": "drop_dup", "params": {"cols": []}},
    ]
    return [{"id": "p%d" % i, "type": s["type"], "enabled": True, "params": s["params"]}
            for i, s in enumerate(steps, 1)]


def preset(name, cols, target="tb"):
    """按列名猜一套默认步骤。cols 是原始表列名。"""
    if target == "flow":
        return _preset_flow()
    if target in ("fa", "fa_beg"):
        return _preset_fa(cols)
    code = _find(cols, "科目编码", "科目代码", "科目号", "账户编码", "编码") or _find(cols, "代码")
    acct = _find(cols, "科目名称", "账户名称", "科目", exclude=("编码", "代码", "级次"))
    aux = _find(cols, "辅助核算", "核算项目", "辅助项")

    common = [
        {"type": "drop_empty", "params": {"cols": []}},
        # 综合本位币行必须在删合计行之前处理：它本身金额不小，被当成「合计行」
        # 命中与否取决于写法，先按币别明确删掉，后面就不会有歧义。
        {"type": "drop_base_cur", "params": {"col": ""}},
        {"type": "drop_keyword", "params": {
            "col": acct or code, "kw": "合计,小计,总计,本页合计,本期合计,本年累计",
            "mode": "包含", "reason": "合计/小计行"}},
        {"type": "drop_keyword", "params": {
            "col": code or acct, "kw": "科目编码,科目代码,科目名称,科目", "mode": "全等",
            "reason": "重复表头"}},
        {"type": "drop_keyword", "params": {
            "col": "", "kw": "制表人,制表单位,打印日期,打印时间,制表日期,复核人,财务负责人",
            "mode": "包含", "reason": "页脚"}},
        {"type": "trim_text", "params": {"cols": [c for c in (acct,) if c]}},
    ]
    if name == "empty":
        common = []
    elif name == "literal":          # 只做最基本的：空行 + 去空格
        common = [common[0], common[5]]
    elif name == "aux_rows":         # 辅助核算明细成行
        common += [
            {"type": "fill_down", "params": {"col": code}} if code else None,
            {"type": "fill_down", "params": {"col": acct}} if acct else None,
            {"type": "derive_level", "params": {"col": code, "name": "科目级次"}} if code else None,
        ]
    elif name == "aux_split":        # 一格多值的辅助核算
        common += [
            {"type": "split_aux", "params": {"col": aux, "sep": "/", "strip": "0"}} if aux else None,
        ]
    common = [c for c in common if c]
    if name != "empty" and name != "literal":
        common.append({"type": "drop_dup", "params": {"cols": []}})
    out = []
    for i, s in enumerate(common, 1):
        out.append({"id": "p%d" % i, "type": s["type"], "enabled": True, "params": s["params"]})
    return out


PRESETS = [
    {"key": "generic", "name": "通用余额表（推荐）",
     "desc": "去空行 → 删合计/小计 → 删重复表头 → 去空格 → 整行去重"},
    {"key": "aux_rows", "name": "辅助核算明细成行",
     "desc": "通用清洗 + 科目编码/名称向下填充 + 级次推断"},
    {"key": "aux_split", "name": "单列辅助核算拆分",
     "desc": "通用清洗 + 把「客户:A/部门:B」拆成辅助核算1/2/3"},
    {"key": "literal", "name": "最轻清洗（只去空行和空格）",
     "desc": "不确定表结构时用这个，动手最少"},
    {"key": "empty", "name": "空白配方（自己配）", "desc": "从零开始一个个加动作"},
]


# ---------------------------------------------------------------- 执行
def columns_of(conn, table):
    cur = conn.execute("SELECT * FROM %s LIMIT 0" % qi(table))
    return [d[0] for d in cur.description]


def _auto_len(conn, table, col, limit=5000):
    """该列最常见的长度（只看纯数字值）。

    长度并列时取**最短**的那个。余额表里 4 位一级码和 10 位明细码并存是常态，
    并列时若不指定次序，同一份数据两次跑可能得到不同结果（实测 4 和 10 都出现过），
    而取到 10 会把 2 位的短码补成 0000000022。
    """
    if col not in columns_of(conn, table):
        return 0
    sql = ("SELECT length(trim(%s)) AS L, count(*) c FROM %s "
           "WHERE regexp_full_match(trim(coalesce(%s, '')), '\\d+') "
           "GROUP BY 1 ORDER BY c DESC, L ASC LIMIT 1") % (qi(col), qi(table), qi(col))
    try:
        r = conn.execute(sql).fetchone()
        return int(r[0]) if r and r[0] else 0
    except Exception:
        return 0


def _auto_l1_len(conn, table, col, std):
    """推断「一级编码取几位」。返回 (位数, 失败原因)，位数为 0 表示推不出。

    不能沿用 _auto_len（该列最常见的长度）：53 家客户全是 8 位明细码时，最常见长度
    就是 8，切出来的「一级科目」等于科目编码本身——检查全绿、核对全错，最难发现。
    改用标准科目表的键投票：8 位码的前 4 位能命中 1002/1122，8 位自身命中不了，于是得 4。
    """
    if not col or col not in columns_of(conn, table):
        return 0, "所选列不在当前表里"
    try:
        rows = conn.execute(
            "SELECT c.L, substr(s.d, 1, c.L) AS k, count(*) AS n FROM "
            "(SELECT regexp_extract(trim(coalesce(%s, '')), '[0-9]+', 0) AS d FROM %s) s "
            "JOIN (VALUES (2),(3),(4),(5),(6),(7),(8),(9),(10),(11),(12)) c(L) "
            "  ON length(s.d) >= c.L "
            "WHERE s.d <> '' GROUP BY 1, 2" % (qi(col), qi(table))).fetchall()
    except Exception as e:
        return 0, "推断位数时读取编码列失败：%s" % str(e)[:80]

    per = {}
    for L, k, n in rows:
        a = per.setdefault(int(L), [0, 0])
        a[0] += int(n)
        if k in std:
            a[1] += int(n)
    if not per:
        return 0, ("所选列里没有以数字开头的科目编码，无法推断位数；"
                   "请确认选中的是科目编码列（如 10020101、1002.01）")

    # 数据里真实出现过的编码长度（不是前缀长度）——「数据有没有层级」只能看这个
    try:
        actual = {int(L): int(n) for L, n in conn.execute(
            "SELECT length(d) AS L, count(*) AS n FROM (SELECT regexp_extract("
            "trim(coalesce(%s, '')), '[0-9]+', 0) AS d FROM %s) WHERE d <> '' GROUP BY 1"
            % (qi(col), qi(table))).fetchall()}
    except Exception:
        actual = {}

    top = max(v[0] for v in per.values()) or 1
    maxlen = min(max(actual) if actual else 12, 12)
    for L in range(2, maxlen + 1):
        tot, hit = per.get(L, (0, 0))
        if tot < top * 0.02:            # 滤掉零星脏码（如个别 2 位异常值）
            continue
        if tot and hit / tot >= 0.5:     # 标准科目表能认出来的最短一档，就是一级
            return L, None
    if len(actual) > 1:
        return min(actual), None         # 标准表没命中，但数据里长短码并存 → 取最短那档
    only = maxlen
    return 0, ("编码长度统一为 %d 位、内置标准科目表也没命中它的前几位，判断不出一级科目几位；"
               "请把「一级编码位数」手填上（8 位明细码通常填 4）" % only)


def _resolve(conn, table, cols, step, data_dir=None):
    p = dict(step.get("params") or {})
    ty = step.get("type")
    if ty == "pad_code":
        col = p.get("col") or ""
        if not int(p.get("length") or 0) and col:
            p["_auto_len"] = _auto_len(conn, table, col)
    elif ty == "derive_l1":
        if data_dir:
            p["_data_dir"] = data_dir
        std = std_accounts.load(data_dir)      # 带 mtime 缓存，不每步读盘
        p["_std"] = std
        col = p.get("col") or ""
        try:
            n = int(p.get("length") or 0)
        except (TypeError, ValueError):
            n = 0
        if col and n <= 0:
            ln, why = _auto_l1_len(conn, table, col, std)
            p["_auto_l1_len"] = ln
            if not ln:
                p["_auto_l1_reason"] = why
    return p


def run(conn, sid, raw_table, steps, upto=None, keep=False, sample=30, data_dir=None):
    """执行流水线。keep=True 时把结果落成 stg_<sid>，否则只算不落。

    返回 {steps:[{i,type,name,before,after,dropped,note}], columns, rows, ...}
    """
    steps = [s for s in (steps or []) if s.get("enabled", True)]
    if upto:
        steps = steps[:int(upto)]
    cols = columns_of(conn, raw_table)
    drop_tbl = "_dropped_" + sid
    conn.execute("DROP TABLE IF EXISTS %s" % qi(drop_tbl))
    conn.execute('CREATE TABLE %s(_src_row BIGINT, step INTEGER, step_name VARCHAR, '
                 'reason VARCHAR)' % qi(drop_tbl))
    cur = raw_table
    owned = []
    stats = []
    checks = []
    total_before = conn.execute("SELECT count(*) FROM %s" % qi(cur)).fetchone()[0]
    error = None
    for i, st in enumerate(steps, 1):
        spec = BY_TYPE.get(st.get("type"))
        if not spec:
            continue
        name = spec["name"]
        try:
            # 每步都重新取列：上一步可能新增了列（填充固定值/级次推断/辅助核算拆分），
            # 后面的动作要能看到它们，否则投影类动作会把新列丢掉
            cur_cols = columns_of(conn, cur)
            p = _resolve(conn, cur, cur_cols, st, data_dir)
            p["_conn"] = conn          # 少数动作要回查数据（如父行递归判定），只在本进程内用
            plan = spec["build"](p, cur_cols, cur)
        except Exception as e:
            stats.append({"i": i, "type": st.get("type"), "name": name,
                          "before": None, "after": None, "dropped": 0,
                          "error": str(e)[:200]})
            error = "%s：%s" % (name, str(e)[:200])
            break
        if plan["kind"] == "noop":
            stats.append({"i": i, "type": st.get("type"), "name": name,
                          "before": None, "after": None, "dropped": 0,
                          "note": plan.get("reason") or "跳过"})
            continue
        if plan["kind"] == "check":
            try:
                ck = _run_check(conn, cur, cur_cols, p)
            except Exception as e:
                ck = {"error": str(e)[:200]}
            ck["i"] = i
            ck["name"] = name
            checks.append(ck)
            stats.append({"i": i, "type": st.get("type"), "name": name,
                          "before": None, "after": None, "dropped": 0,
                          "note": "校验", "check": ck})
            continue
        before = conn.execute("SELECT count(*) FROM %s" % qi(cur)).fetchone()[0]
        nxt = "_cl_%s_%d" % (sid, i)
        try:
            if plan["kind"] == "filter":
                conn.execute(
                    "INSERT INTO %s SELECT _src_row, %d, %s, %s FROM %s WHERE NOT (%s)"
                    % (qi(drop_tbl), i, ql(name), ql(plan.get("reason") or name), qi(cur), plan["cond"]))
                conn.execute("CREATE OR REPLACE TABLE %s AS SELECT * FROM %s WHERE %s"
                             % (qi(nxt), qi(cur), plan["cond"]))
            else:
                sql = plan["sql"].replace("{t}", qi(cur))
                conn.execute("CREATE OR REPLACE TABLE %s AS %s" % (qi(nxt), sql))
                if plan.get("drops"):
                    # 分类剔除：每类带自己的原因，留痕里能看出删的是汇总行还是真实明细
                    for _label, _dsql in plan["drops"]:
                        conn.execute(
                            "INSERT INTO %s SELECT _src_row, %d, %s, %s FROM (%s)"
                            % (qi(drop_tbl), i, ql(name), ql(_label),
                               _dsql.replace("{t}", qi(cur))))
                elif plan.get("drop_sql"):
                    conn.execute(
                        "INSERT INTO %s SELECT _src_row, %d, %s, %s FROM (%s)"
                        % (qi(drop_tbl), i, ql(name), ql(plan.get("reason") or name),
                           plan["drop_sql"].replace("{t}", qi(cur))))
        except Exception as e:
            stats.append({"i": i, "type": st.get("type"), "name": name,
                          "before": before, "after": None, "dropped": 0,
                          "error": str(e)[:200]})
            error = "%s：%s" % (name, str(e)[:200])
            break
        after = conn.execute("SELECT count(*) FROM %s" % qi(nxt)).fetchone()[0]
        _note = plan.get("note") or plan.get("reason") or ""
        if before > 0 and after == 0:
            # 这类结果几乎都是参数或前置顺序问题，不是「数据本来就是空的」，
            # 静默输出 0 行会让人以为是数据问题，必须说清楚
            _note = ("⚠ 本步把 %d 行清成了 0 行。通常是参数或前置步骤顺序不对"
                     "（例如先删了「所有账簿」汇总层，导致科目编码全空、无值可填）"
                     % before) + ((" " + _note) if _note else "")
        stats.append({"i": i, "type": st.get("type"), "name": name,
                      "before": before, "after": after, "dropped": before - after,
                      "note": _note})
        if cur != raw_table and cur in owned:
            conn.execute("DROP TABLE IF EXISTS %s" % qi(cur))
        cur = nxt
        owned.append(nxt)

    n_after = conn.execute("SELECT count(*) FROM %s" % qi(cur)).fetchone()[0]
    rows = []
    if sample:
        cur_r = conn.execute("SELECT * FROM %s LIMIT %d" % (qi(cur), int(sample)))
        all_cols = [d[0] for d in cur_r.description]
        rows = [["" if v is None else str(v) for v in r] for r in cur_r.fetchall()]
    else:
        all_cols = columns_of(conn, cur)
        rows = []
    keep = [i for i, c in enumerate(all_cols) if c != "_src_row"]
    out_cols = [all_cols[i] for i in keep]
    rows = [[r[i] for i in keep] for r in rows]

    if keep:
        stg = "stg_" + sid
        conn.execute("DROP TABLE IF EXISTS %s" % qi(stg))
        if cur == raw_table:
            conn.execute("CREATE TABLE %s AS SELECT * FROM %s" % (qi(stg), qi(raw_table)))
        else:
            conn.execute("ALTER TABLE %s RENAME TO %s" % (qi(cur), qi(stg)))
            owned = [o for o in owned if o != cur]
    for o in owned:
        conn.execute("DROP TABLE IF EXISTS %s" % qi(o))

    return {"columns": out_cols, "rows": rows, "total_before": total_before,
            "total_after": n_after, "steps": stats, "checks": checks,
            "dropped": total_before - n_after, "error": error,
            "applied": bool(keep)}


def _run_check(conn, table, cols, p):
    def e(k):
        c = p.get(k) or ""
        return "_num(%s)" % qi(c) if c in cols else "0"

    expr = "COALESCE(%s, 0) + COALESCE(%s, 0) - COALESCE(%s, 0) - COALESCE(%s, 0)" % (
        e("open"), e("dr"), e("cr"), e("end"))
    tol = float(p.get("tol") or 0.01)
    sql = ("SELECT count(*) AS n, count(*) FILTER (WHERE abs(_d) > %f) AS bad, "
           "sum(_d) AS diff FROM (SELECT %s AS _d FROM %s) t" % (tol, expr, qi(table)))
    r = conn.execute(sql).fetchone()
    return {"rows": int(r[0] or 0), "bad": int(r[1] or 0),
            "diff": float(r[2] or 0), "tol": tol}


def dropped_rows(conn, sid, raw_table, limit=200):
    t = "_dropped_" + sid
    try:
        cur = conn.execute(
            "SELECT d._src_row, d.step, d.step_name, d.reason FROM %s d "
            "ORDER BY d._src_row LIMIT %d" % (qi(t), int(limit)))
    except Exception:
        return {"columns": [], "rows": [], "total": 0}
    cols = [d[0] for d in cur.description]
    rows = [list(r) for r in cur.fetchall()]
    total = conn.execute("SELECT count(*) FROM %s" % qi(t)).fetchone()[0]
    return {"columns": cols, "rows": rows, "total": total}
