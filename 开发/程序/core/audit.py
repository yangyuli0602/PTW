# -*- coding: utf-8 -*-
"""清洗达标自检：把「洗得对不对」变成可判定的清单。

为什么要有这个：清洗是整条链路的分水岭——洗过头（删掉真实行）和洗不够
（留着父科目/合计行）都会让后面的五组检查、BKD 上数全盘失真，但两种错误
在页面上都只表现为「数字对不上」，看不出根因。

所以这里把标准写成可执行判定，逐项给「达标 / 不达标 / 不适用」，
并配一段给 AI 的提示词（PROMPT），让接手的 AI 知道自己要达到什么状态。
"""
import json

# ------------------------------------------------------------------ 提示词
# 页面上「复制 AI 提示词」复制的是这一份，不要在别处再写一遍。
PROMPT = """你是审计数据处理助手，正在使用 PTW 平台把客户的科目余额表和序时账
清洗、映射、合并成标准表。请严格按下面的标准把数据洗到「完整且最细颗粒度」，
不合格就继续修，不要带着问题往下走。

【目标状态 · 科目余额表】
1. 只保留末级科目。不允许父科目行与明细行同时存在（例如同时有 1002 和 100201）。
2. 不含任何合计 / 小计 / 总计行，也不含「核算维度=合计」这类汇总行。
   判定口径：「合计 / 小计 / 总计」出现在单元格的**开头或结尾**才算汇总行。
   名称中间夹着这两个字的（例如「其他合计性支出」「结转本年合计支出」）是真实科目，
   必须保留——按关键词删行时不要用「包含」模式把这类行一起删掉。
3. 每个科目的辅助核算要完整列出，不能只剩一级科目名：
   - 银行存款 / 其他货币资金：必须有账号（bank_account 或辅助核算里的账号）
   - 应收账款 / 应付账款 / 预收 / 预付 / 其他应收 / 其他应付：必须有客商
   - 有部门、项目核算的科目：对应的辅助核算列不能整列为空
4. 一级科目已派生（acct_l1 不能为空），科目编码与名称不能有空值。
5. 币别行要全：多币别账套里「综合本位币」汇总行与外币明细行并存时，
   保留明细行并标出币别，不要把汇总行当成一个科目。

【目标状态 · 序时账】
1. 日期、凭证号、摘要、科目编码、借贷金额这五项不能有空值。
2. 借贷金额不能出现「两边都为空」的行——那说明金额列没映射对。
3. 一级科目已派生，科目全称（acct_full）尽量有，用于第四组按全称核对。
4. 辅助核算（客商 / 部门 / 项目）不要整列为空，否则按账户、按往来核对做不了。

【操作步骤】
1. 导入后给每张工作表指定正确目标表（科目余额表 / 序时账 / 银行流水）。
2. 在「清洗与映射」页逐张确认：表头行要选对，两层表头填 2。
3. 清洗动作按需启用：删空行 → 删合计小计 → 删父科目（或只导末级）
   → 补一级科目 → 辅助核算拆分 → 整行去重。每次都看「剔除留痕」，
   确认删掉的是该删的，不是把真实行删了。
4. 字段映射必须把带 * 的必填字段映射上；金额列要注意「原币 / 本位币」不要配反。
5. 点平台上的「检查清洗达标」，逐项看结论；不达标的按提示语修，再检查。
6. 达标后合并成标准表，再跑「数据检查」五组，确认全部通过。
7. 最后按下面的格式输出结论，不要省略任何一项。

【输出格式 · 必须照这个写】
一、清洗达标情况
  科目余额表：末级科目 [达标/不达标] / 无合计小计 [达标/不达标] / 辅助核算 [达标/不达标] / 必备字段 [达标/不达标]
  序时账：    必备字段 [达标/不达标] / 金额完整 [达标/不达标] / 一级科目 [达标/不达标]
二、五组检查结果（每项写：组名、总行数、差异数、结论）
  第一组 余额表自平衡：…
  第二组 序时账自平衡：…
  第三组 两表一级科目：…
  第四组 两表科目全称：…
  第五组 货币资金辅助核算可用性：…
三、整体结论：通过 / 不通过。不通过时逐条写「哪个表、哪一列、差多少、下一步改什么」。

【硬约束】
- 不要为了凑平衡去改数据、删行、加调整分录。差额必须解释，解释不了就如实报告。
- 不确定的行宁可留着并标出来，也不要静默删掉。
- 数字对不上时先怀疑自己的清洗和映射（父科目重复、汇总层重复、期间错位、
  原币与本币配反），确认不是这些之后再考虑是客户数据本身的问题。
"""


def _t(v):
    return "" if v is None else str(v).strip()


def _has(conn, table):
    try:
        n = conn.execute('SELECT count(*) FROM "%s"' % table).fetchone()[0]
        return n > 0
    except Exception:
        return False


def _scalar(conn, sql, default=0):
    try:
        v = conn.execute(sql).fetchone()
        return v[0] if v else default
    except Exception:
        return default


def _item(group, key, title, ok, value, detail, fix="", level="fail", block=True):
    """block=True：不达标会阻断合并（属「会导致假不符/假数据」的硬问题）。
    block=False：只提示不阻断（客户数据本身缺失，如辅助核算没填——平台造不出来，
    但不影响两表对账，拦下来只会让人无法推进）。"""
    return {"group": group, "key": key, "title": title, "ok": bool(ok),
            "value": value, "detail": detail, "fix": fix, "level": level,
            "block": bool(block)}


# ---------------------------------------------------------------- 科目余额表
def _audit_tb(conn, items):
    G = "科目余额表"
    if not _has(conn, "tb"):
        items.append(_item(G, "tb_missing", "已合并成标准表", False, "0 行",
                           "还没有合并科目余额表", "导入后在「合并与检查」点「生成标准表」"))
        return

    total = _scalar(conn, "SELECT count(*) FROM tb")
    items.append(_item(G, "tb_rows", "标准表已生成", True, "%s 行" % format(total, ","),
                       "科目余额表已合并成 %s 行" % format(total, ",")))

    # 1) 末级科目：**同一公司内**存在某编码是另一编码的前缀 → 父科目与明细并存
    #    必须按公司分组：多主体合并后，A 公司用一级科目 1002、B 公司用明细码 100201，
    #    全表 DISTINCT 会把它们当成「父子并存」误报（2026-11-12 瑞云实测 8 个里 7 个是误报）。
    parents = _scalar(conn, """
        WITH c AS (SELECT DISTINCT coalesce(trim(company_name), '') AS co,
                          trim(acct_code) AS cd FROM tb
                   WHERE nullif(trim(acct_code), '') IS NOT NULL)
        SELECT count(*) FROM c a
        WHERE EXISTS (SELECT 1 FROM c b
                      WHERE b.co = a.co AND b.cd <> a.cd AND starts_with(b.cd, a.cd))
    """)
    sample = []
    if parents:
        try:
            sample = [r[0] for r in conn.execute("""
                WITH c AS (SELECT DISTINCT coalesce(trim(company_name), '') AS co,
                                  trim(acct_code) AS cd FROM tb
                           WHERE nullif(trim(acct_code), '') IS NOT NULL)
                SELECT a.cd FROM c a
                WHERE EXISTS (SELECT 1 FROM c b
                              WHERE b.co = a.co AND b.cd <> a.cd AND starts_with(b.cd, a.cd))
                ORDER BY a.cd LIMIT 5
            """).fetchall()]
        except Exception:
            pass
    items.append(_item(
        G, "tb_leaf", "仅含末级科目", parents == 0,
        "父科目 %d 个" % parents if parents else "无父科目",
        ("父科目与明细并存：%s 等" % "、".join(sample)) if parents else "所有科目都是末级科目",
        "启用「生成一级科目」动作并设「有明细时只保留明细行=1」，"
        "或只导出末级科目余额表" if parents else ""))

    # 2b) 综合本位币汇总行与分币别明细并存
    #     两者净额往往完全相等（实测 4.7 亿 vs 4.7 亿）——同时取会让金额翻倍，
    #     只取汇总层又拿不到账户明细。必须由清洗配方显式二选一，不能平台替它猜。
    base_mix = _scalar(conn, """
        SELECT count(*) FROM (
          SELECT trim(acct_code) AS cd FROM tb
          WHERE nullif(trim(acct_code), '') IS NOT NULL
          GROUP BY 1
          HAVING sum(CASE WHEN coalesce(currency,'') LIKE '%综合本位币%' THEN 1 ELSE 0 END) > 0
             AND sum(CASE WHEN coalesce(currency,'') NOT LIKE '%综合本位币%' THEN 1 ELSE 0 END) > 0
        )
    """)
    items.append(_item(
        G, "tb_base_cur_mix", "综合本位币行不与分币别明细并存", base_mix == 0,
        "混用科目 %d 个" % base_mix if base_mix else "无混用",
        ("有 %d 个科目同时存在「综合本位币」汇总行与分币别明细行——两者金额通常相等，"
         "同时上数会翻倍，只取汇总层又丢账户明细" % base_mix) if base_mix
        else "综合本位币层与明细层没有同时存在",
        "在「清洗与映射」二选一：「删除综合本位币行」（保留原币明细，底稿有原币列）；"
        "或「只保留综合本位币行」（按汇总层上数，底稿无原币明细）" if base_mix else "",
        level="warn"))
    # level=warn：并存本身不会让两表对账出错（recon 按币别分组、底稿取数会排除汇总层），
    # 但它决定了底稿有没有原币明细——所以提示，不阻断。

    # 2) 合计 / 小计行
    #    口径必须和 clean._b_drop_subtotal 完全一致：关键词在格子开头或结尾才算。
    #    两边都写死了各自的匹配，一旦不同就会出现「体检说还有合计行 → 用户去点删除
    #    → 把「其他合计性支出」这种真科目删了」的连锁误伤。
    _hit = r"(合计|小计|总计)\s*$|^\s*(合计|小计|总计)"
    sub = _scalar(conn, """
        SELECT count(*) FROM tb
        WHERE regexp_matches(trim(coalesce(acct_name, '')), '%s')
           OR coalesce(aux1, '') IN ('合计','小计','总计','合计数')
           OR coalesce(aux2, '') IN ('合计','小计','总计','合计数')
           OR coalesce(aux3, '') IN ('合计','小计','总计','合计数')
           OR coalesce(aux_text, '') IN ('合计','小计','总计','合计数')
    """ % _hit)
    items.append(_item(G, "tb_subtotal", "无合计 / 小计行", sub == 0,
                       "%d 行" % sub,
                       "命中 %d 行合计/小计（含核算维度为「合计」的汇总行）；"
                       "这类行会让金额翻倍" % sub if sub else "没有合计/小计行残留",
                       "用「删除小计行」动作：它只认出现在格子开头或结尾的合计/小计，"
                       "不会误删「其他合计性支出」这类真实科目" if sub else ""))

    # 3) 必备字段
    #    口径必须和「字段映射」页的必备字段一致：公司名称、账簿名称也是必备。
    #    这两列是后面 BKD 分主体、货币资金分户的前提，空着照样能合并，
    #    但下游会把几个主体加在一起比，错得还很安静，所以在这里就卡住。
    no_code = _scalar(conn, "SELECT count(*) FROM tb WHERE nullif(trim(acct_code),'') IS NULL")
    no_name = _scalar(conn, "SELECT count(*) FROM tb WHERE nullif(trim(acct_name),'') IS NULL")
    no_l1 = _scalar(conn, "SELECT count(*) FROM tb WHERE nullif(trim(acct_l1),'') IS NULL")
    no_company = _scalar(conn, "SELECT count(*) FROM tb "
                              "WHERE nullif(trim(coalesce(company_name,'')),'') IS NULL")
    no_book = _scalar(conn, "SELECT count(*) FROM tb "
                           "WHERE nullif(trim(coalesce(book_name,'')),'') IS NULL")
    miss = []
    if no_code:
        miss.append("科目编码空 %d 行" % no_code)
    if no_name:
        miss.append("科目名称空 %d 行" % no_name)
    if no_l1:
        miss.append("一级科目空 %d 行" % no_l1)
    if no_book:
        miss.append("账簿名称空 %d 行" % no_book)
    if no_company:
        miss.append("公司名称空 %d 行" % no_company)
    items.append(_item(G, "tb_required", "必备字段无缺失", not miss,
                       "、".join(miss) if miss else "编码/名称/一级科目/账簿/公司齐全",
                       "；".join(miss) if miss else "科目编码、科目名称、一级科目、账簿名称、"
                                                    "公司名称都没有空值",
                       "一级科目空：定长编码（如 10020101）用「按编码派生一级科目」；"
                       "编码带分隔符或名称列可拆时用「提取一级科目 / 生成一级科目」；"
                       "账簿名称空：在映射页把"
                       "「账簿名称来源」设成某一列 / 文件名 / 手工填写；公司名称空："
                       "导入账簿对照表，或把「公司名称来源」设成等于账簿名称"))

    # 3b) 一级科目里混着纯编码：多半是位数填成了整串，或标准科目表没命中
    code_only = _scalar(conn, "SELECT count(*) FROM tb WHERE regexp_full_match("
                              "trim(coalesce(acct_l1,'')), '\\d+')")
    l1_eq_code = _scalar(conn, "SELECT count(*) FROM tb WHERE "
                               "length(trim(coalesce(acct_l1,''))) = length(trim(coalesce(acct_code,''))) "
                               "AND length(trim(coalesce(acct_code,''))) >= 6")
    bad_l1 = max(code_only, l1_eq_code)
    items.append(_item(
        G, "tb_l1_code_only", "一级科目不是纯编码", total == 0 or bad_l1 * 2 <= total,
        "纯编码 %d 行 / 与科目编码等长 %d 行" % (code_only, l1_eq_code),
        ("一级科目里 %d 行是纯编码（其中 %d 行与科目编码完全等长）——多半是「一级编码位数」"
         "填成了整串长度，或标准科目表没命中这些编码" % (code_only, l1_eq_code))
        if bad_l1 else "一级科目都带文字名称",
        "把「一级编码位数」改对（8 位明细码通常取 4 位）；缺的编码补进 data/std_accounts.json"
        if bad_l1 else "",
        level="warn"))

    # 4) 辅助核算完整度（按科目类型分别看，不适用就不报）
    aux_rules = [
        ("货币资金", "acct_l1 IN ('银行存款','其他货币资金')",
         # 只认「真账号」：专用列有值，或辅助核算里出现 8 位以上连续数字。
         # 不能只判「含账号二字」——「账号010」这种占位值会让检查静默通过。
         "(coalesce(bank_account,'') <> '' "
         "OR regexp_matches(coalesce(aux_text,''), '\\d{8,}'))",
         "银行存款/其他货币资金缺少账号", "把「核算维度编码」映射到辅助核算1，"
         "或把账号列映射到 bank_account"),
        ("往来", "acct_l1 IN ('应收账款','应付账款','预收账款','预付账款',"
                "'其他应收款','其他应付款')",
         "(coalesce(counterparty_name,'') <> '' OR coalesce(aux_text,'') <> '')",
         "往来科目缺少客商", "把客商/往来单位列映射到辅助核算，并在映射页标注角色为「客商」"),
    ]
    for label, where, filled, bad_txt, fix in aux_rules:
        tot = _scalar(conn, "SELECT count(*) FROM tb WHERE %s" % where)
        if not tot:
            continue
        got = _scalar(conn, "SELECT count(*) FROM tb WHERE %s AND %s" % (where, filled))
        rate = got / tot if tot else 1.0
        ok = rate >= 0.95
        items.append(_item(
            G, "tb_aux_" + ("cash" if label == "货币资金" else "party"),
            "%s辅助核算完整" % label, ok, "%.0f%%（%d/%d）" % (rate * 100, got, tot),
            ("%s：仅 %d/%d 行有辅助核算" % (bad_txt, got, tot)) if not ok
            else "%s %d 行中 %d 行带辅助核算" % (label, tot, got),
            fix if not ok else "", block=False))
        # block=False：客户没填辅助核算，平台造不出来；它影响的是底稿账户级明细，
        # 不影响两表对账——拦下来只会让人无法推进，所以在页面上如实提示、但放行合并。


# ---------------------------------------------------------------- 序时账
def _audit_jr(conn, items):
    G = "序时账"
    if not _has(conn, "journal"):
        items.append(_item(G, "jr_missing", "已合并成标准表", False, "0 行",
                           "还没有合并序时账", "导入后在「合并与检查」点「生成标准表」"))
        return

    total = _scalar(conn, "SELECT count(*) FROM journal")
    items.append(_item(G, "jr_rows", "标准表已生成", True, "%s 行" % format(total, ","),
                       "序时账已合并成 %s 行" % format(total, ",")))

    cols = set()
    try:
        cols = {r[0] for r in conn.execute("DESCRIBE journal").fetchall()}
    except Exception:
        pass

    def empty(col):
        return _scalar(conn, "SELECT count(*) FROM journal "
                             "WHERE nullif(trim(coalesce(%s::VARCHAR,'')), '') IS NULL" % col) \
            if col in cols else total

    # 凭证号允许用「唯一凭证号」代替
    no_vno = empty("vno")
    if "voucher_key" in cols:
        no_vno = _scalar(conn, "SELECT count(*) FROM journal "
                               "WHERE nullif(trim(coalesce(vno,'')), '') IS NULL "
                               "AND nullif(trim(coalesce(voucher_key,'')), '') IS NULL")
    miss = []
    for col, label in (("vdate", "日期"), ("summary", "摘要"), ("acct_code", "科目编码"),
                       ("book_name", "账簿名称")):
        n = empty(col)
        if n:
            miss.append("%s空 %d 行" % (label, n))
    if no_vno:
        miss.append("凭证号空 %d 行" % no_vno)
    items.append(_item(G, "jr_required", "必备字段无缺失", not miss,
                       "、".join(miss) if miss else "日期/凭证号/摘要/科目编码/账簿名称齐全",
                       "；".join(miss) if miss else "五项必备字段都没有空值",
                       "去「字段映射」把缺的字段配上；账簿名称没这一列时，改用"
                       "「账簿名称来源」= 文件名 / 工作表名 / 手工填写"))

    # 金额：两边都空说明金额列没映射对
    if "debit" in cols or "credit" in cols:
        no_amt = _scalar(conn, "SELECT count(*) FROM journal "
                               "WHERE coalesce(debit,0) = 0 AND coalesce(credit,0) = 0")
    else:
        no_amt = total
    rate = (total - no_amt) / total if total else 1.0
    items.append(_item(
        G, "jr_amount", "借贷金额完整", no_amt == 0, "%d 行无金额" % no_amt,
        ("%d 行借贷两边都为空（%.1f%%），金额列可能没映射对，或存在只有数量的凭证"
         % (no_amt, (1 - rate) * 100)) if no_amt else "每行都有借方或贷方金额",
        "检查「借方金额 / 贷方金额」是否映射到了正确列，注意别和原币列配反",
        level="fail"))

    # 一级科目
    no_l1 = empty("acct_l1")
    items.append(_item(
        G, "jr_l1", "一级科目已派生", no_l1 == 0, "%d 行缺一级科目" % no_l1,
        ("%d 行一级科目为空，第三、四组检查会跑不了" % no_l1) if no_l1
        else "一级科目已全部派生",
        "加「提取一级科目」（分隔符 _ - / . : | 、）或「按编码派生一级科目」"
        "（定长编码按前 4 位，位数可自动推断）" if no_l1 else ""))

    # 整百万行：导出工具分卷或截断的典型信号
    if total and total % 1000000 == 0:
        items.append(_item(
            G, "jr_million", "行数不是整百万", False, "%s 行" % format(total, ","),
            "恰好 %s 行，很可能是导出时被截断或只导了第一卷" % format(total, ","),
            "确认客户是不是分卷导出；有分卷要把其余卷一起导入", level="warn"))


# ---------------------------------------------------------------- 五组检查
def _audit_check(conn, items):
    G = "数据检查"
    row = None
    try:
        row = conn.execute("SELECT ran_at, params, summary FROM _recon_last LIMIT 1").fetchone()
    except Exception:
        pass
    if not row:
        items.append(_item(G, "check_missing", "已跑过五组检查", False, "未运行",
                           "还没有跑过数据检查", "清洗达标后到「合并与检查 → 数据检查」跑一次"))
        return
    try:
        summary = json.loads(row[2] or "{}")
        params = json.loads(row[1] or "{}")
    except Exception:
        items.append(_item(G, "check_bad", "已跑过五组检查", False, "解析失败",
                           "上一次检查结果读不出来", "重新跑一次数据检查"))
        return
    items.append(_item(G, "check_ran", "已跑过五组检查", True,
                       "运行于 %s" % str(row[0])[:19],
                       "期间 %s ~ %s，容差 %s"
                       % ((params.get("period") or {}).get("from") or "-",
                          (params.get("period") or {}).get("to") or "-",
                          params.get("tol"))))
    for it in summary.get("items") or []:
        ok = bool(it.get("ok"))
        items.append(_item(
            G, "chk_" + str(it.get("key") or it.get("no")),
            "第%s组 %s" % (it.get("no"), it.get("title")), ok,
            "总 %s / 差异 %s" % (it.get("total", 0), it.get("bad", 0)),
            ("检查本身报错：%s" % it.get("error")) if it.get("error")
            else ("%s 项差异" % it.get("bad")) if not ok else "全部一致",
            "按组内明细表逐行核对；先排除父科目重复、汇总层重复、期间错位、原币本币配反"))


# ---------------------------------------------------------------- 对外入口
def audit(conn, scope="all"):
    """清洗达标体检。

    scope="all"        全量（含五组检查结果）——页面上的「检查清洗达标」用这个
    scope="pre_merge"  只判清洗项（跳过 check_*）——合并门禁用这个：
                       合并刚完成时五组检查还没跑，拿它判会把所有人都拦下
    """
    items = []
    fns = (_audit_tb, _audit_jr) if scope == "pre_merge" else (_audit_tb, _audit_jr, _audit_check)
    for fn in fns:
        try:
            fn(conn, items)
        except Exception as e:
            items.append(_item("系统", "err_" + fn.__name__, "检查执行", False,
                               "异常", str(e)[:200], "把这条报给维护方"))

    # 只有 fail 级的不达标才否决整体结论；warn 单独列为「注意」
    fails = [i for i in items if not i["ok"] and i["level"] == "fail"]
    warns = [i for i in items if not i["ok"] and i["level"] == "warn"]
    # blocked：fail 里会阻断合并的那部分（导致假不符/假数据的硬问题）。
    # 辅助核算缺失这类「客户数据本身不完整」的 fail 只提示不阻断。
    blocked = [i for i in fails if i.get("block", True)]
    passed = len([i for i in items if i["ok"]])
    ok = not fails

    if ok and not warns:
        verdict = "达标：数据可以达到「完整、最细颗粒度」标准，可以进入合并与检查。"
    elif ok:
        verdict = ("基本达标，但有 %d 项需要注意（不阻塞后续流程）：%s"
                   % (len(warns), "；".join(i["title"] for i in warns)))
    elif not blocked:
        verdict = ("可以合并，但数据不完整（%d 项，不影响两表对账）：%s。"
                   "这些缺口会让底稿的账户/往来明细变粗，需要项目组知悉。"
                   % (len(fails), "；".join(i["title"] for i in fails)))
    else:
        verdict = ("未达标：%d 项必须修 —— %s。修完再点一次「检查清洗达标」。"
                   % (len(blocked), "；".join(i["title"] for i in blocked)))
    return {
        "ok": ok, "items": items, "prompt": PROMPT,
        "stat": {"total": len(items), "passed": passed,
                 "failed": len(fails), "warn": len(warns),
                 "blocked": len(blocked)},
        "verdict": verdict,
    }
