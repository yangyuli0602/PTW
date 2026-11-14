# -*- coding: utf-8 -*-
"""标准字段字典 + 列名自动匹配。

只做"猜"，猜错了人工改。规则按命中词组的完整度打分，再一对一贪心分配。
"""
import re

# ---------------------------------------------------------------- 序时账
JOURNAL_FIELDS = [
    # key,        中文标签,       类型,     是否常用
    ("vdate",     "日期",         "date",   1),
    ("period",    "会计期间",      "text",   0),
    ("vtype",     "凭证类型",      "text",   0),
    ("vno",       "凭证号",        "text",   1),
    ("voucher_key", "唯一凭证号",   "text",   0),
    ("summary",   "摘要",         "text",   1),
    ("acct_code", "科目编码",      "text",   1),
    ("acct_name", "科目名称",      "text",   1),
    ("acct_full", "科目全称",      "text",   0),
    ("acct_l1",   "一级科目",      "text",   0),
    ("debit",     "借方金额(本币)", "amount", 1),
    ("credit",    "贷方金额(本币)", "amount", 1),
    ("debit_fx",  "借方金额(原币)", "amount", 0),
    ("credit_fx", "贷方金额(原币)", "amount", 0),
    ("amount",    "金额(单列)",    "amount", 0),
    ("amount_fx", "原币金额(单列)", "amount", 0),
    ("direction", "借贷方向",      "text",   0),
    ("qty",       "数量",         "num",    0),
    ("price",     "单价",         "num",    0),
    ("currency",  "币种",         "text",   0),
    ("book_name", "账簿名称",      "text",   1),
    ("company_name", "公司名称",    "text",   0),
    ("dept",      "部门",         "text",   0),
    ("counterparty", "往来单位",   "text",   0),
    ("person",    "制单人",        "text",   0),
    ("checker",   "审核人",        "text",   0),
    ("poster",    "过账人",        "text",   0),
    ("attach",    "附单据数",      "text",   0),
    ("aux1",      "辅助核算1",     "text",   0),
    ("aux2",      "辅助核算2",     "text",   0),
    ("aux3",      "辅助核算3",     "text",   0),
]

# ---------------------------------------------------------------- 科目余额表
# 三层辅助核算：
#   位置槽 aux1/aux2/aux3  —— 原样承接任意导出格式的 1/2/3 号辅助核算列
#   角色标注（存 mapping["_aux"]，不占列）—— 说明这个槽是银行账号还是客商
#   语义列 bank_account 等  —— 合并时按角色从位置槽里抽出来，供直接 SQL 使用
# 外币（原币）金额字段：和 10 个本币金额一一对应，全部选填。
# 有外币核算的科目（银行存款外币户、外币往来）才用得上，没这几列也不影响合并。
FX_BASE = [
    ("open_debit",  "期初借方(外币)"),
    ("open_credit", "期初贷方(外币)"),
    ("open_bal",    "期初余额净额(外币)"),
    ("cur_debit",   "本期借方发生(外币)"),
    ("cur_credit",  "本期贷方发生(外币)"),
    ("ytd_debit",   "本年累计借方(外币)"),
    ("ytd_credit",  "本年累计贷方(外币)"),
    ("end_debit",   "期末借方(外币)"),
    ("end_credit",  "期末贷方(外币)"),
    ("end_bal",     "期末余额净额(外币)"),
]
FX_KEYS = tuple(k + "_fx" for k, _ in FX_BASE)

TB_FIELDS = [
    ("company_name", "公司名称",      "text",   1),
    ("book_name",  "账簿名称",        "text",   1),
    ("acct_code",  "科目编码",        "text",   1),
    ("acct_name",  "科目名称",        "text",   1),
    ("acct_full",  "科目全称",        "text",   0),
    ("acct_l1",    "一级科目",        "text",   1),
    ("level",      "科目级次",        "text",   0),
    ("currency",   "币种",            "text",   1),
    ("direction",  "余额方向",        "text",   0),
    ("open_debit", "期初借方(本币)",   "amount", 1),
    ("open_credit","期初贷方(本币)",   "amount", 1),
    ("open_bal",   "期初余额净额(本币)", "amount", 0),
    ("cur_debit",  "本期借方发生(本币)", "amount", 1),
    ("cur_credit", "本期贷方发生(本币)", "amount", 1),
    ("ytd_debit",  "本年累计借方(本币)", "amount", 1),
    ("ytd_credit", "本年累计贷方(本币)", "amount", 1),
    ("end_debit",  "期末借方(本币)",   "amount", 1),
    ("end_credit", "期末贷方(本币)",   "amount", 1),
    ("end_bal",    "期末余额净额(本币)", "amount", 0),
] + [(k + "_fx", lb, "amount", 0) for k, lb in FX_BASE] + [
    # ---- 辅助核算（都不是必备字段）
    ("aux1",       "辅助核算1",       "text",   0),
    ("aux2",       "辅助核算2",       "text",   0),
    ("aux3",       "辅助核算3",       "text",   0),
    # 开户行在账号上面：货币资金 BKD 的「银行/存款机构名称」直接吃这一列，
    # 只给账号不给行名时，一眼看不出是哪个行的哪个户。
    ("bank_name",         "银行开户行名", "text",  0),
    ("bank_account",      "银行账号",  "text",  0),
    ("counterparty_name", "客商名称",  "text",  0),
    ("dept_name",         "部门",     "text",  0),
    ("project_name",      "项目",     "text",  0),
    ("person_name",       "职员",     "text",  0),
    ("item_name",         "存货",     "text",  0),
]

# ---------------------------------------------------------------- 银行流水
# 和序时账/余额表并列的第三类数据源。设计要点：
#   acct_no / counterparty_acct 是主键性质的两列——前者决定"这笔流水属于哪个账户"，
#   后者决定"这笔转账是不是本公司的内部划转"。两列都缺，这条流水就没有分析价值。
#   金额支持单列（amount，带方向）和双列（income/expense）两种银行导出习惯。
FLOW_FIELDS = [
    ("acct_no",           "本方账号",    "text",   1),
    ("acct_name",         "本方户名",    "text",   1),
    ("bank_name",         "本方开户行",  "text",   0),
    ("company_name",      "公司名称",    "text",   0),
    ("txn_date",          "交易日期",    "date",   1),
    ("txn_time",          "交易时间",    "text",   0),
    ("serial",            "交易流水号",  "text",   0),
    ("direction",         "收付方向",    "text",   1),
    ("income",            "收入金额",    "amount", 0),
    ("expense",           "支出金额",    "amount", 0),
    ("amount",            "交易金额",    "amount", 1),
    ("balance",           "账户余额",    "amount", 0),
    ("currency",          "币种",        "text",   0),
    ("counterparty_acct", "对方账号",    "text",   1),
    ("counterparty_name", "对方户名",    "text",   1),
    ("counterparty_bank", "对方开户行",  "text",   0),
    ("summary",           "摘要用途",    "text",   1),
    ("channel",           "交易渠道",    "text",   0),
    ("remark",            "备注",        "text",   0),
]

# ---------------------------------------------------------------- 固定资产台账
# 和序时账/余额表并列的第四类数据源：客户资产清单（FA list）。
# 设计要点：
#   category_code / category_name / asset_group 三列是"分类三元组"——
#     台账可能只给编码（G0100）、只给名称（固定资产-办公用房）、或两者都给，
#     也可能自带大类（机器设备）。三者都可空，缺的由 asset_map 推断。
#   orig_cost / acc_dep / impair / reval / net_book 五个金额：
#     不同客户台账口径不同（购置值 / 原值 / 累计折旧 / 减值 / 重估 / 净值），
#     五个都留位置，用哪个由生成时按"哪个有值"决定，不在字段层写死。
#   fa / fa_beg 是同一套字段的两张表：前者期末，后者期初。
FA_FIELDS = [
    ("company_name",  "公司名称",     "text",   0),
    ("acct_code",     "总账科目",     "text",   0),
    ("category_code", "分类编码",     "text",   1),
    ("category_name", "分类名称",     "text",   0),
    ("asset_group",   "资产类型",     "text",   0),
    ("asset_code",    "资产编码",     "text",   0),
    ("asset_name",    "资产名称",     "text",   0),
    ("dept_name",     "使用部门",     "text",   0),
    ("location",      "存放地点",     "text",   0),
    ("start_date",    "启用日期",     "date",   0),
    ("disposal_date", "处置日期",     "date",   0),
    ("life_months",   "使用寿命(月)", "num",    0),
    ("residual_rate", "残值率",       "num",    0),
    ("orig_cost",     "原值",         "amount", 1),
    ("acc_dep",       "累计折旧",     "amount", 1),
    ("impair",        "减值准备",     "amount", 0),
    ("reval",         "重估/调整",    "amount", 0),
    ("net_book",      "账面净值",     "amount", 0),
    ("status",        "使用状态",     "text",   0),
]

# 辅助核算角色：key -> 中文名。角色存在 mapping["_aux"] 里，形如 {"aux1": "bank_account"}
AUX_ROLES = [
    ("bank_name",         "银行开户行名"),
    ("bank_account",      "银行账号"),
    ("counterparty_name", "客商"),
    ("dept_name",         "部门"),
    ("project_name",      "项目"),
    ("person_name",       "职员"),
    ("item_name",         "存货"),
    ("mixed",             "混合（一格多类型，按「客户:xx」前缀自动识别）"),
    ("other",             "其他"),
]
AUX_ROLE_LABEL = dict(AUX_ROLES)
# 角色 -> 该角色在「类型:名称」里的类型关键词（正则 alternation）
ROLE_KEYS = {
    # 「账号」类只认账号词；「开户行」单独一类——以前开户行也归 bank_account，
    # 结果「开户行」那一格被当成账号去抽数字，抽不出来还把真正的账号列挤掉。
    "bank_name":         "开户行|开户银行|银行名称|支行|分行|银行",
    "bank_account":      "银行账号|银行帐号|银行账户|账号|帐号|卡号",
    "counterparty_name": "客户|供应商|客商|往来单位|往来户|对方单位|单位名称",
    "dept_name":         "部门|成本中心|事业部|车间",
    "project_name":      "项目|工程|课题",
    "person_name":       "职员|员工|职工|人员|经办人|个人",
    "item_name":         "存货|物料|商品|产品",
}
# 位置槽 -> 只有这三个 key 参与角色抽取
AUX_SLOTS = ("aux1", "aux2", "aux3")

# 字段分组：main 主表字段 / fx 外币金额 / aux 辅助核算
FIELD_GROUP = {k: "aux" for k in (
    "aux1", "aux2", "aux3", "bank_name", "bank_account", "counterparty_name",
    "dept_name", "project_name", "person_name", "item_name")}
FIELD_GROUP.update({k: "fx" for k in FX_KEYS})
# 序时账也有原币金额：和本币 debit/credit 对应，选填
FIELD_GROUP.update({"debit_fx": "fx", "credit_fx": "fx"})

TARGETS = {
    "journal": {"label": "序时账", "fields": JOURNAL_FIELDS},
    "tb":      {"label": "科目余额表", "fields": TB_FIELDS},
    "flow":    {"label": "银行流水", "fields": FLOW_FIELDS},
    # 固定资产台账：同一套字段挂两个 target，因为期初/期末是两份文件、两张表。
    # 合成一张会互相覆盖（merge 按 target 建表，后合并的盖掉先合并的）。
    "fa":      {"label": "固定资产台账（期末）", "fields": FA_FIELDS},
    "fa_beg":  {"label": "固定资产台账（期初）", "fields": FA_FIELDS},
}


def field_dict(target):
    return [{"key": k, "label": lb, "type": t, "common": bool(c),
             "group": FIELD_GROUP.get(k, "main")}
            for k, lb, t, c in TARGETS[target]["fields"]]


# ---------------------------------------------------------------- 导出用中文表头
# 标准表里还有几个不在字段字典里的内部列，单独给中文名。
EXTRA_LABEL = {
    "src": "来源",
    "src_row": "来源行号",
    "aux_text": "辅助核算全文",
    "short_name": "公司简称",
}

# ---------------------------------------------------------------- 核对结果表头
# 核对结果是自己拼的一张表（不是 tb / journal），列名同样走 to_cn 翻中文。
RECON_LABEL = {
    "entity": "主体(公司/账簿)",
    "currency": "币别",
    "acct_code": "一级科目编码",
    "acct_name": "一级科目名称",
    "acct_full": "科目全称",
    # ---- 余额表自平衡（本币）
    "tb_cur_debit": "本期借(本币)",
    "tb_cur_credit": "本期贷(本币)",
    "tb_ytd_debit_sum": "本年累计借(本币)",
    "tb_ytd_credit_sum": "本年累计贷(本币)",
    "tb_ytd_diff": "本年累计借贷差",
    "tb_open_net": "期初余额(本币)",
    "tb_end_net": "期末余额(本币)",
    "tb_roll_diff": "余额勾稽差(本币)",
    # ---- 余额表自平衡（原币）
    "tb_cur_debit_fx": "本期借(原币)",
    "tb_cur_credit_fx": "本期贷(原币)",
    "tb_open_net_fx": "期初余额(原币)",
    "tb_end_net_fx": "期末余额(原币)",
    "tb_roll_diff_fx": "余额勾稽差(原币)",
    "n_tb_rows": "余额表行数",
    # ---- 序时账自平衡
    "jr_debit_sum": "序时账借(本币)",
    "jr_credit_sum": "序时账贷(本币)",
    "jr_debit_fx": "序时账借(原币)",
    "jr_credit_fx": "序时账贷(原币)",
    "jr_diff": "借贷差(本币)",
    "jr_diff_fx": "原币差(仅展示)",
    "n_jr_rows": "序时账行数",
    "voucher_key": "唯一凭证号",
    "period": "会计期间",
    "vno": "凭证号",
    "n_vouchers": "凭证张数",
    "n_bad_vouchers": "不平凭证数",
    "n_months": "月份数",
    "n_bad_months": "不平月份数",
    "tb_ytd_debit": "余额表借方(本币)",
    "tb_ytd_credit": "余额表贷方(本币)",
    "tb_ytd_net": "余额表净额(本币)",
    "jr_ytd_debit": "序时账借方(本币)",
    "jr_ytd_credit": "序时账贷方(本币)",
    "jr_ytd_net": "序时账净额(本币)",
    "diff_debit": "借方差异(本币)",
    "diff_credit": "贷方差异(本币)",
    "diff_net": "净额差异",
    "tb_ytd_debit_fx": "余额表借方(原币)",
    "tb_ytd_credit_fx": "余额表贷方(原币)",
    "tb_ytd_net_fx": "余额表净额(原币)",
    "jr_ytd_debit_fx": "序时账借方(原币)",
    "jr_ytd_credit_fx": "序时账贷方(原币)",
    "jr_ytd_net_fx": "序时账净额(原币)",
    "diff_debit_fx": "借方差异(原币)",
    "diff_credit_fx": "贷方差异(原币)",
    "diff_net_fx": "净额差异(原币)",
    "status": "核对结果",
    "n_tb": "余额表行数",
    "n_jr": "序时账行数",
}


def _build_labels(target=None):
    """组装"英文列名 -> 中文表头"。

    direction 在两张表里含义不同（序时账=借贷方向、余额表=余额方向），
    所以优先用目标表自己的字典；不知道目标表时按余额表优先合并。
    target="recon" 是核对结果表：列名带 tb_/jr_ 前缀，自己有全套中文名。
    """
    if target == "recon":
        m = dict(EXTRA_LABEL)
        m.update(RECON_LABEL)
        return m
    m = dict(EXTRA_LABEL)
    order = [target] if target in TARGETS else ["tb", "journal"]
    for t in order:
        for _k, _lb, _ty, _c in TARGETS[t]["fields"]:
            m.setdefault(_k, _lb)
    return m


# 默认（合并版）对照表；导出时按实际目标表再取一次精确版
LABELS = _build_labels()
_LABELS_BY_TARGET = {t: _build_labels(t) for t in TARGETS}
_LABELS_BY_TARGET["recon"] = _build_labels("recon")


def to_cn(name, target=None):
    """内部列名转中文表头。

    只有"整名完全等于某个标准列名"才翻译，用户自己起的别名（如 AS 编码）
    原样返回——这样手写 SQL 的别名不会被系统抢改。
    """
    if name is None:
        return ""
    s = str(name).strip()
    m = _LABELS_BY_TARGET.get(target, LABELS)
    return m.get(s, s)


def cn_headers(cols, target=None):
    """批量转中文表头，重名自动加 (2)(3) 后缀，避免两列都叫"部门"。"""
    out, seen = [], {}
    for c in cols:
        lb = to_cn(c, target)
        n = seen.get(lb, 0) + 1
        seen[lb] = n
        out.append(lb if n == 1 else "%s(%d)" % (lb, n))
    return out


def table_columns(target):
    """标准表有哪些列 → [(列名, 中文名)]，顺序照标准字段定义走。

    给「按列值删除行」这类手工操作用：下拉里要显示中文，
    不能让人去记 acct_code / company_name 这种内部列名。
    """
    fs = (TARGETS.get(target) or {}).get("fields") or []
    out, seen = [], set()
    for k, lb, _ty, _c in fs:
        if k in seen:
            continue
        seen.add(k)
        out.append((k, lb))
    return out


def _norm(s):
    """归一化列名：去空白/标点/括号内容，全角转半角，统一小写。"""
    if s is None:
        return ""
    s = str(s)
    s = s.translate(str.maketrans(
        "（）：，．－　", "()::.- "))
    s = re.sub(r"[\(\)\[\]【】<>《》\s_\-\.:：/\\|,，、'\"]", "", s)
    return s.lower()


# 每个字段的匹配词组：其中每个 tuple 是一个"概念"，概念内元素是任一即可(OR)，
# 所有概念之间是 AND。权重用于区分优先级。
JOURNAL_RULES = {
    "vdate":     [(("日期", "业务日期", "凭证日期", "记账日期", "date"), 3)],
    "period":    [(("期间", "会计期间", "月份", "所属期间"), 2)],
    "vtype":     [(("凭证类型", "凭证类别", "凭证种类"), 3)],
    "vno":       [(("凭证号", "凭证字号", "凭证编号", "voucherno", "凭证号码"), 3)],
    "voucher_key": [(("唯一凭证号", "唯一凭证", "凭证唯一号", "凭证主键"), 5)],
    "summary":   [(("摘要", "业务说明", "说明", "事由", "备注内容"), 3)],
    "acct_code": [(("科目编码", "科目代码", "科目号", "会计科目编码", "科目编号", "accountcode"), 3)],
    "acct_name": [(("科目名称", "会计科目名称", "科目", "账户名称", "accountname"), 2)],
    "acct_full": [(("科目全称", "完整科目", "科目路径", "全称科目", "科目全路径"), 5)],
    "acct_l1":   [(("一级科目", "总账科目", "大类科目", "一级", "L1"), 5)],
    # 主体：多主体账套靠它隔离配对与填数，缺了会把不同公司的凭证配到一起
    "company_name": [(("公司名称", "单位名称", "主体名称", "核算主体", "公司全称",
                       "公司", "company"), 2)],
    # 银行账号：序时账里常单独一列，直接映射比走辅助核算更省事
    "bank_account": [(("银行账号", "银行帐号", "银行账户", "账号", "帐号", "卡号"), 2)],
    "debit":     [(("借方", "借", "debit"), 2), (("金额", "发生额", "amount", ""), 1)],
    "credit":    [(("贷方", "贷", "credit"), 2), (("金额", "发生额", "amount", ""), 1)],
    "amount":    [(("金额", "发生额", "amount"), 2)],
    "amount_fx": [(("原币金额", "外币金额", "原币发生额", "amountfx"), 5)],
    "direction": [(("方向", "借贷", "借贷标志"), 2)],
    "qty":       [(("数量", "qty", "件数"), 2)],
    "price":     [(("单价", "price"), 2)],
    "currency":  [(("币种", "货币", "currency"), 2)],
    "book_name": [(("账簿名称", "账套名称", "账簿", "账套", "核算账套", "book"), 3)],
    "dept":      [(("部门", "成本中心", "事业部"), 2)],
    "counterparty": [(("往来单位", "客户名称", "供应商", "客商", "往来户", "对方单位"), 2)],
    "person":    [(("制单", "制表", "录入", "经办人"), 2)],
    "checker":   [(("审核人", "审核", "复核"), 2)],
    "poster":    [(("过账", "记账人"), 2)],
    "attach":    [(("附单据", "单据数", "附件数"), 2)],
    # ---- 辅助核算位置槽：与科目余额表一致，只认序号明确的列，泛指列留给 aux1
    "aux1":      [],
    "aux2":      [],
    "aux3":      [],
}

TB_RULES = {
    "company_name":[(("公司名称", "企业名称", "主体名称", "核算主体", "单位名称",
                      "公司简称", "company"), 2)],
    "acct_code":   [(("科目编码", "科目代码", "科目号", "科目编号", "accountcode"), 3)],
    "acct_name":   [(("科目名称", "账户名称", "科目"), 2)],
    "acct_full":   [(("科目全称", "完整科目", "科目路径", "全称科目", "科目全路径"), 5)],
    "acct_l1":     [(("一级科目", "总账科目", "大类科目", "一级", "L1"), 5)],
    "level":       [(("级次", "级别", "层次", "level"), 2)],
    "book_name":   [(("账簿名称", "账套名称", "账簿", "账套", "核算账套", "book"), 3)],
    "currency":    [(("币种", "货币", "currency"), 3)],
    "direction":   [(("方向", "余额方向", "借贷方向"), 2)],
    "open_debit":  [(("期初", "年初", "上期"), 2), (("借", "debit"), 2)],
    "open_credit": [(("期初", "年初", "上期"), 2), (("贷", "credit"), 2)],
    "open_bal":    [(("期初", "年初", "上期"), 2), (("余额",), 1)],
    "cur_debit":   [(("本期", "本月", "本期发生", "本期发生额", "发生额"), 2), (("借", "debit"), 2)],
    "cur_credit":  [(("本期", "本月", "本期发生", "本期发生额", "发生额"), 2), (("贷", "credit"), 2)],
    # 本年累计的权重必须高于「本期发生额」：像「本年累计借方发生额」这种列名，
    # cur_debit 靠「发生额」也能命中，两者会打成同分，再按规则表顺序 cur 在前就抢走了
    # ——结果本年累计数被当成本期数，凡是用 ytd 的分析全错。
    "ytd_debit":   [(("本年", "累计", "本年累计", "年初至今", "本年发生"), 3), (("借", "debit"), 2)],
    "ytd_credit":  [(("本年", "累计", "本年累计", "年初至今", "本年发生"), 3), (("贷", "credit"), 2)],
    "end_debit":   [(("期末", "本期期末", "月末", "余额"), 2), (("借", "debit"), 2)],
    "end_credit":  [(("期末", "本期期末", "月末", "余额"), 2), (("贷", "credit"), 2)],
    "end_bal":     [(("期末", "月末"), 2), (("余额",), 1)],
    # ---- 辅助核算位置槽：只认序号明确的列，泛指的那一列留给 aux1
    "aux1":        [],
    "aux2":        [],
    "aux3":        [],
    # ---- 语义列：某些系统直接给了独立列，能认就直接认
    "bank_account":      [(("银行账号", "银行账户", "开户行", "账号"), 3)],
    "counterparty_name": [(("客户名称", "供应商名称", "客商名称", "往来单位名称"), 3)],
    "dept_name":         [(("部门名称", "部门"), 2)],
    "project_name":      [(("项目名称", "项目"), 2)],
    "person_name":       [(("职员", "员工", "职工", "个人"), 2)],
    "item_name":         [(("存货名称", "存货", "商品名称"), 2)],
}
# 外币字段的认词和本币完全一样，靠"外币/原币"标记在 auto_map 里做闸门分开
for _k, _ in FX_BASE:
    TB_RULES[_k + "_fx"] = TB_RULES[_k]
# 序时账的原币借贷：认词同本币借贷，同样靠闸门分开
JOURNAL_RULES["debit_fx"] = JOURNAL_RULES["debit"]
JOURNAL_RULES["credit_fx"] = JOURNAL_RULES["credit"]

# ---------------------------------------------------------------- 银行流水
# 这张表的难点全在"本方"和"对方"要分开：银行导出的表头经常只写「账号」「开户行」，
# 而同一张表里还有一个「对方账号」「对方开户行」。auto_map 是一列只能用一次的贪心，
# 所以给"对方"那几条加更高的权重——谁分高谁先拿走，剩下的才归本方。
FLOW_RULES = {
    # ---- 本方账户
    # 只列"账号"类词。"账户"不能放进来——它会让「账户名称」被认成本方账号，
    # 把户名列抢走；宁可认不出让人手工映射。
    "acct_no":       [(("本方账号", "本方帐号", "本方账户", "查询账号", "客户账号",
                        "账户号码", "银行账号", "帐号", "账号", "卡号"), 5)],
    "acct_name":     [(("本方户名", "本方账户名称", "本方名称", "开户名称",
                        "账户名称", "户名", "客户名称"), 4)],
    "bank_name":     [(("本方开户行", "开户行", "开户银行", "银行名称"), 3)],
    "company_name":  [(("公司名称", "单位名称", "主体名称", "企业名称"), 2)],
    # ---- 交易要素
    "txn_date":      [(("交易日期", "记账日期", "入账日期", "业务日期",
                        "交易日", "日期", "交易时间"), 3)],
    "txn_time":      [(("交易时刻", "时分秒"), 2)],
    "serial":        [(("流水号", "交易编号", "业务编号", "交易参考号",
                        "回单编号", "凭证号", "交易序号"), 3)],
    "direction":     [(("收付方向", "收支方向", "收付标志", "借贷标志",
                        "借贷方向", "收支标志", "方向", "收支"), 3)],
    "income":        [(("收入金额", "贷方发生额", "贷方金额", "存入金额",
                        "进账金额", "收款金额", "收入", "存入", "贷方"), 4)],
    "expense":       [(("支出金额", "借方发生额", "借方金额", "支取金额",
                        "出账金额", "付款金额", "支出", "支取", "借方"), 4)],
    "amount":        [(("交易金额", "发生额", "金额"), 2)],
    "balance":       [(("账户余额", "交易后余额", "余额"), 3)],
    "currency":      [(("币种", "货币", "currency"), 3)],
    # ---- 对方（权重高于本方，保证"对方账号"不被我方抢走）
    "counterparty_acct": [(("对方账号", "对方帐号", "对方账户", "对手账号",
                            "收方账号", "付方账号", "收款方账号", "付款方账号",
                            "收款人账号", "付款人账号"), 6)],
    "counterparty_name": [(("对方户名", "对方账户名称", "对方单位名称", "对方名称",
                            "对方客户名称", "收款人名称", "付款人名称",
                            "收款人", "付款人", "对方"), 6)],
    "counterparty_bank": [(("对方开户行", "对方开户银行", "对方银行",
                            "收方开户行", "付方开户行"), 6)],
    # ---- 其他
    "summary":       [(("摘要", "用途", "附言", "交易用途", "业务摘要", "说明"), 3)],
    "channel":       [(("渠道", "交易渠道", "交易方式", "交易类型"), 2)],
    "remark":        [(("备注",), 2)],
}

# ---------------------------------------------------------------- 固定资产台账
# 客户台账列名千差万别（购置值/原值/原价、净帐面值/净值/净额…），这里只收"能确定"的词。
# 两条纪律：
#   1) impair 不接裸词「减值」——期初表有「上期末减值」「当期计提减值」两列，
#      裸词会把它们抢成减值准备，导致一张表里三个减值列争同一个字段。
#   2) net_book 权重高于 orig_cost——「当前净帐面值」同时含"值"和"净"，不压住会被原值抢走。
FA_RULES = {
    "company_name":  [(("公司名称", "单位名称", "主体名称", "企业名称"), 2)],
    "acct_code":     [(("总账科目", "科目编码", "会计科目", "科目代码"), 3)],
    # ---- 分类三元组
    # 注意：_score 里多个 concept 是 AND（要同时命中），一个 concept 内多个词才是 OR。
    # 编码列只认"编码/代码/码"——不能带裸「分类」：客户常用「资产分类」装的是**名称**
    # （固定资产-办公用房），裸词会让名称被抢进编码列。
    # 权重必须高于 category_name：精确词「分类编码」要先于泛词「分类」被选中，
    # 否则「分类编码」列会被名称规则抢走，编码列永远空着。
    "category_code": [(("分类编码", "分类代码", "资产分类代码", "类别编码",
                       "资产编码类别", "分类码", "代码"), 8)],
    "category_name": [(("细分类型", "分类名称", "资产类别名称", "明细类别", "类别名称",
                        "资产分类", "固定资产类别", "资产类别", "明细分类",
                        "分类", "类别"), 6)],
    "asset_group":   [(("资产类型", "资产大类", "资产组", "固定资产类别"), 5)],
    # ---- 资产标识
    "asset_code":    [(("资产编号", "资产编码", "卡片编号", "固定资产编号",
                        "设备编号", "资产号"), 4)],
    "asset_name":    [(("资产名称", "固定资产名称", "设备名称", "资产描述"), 4)],
    "dept_name":     [(("使用部门", "管理部门", "使用单位", "部门"), 3)],
    "location":      [(("存放地点", "使用地点", "存放位置", "地点"), 3)],
    # ---- 日期与折旧参数
    "start_date":    [(("启用日期", "开始使用日期", "资本化日期", "开始使用",
                        "入账日期", "购置日期"), 4)],
    "disposal_date": [(("处置日期", "报废日期", "减少日期", "清理日期"), 4)],
    "life_months":   [(("使用寿命", "折旧年限", "使用年限", "使用月份", "摊销年限"), 3)],
    "residual_rate": [(("残值率", "预计净残值率", "残值比例"), 4)],
    # ---- 金额（顺序即优先级：净额优先，避免被原值抢）
    "net_book":      [(("当前净帐面值", "净帐面值", "账面净值", "帐面净值",
                        "资产净值", "净值", "净额"), 6)],
    "orig_cost":     [(("年末购置值", "购置值", "资产原值", "原值", "原价",
                        "购置价值", "购买价值", "入账价值"), 5)],
    "acc_dep":       [(("累计折旧", "已提折旧", "折旧累计", "累计折旧额"), 5)],
    "impair":        [(("减值准备", "资产减值准备", "减值准备金额"), 4)],
    "reval":         [(("会计年度末重估", "末重估", "重估增值", "重估", "评估增值"), 4)],
    "status":        [(("使用状态", "资产状态", "卡片状态"), 3)],
}

RULES = {"journal": JOURNAL_RULES, "tb": TB_RULES, "flow": FLOW_RULES,
         "fa": FA_RULES, "fa_beg": FA_RULES}

# 列名里出现这些词 → 视为外币(原币)金额列
FX_HINT = re.compile(r"外币|原币|外汇|foreign|美元|港币|欧元|日元|英镑")


def _score(ncol, concepts):
    """返回 (score, 命中概念数)。未命中全部概念返回 -1。"""
    total = 0
    hit = 0
    for words, w in concepts:
        matched = False
        for wd in words:
            if wd and wd in ncol:
                matched = True
                break
        if matched:
            hit += 1
            total += w
        elif w == 0:
            continue
        else:
            return -1, 0
    # 概念都命中才有效；词越短越可能误命中，给一点长度惩罚
    return total - len(ncol) * 0.01, hit


def auto_map(target, columns):
    """把原始列名自动映射到标准字段。返回 {field: column}。

    一对一贪心：先给高分字段选列，选过的列不再复用。
    科目余额表的辅助核算位置槽单独处理（辅助核算1/2/3 → aux1/2/3，泛指的一列 → aux1）。
    """
    rules = RULES.get(target)
    if not rules:
        return {}
    norms = [(_norm(c), c) for c in columns]
    # 外币闸门：列名带"外币/原币"的只准进 *_fx 字段，本币金额字段反过来不接它们。
    # 没有这一层，"期初借方(外币)"会被同样认"期初+借"的 open_debit 抢走。
    gate = any(f.endswith("_fx") for f in rules)
    amount_f = {k for k, _lb, ty, _c in TARGETS[target]["fields"] if ty == "amount"}
    cands = []
    for f, concepts in rules.items():
        mode = -1                      # -1 不设限 / 0 只收本币 / 1 只收外币
        if gate:
            if f.endswith("_fx"):
                mode = 1
            elif f in amount_f:
                mode = 0
        for i, (nc, raw) in enumerate(norms):
            if mode >= 0:
                is_fx = bool(FX_HINT.search(nc))
                if mode == 1 and not is_fx:
                    continue
                if mode == 0 and is_fx:
                    continue
            sc, hit = _score(nc, concepts)
            if sc > 0:
                cands.append((sc, hit, f, i))
    cands.sort(key=lambda x: (-x[0], -x[1]))
    used_col = set()
    used_field = set()
    out = {}
    for sc, hit, f, i in cands:
        if f in used_field or i in used_col:
            continue
        out[f] = columns[i]
        used_field.add(f)
        used_col.add(i)
    if target in ("tb", "journal"):
        for slot, i in _auto_aux(norms).items():
            if slot in used_field or i in used_col:
                continue
            out[slot] = columns[i]
            used_field.add(slot)
            used_col.add(i)
    return out


_AUX_SEQ = re.compile(r"(?:辅助核算|核算项目|辅助项|aux)[_\-]?([1-3一二三])")
_AUX_ANY = re.compile(r"(辅助核算|核算项目|辅助项|辅助$|auxiliary)")


def _auto_aux(norms):
    """返回 {slot_index: 列下标}。序号明确的优先，剩下泛指的列给 aux1。"""
    hit = {}
    generic = []
    for i, (nc, raw) in enumerate(norms):
        m = _AUX_SEQ.search(nc)
        if m:
            ch = m.group(1)
            n = {"一": 1, "二": 2, "三": 3}.get(ch, ch)
            hit["aux%s" % n] = i
        elif _AUX_ANY.search(nc):
            generic.append(i)
    if generic:
        for slot in AUX_SLOTS:
            if slot not in hit:
                hit[slot] = generic.pop(0)
                if not generic:
                    break
    return hit


def detect_header(rows, target=None):
    """在前若干行里猜表头位置，返回 (表头行号 1-based, 表头行数)。

    评分：非空格子数 + 命中标准字段词的数量 + 全为文本行的加成 - 含数字行的惩罚。
    如果胜出的行左侧有空单元格（典型的第二层表头：'' '' 借方 贷方），
    则把上一行一起算作表头（两层表头）。
    """
    best_i, best_sc = 1, -1e9
    rules = RULES.get(target) if target else None
    keywords = set()
    for tgt in (["journal", "tb", "flow"] if not rules else [target]):
        for concepts in RULES[tgt].values():
            for words, _ in concepts:
                for w in words:
                    if w:
                        keywords.add(w)
        if rules:
            break
    for idx, row in enumerate(rows[:30]):
        cells = [str(c).strip() for c in row if c is not None and str(c).strip() != ""]
        if not cells:
            continue
        # 被铺平的合并单元格：整行 ≥3 格却只有一种内容（导出工具把「科目余额表」铺满 9 格）。
        # 这类是标题行，不是表头。必须**直接跳过**而不是扣分——9 个格子都含「科目」，
        # 关键词加分能加 36 分，扣 10 分根本压不住（实测仍被判成第 1 行）。
        if len(cells) >= 3 and len({str(c).strip() for c in cells}) == 1:
            continue
        sc = len(cells) * 1.0
        kn = 0
        for c in cells:
            nc = _norm(c)
            if any(k in nc for k in keywords):
                kn += 1
        sc += kn * 4
        if len(cells) == 1:
            sc -= 3          # 单格大标题
        numish = sum(1 for c in cells if re.fullmatch(r"-?[\d,\.]+", c))
        sc -= numish * 2     # 数据行通常是数字
        if idx == 0:
            sc += 0.5
        if sc > best_sc:
            best_sc, best_i = sc, idx + 1

    # 两层表头识别：
    # 1) 胜出行左侧有空单元格，且上一行有内容（合并单元格被空列顶出的第二层表头）
    # 2) 胜出行本身就是「借/贷/期初/期末」二级行 → 上一行是分类表头
    # 3) 胜出行的下一行是「借/贷/期初/期末」二级行 → 胜出行是分类表头
    #    （导出工具把合并单元格拆开、前两列也写了文字时，判据 1 会失效，靠 2/3 兜底）
    # 「方向 / 期间点 / 期间段」三类都算二级或分类表头的词。
    # 加「本期/本年/累计」是必要的：分类层常写「期初余额 | 本期发生额 | 本年累计」，
    # 少这两个词就认不出这一行是分类层（实测 rows 用例退化成 (4,1)）。
    # 它们会不会把单层表头误判成二级层？不会——判据看的是**上一行**，
    # 单层表头的上一行是标题（1 格），`len(_has_content) >= 2` 直接挡住。
    SECOND = ("借方", "贷方", "借", "贷", "期初", "期末", "年初", "年末",
              "本期", "本年", "累计", "上月", "本月")
    _SEC_SUFFIX = ("发生额", "余额", "金额", "本位币", "原币", "数", "合计")
    # 出现这些词说明这一行是「主表头」（带着科目编码、名称这些主体列），不是二级层
    _SUBJ = ("科目", "编码", "名称", "账簿", "摘要", "凭证", "级次")

    def _is_second_cell(c):
        """这一格像不像二级表头的「方向 / 期间」格子。

        2026-11-11 扩：原来要求整格恰好等于 SECOND，遇到「借方金额」「期初余额(借)」
        就认不出来——双层表头会按单层读，列名整体错位。
        （金斯耐两张海外余额表第二层写的就是「借方金额 / 贷方金额」，平台认成第 1 行
        大标题，只能手工指定 `header_row=3, header_rows=2`。）
        现在先去掉括号内容和「金额 / 余额 / 发生额」这类口径后缀再比对。
        """
        t = re.sub(r"\s+", "", str(c or ""))
        if not t or len(t) > 12:
            return False
        if t in SECOND:
            return True
        core = re.sub(r"[（(][^）)]*[）)]", "", t)      # 期初余额(借) → 期初余额
        for suf in _SEC_SUFFIX:                         # 借方金额 → 借方
            core = core.replace(suf, "")
        core = core.strip("()（）_-—·:：、,，")
        return bool(core) and core in SECOND

    def _is_second(row):
        """整行像不像「二级表头」：≥2 格是方向/期间词，**且不含科目编码/名称这类主体列**。

        主体词闸门是关键（2026-11-11 补）：单层表头会把
        「科目编码 | 科目名称 | 期初本位币(借) | 期初本位币(贷)…」写在同一行，
        只数方向词会把它误判成二级行，进而把上面那行标题也并进来——
        实测 20 个样本从 (2,1) 变成 (1,2)，整行错位。
        """
        if not row:
            return False
        cells = [c for c in row if c is not None and str(c).strip()]
        if any(any(w in str(c) for w in _SUBJ) for c in cells):
            return False
        return sum(1 for c in cells if _is_second_cell(c)) >= 2

    def _has_content(row):
        return [c for c in (row or []) if c is not None and str(c).strip() != ""]

    win = rows[best_i - 1] if best_i - 1 < len(rows) else []
    # 0) 上一行与胜出行「前缀相同、后面不同」→ 标准双层表头。
    #    导出工具会把合并单元格铺平，两行都重复写着「序号|科目编码|科目名称|币种|账簿|核算维度」，
    #    所以既不能只数方向词（会被主体列稀释），也不能要求二级行没有主体列
    #    （金斯耐两张表第 3 行是分类、第 4 行是明细，主体列两行都有）。
    #    只有"前缀一长串相同、尾巴不同"才是这个结构的特征。
    if best_i > 1:
        above = rows[best_i - 2]
        pre = 0
        for x, y in zip(above, win):
            sx = str(x or "").strip()
            sy = str(y or "").strip()
            if sx and sy and sx == sy:
                pre += 1
            else:
                break
        same = [str(c or "").strip() for c in above] == [str(c or "").strip() for c in win]
        if pre >= 2 and not same:
            return best_i - 1, 2
    # 2) 胜出行的**上一行**是纯分类行（只有期间/方向词、没有主体列）→ 两层
    if best_i > 1:
        above = rows[best_i - 2]
        if _is_second(above) and len(_has_content(above)) >= 2:
            return best_i - 1, 2
    # 3) 胜出行与**下一行**前缀相同、后面不同 → 胜出行是分类层、下一行是明细层
    #    （与判据 0 对称：铺平后分类行和明细行都带主体列，谁当胜出行都可能）
    if best_i < len(rows):
        nxt = rows[best_i]
        pre = 0
        for x, y in zip(win, nxt):
            sx = str(x or "").strip()
            sy = str(y or "").strip()
            if sx and sy and sx == sy:
                pre += 1
            else:
                break
        same = [str(c or "").strip() for c in win] == [str(c or "").strip() for c in nxt]
        if pre >= 2 and not same:
            return best_i, 2
    # 4) 胜出行的下一行是纯二级行 → 胜出行是分类表头
    if best_i < len(rows) and _is_second(rows[best_i]):
        return best_i, 2
    # 1) 胜出行左侧有空单元格，且上一行有内容
    lead_empty = 0
    for c in win:
        if c is None or str(c).strip() == "":
            lead_empty += 1
        else:
            break
    if lead_empty > 0 and best_i > 1:
        above = [str(c).strip() for c in rows[best_i - 2] if c is not None]
        if any(above):
            return best_i - 1, 2
    return best_i, 1


# ---------------------------------------------------------------- 辅助核算角色
def role_from_header(name):
    """列名本身带角色词就直接认（科目余额表里「客户名称」这种列）。"""
    nc = _norm(name)
    if not nc:
        return ""
    best, best_len = "", 0
    for key, words in ROLE_KEYS.items():
        for w in words.split("|"):
            if w and w in nc and len(w) > best_len:
                best, best_len = key, len(w)
    return best


_PREFIX_SPLIT = re.compile(r"[/／、;；|,，]")
_COLON = re.compile(r"[:：]")


def role_from_values(values):
    """从样本值里猜角色：值形如「客户:阿里巴巴/部门:销售部」时按前缀统计。

    返回 (role, 命中率)。没有前缀形态的列返回 ("", 0)，由人工指定。
    """
    hits = {}
    total = 0
    for v in values:
        s = str(v or "").strip()
        if not s:
            continue
        total += 1
        if not _COLON.search(s):
            continue
        for part in _PREFIX_SPLIT.split(s):
            p = part.strip()
            if not _COLON.search(p):
                continue
            prefix = _COLON.split(p, 1)[0].strip()
            for key, words in ROLE_KEYS.items():
                if any(w and w in prefix for w in words.split("|")):
                    hits[key] = hits.get(key, 0) + 1
                    break
    if not hits or not total:
        return "", 0.0
    key = max(hits, key=lambda k: hits[k])
    return key, hits[key] / float(total)


# ---------------------------------------------------------------- 自动识别表类型
# 导入时「人工一个个选」是纯体力活：文件名和表头其实已经把答案写在脸上了。
# 这里先猜，猜完前端把结果填进下拉框并写明依据，人只需要瞅一眼、改个别错的。
# 判据分两层：表头（强证据，权重高）→ 文件名（辅助，兜住表头不规范的导出）。
_TGT_HINTS = {
    "journal": {
        "label": "序时账",
        "head": ["凭证号", "凭证日期", "凭证字号", "凭证类型", "凭证编号", "记账凭证",
                 "摘要", "借方金额", "贷方金额", "分录", "日记账"],
        "file": ["序时", "凭证", "分录", "日记账", "明细账", "journal", "voucher"],
    },
    "tb": {
        "label": "科目余额表",
        "head": ["期初余额", "期末余额", "年初余额", "本期发生额", "本年累计",
                 "余额方向", "期初借方", "期末借方", "累计借方", "科目级次"],
        "file": ["余额", "试算", "科目余额", "balance", "trial"],
    },
    "flow": {
        "label": "银行流水",
        "head": ["交易日期", "交易时间", "对方账号", "对方户名", "对方账户",
                 "交易金额", "账户余额", "收支", "借贷标志", "流水号", "本方账号"],
        "file": ["流水", "对账", "银行交易", "交易明细", "statement", "flow"],
    },
    "fa": {
        "label": "固定资产台账",
        "head": ["原值", "累计折旧", "资产编码", "资产名称", "使用年限", "净值",
                 "固定资产", "购置日期", "折旧年限", "残值率"],
        "file": ["固定资产", "台账", "资产清单", "fa", "asset"],
    },
}


def detect_target(file_name="", sheet_name="", columns=()):
    """按表头 + 文件名猜这张表属于哪类标准表。

    返回 {target, label, score, reason, ranking}：
      target 为空串表示没把握（前端保持「未指定」，等人选）；
      reason 是给人看的一句依据，写在旁边方便复核——「自动识别」最怕的是
      猜错了还不说为什么，人只能一格一格去验。
    """
    cols = [str(c or "") for c in (columns or [])]
    # 列名先归一（去空格/全角/标点）再找关键词。人手敲的表头常夹着看不见的空格
    # （「分 类」「原 值」「累 计 折 旧」），按原文找关键词一个都命中不了，
    # 明明表头写着「原值」，却只能靠文件名才认出是固定资产台账。
    # 认列（auto_map）和找表头行（detect_header）本来就都用 _norm，这里补齐同一层，
    # 三处口径一致；只是比对时归一，列名本身原样保留，不改客户看到的名字。
    norms = [_norm(c) for c in cols]
    joined = " ".join(norms)
    fn = "%s %s" % (file_name or "", sheet_name or "")
    fn_l = _norm(fn)

    # 对照表 / 映射表 / 说明表本来就不是账，直接不猜。
    # 这类表常常带着一两个账表列名（瑞恒的「分类对照」就有「累计折旧」），
    # 再加文件名里的「固定资产」，分数会凑够门槛被误判成台账 —— 归一之后再比对，
    # 这种拼凑更容易过线，所以得显式挡一道。
    if any(w in fn_l for w in ("对照", "映射", "mapping")):
        return {"target": "", "label": "", "score": 0,
                "reason": "看起来是分类对照 / 映射表，不是账表，请人工指定",
                "ranking": []}

    scores = {}
    reasons = {}
    for tgt, hint in _TGT_HINTS.items():
        sc = 0
        why = []
        # 表头命中：每个词 2 分，最多算 4 个词，避免一列「摘要」把别的都压死
        hit = [w for w in hint["head"] if _norm(w) in joined][:4]
        if hit:
            sc += 2 * len(hit)
            why.append("表头含「%s」" % "、".join(hit[:3]))
        # 文件名命中：3 分（文件名是导出人自己起的，往往比表头直白）
        fhit = [w for w in hint["file"] if _norm(w) in fn_l]
        if fhit:
            sc += 3
            why.append("文件名含「%s」" % fhit[0])
        if sc:
            scores[tgt] = sc
            reasons[tgt] = "；".join(why)

    # 硬规则：表头里出现「凭证」二字，基本就是序时账（用户原话）
    if any("凭证" in c for c in norms):
        scores["journal"] = scores.get("journal", 0) + 5
        reasons["journal"] = (reasons.get("journal", "") + "；表头有「凭证」列").strip("；")

    if not scores:
        return {"target": "", "label": "", "score": 0, "reason": "文件名和表头都没有可识别的特征，请人工指定",
                "ranking": []}
    best = max(scores, key=lambda k: scores[k])
    # 分不够高、或与第二名拉不开差距，就别硬猜
    order = sorted(scores.items(), key=lambda kv: -kv[1])
    runner = order[1][1] if len(order) > 1 else 0
    if scores[best] < 4 or (runner and scores[best] - runner < 2):
        return {"target": "", "label": "",
                "score": scores[best],
                "reason": "特征不够明确（%s），请人工指定" % "、".join(
                    "%s%d分" % (_TGT_HINTS[k]["label"], v) for k, v in order),
                "ranking": [{"target": k, "label": _TGT_HINTS[k]["label"], "score": v}
                            for k, v in order]}
    # 固定资产：期初/期末是两份表，按名字里的「期初/年初」分
    target = best
    extra = ""
    if best == "fa":
        if any(w in fn_l for w in ("期初", "年初", "上年末")) \
                or any(w in joined for w in ("期初原值", "年初原值")):
            target = "fa_beg"
            extra = "（期初台账）"
    return {"target": target,
            "label": _TGT_HINTS[best]["label"] + extra,
            "score": scores[best],
            "reason": reasons.get(best, ""),
            "ranking": [{"target": k, "label": _TGT_HINTS[k]["label"], "score": v} for k, v in order]}
