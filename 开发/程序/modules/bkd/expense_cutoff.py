# -*- coding: utf-8 -*-
"""费用截止性测试（VC&VD.01.4）：样本池 + 遴选。

设计要点
--------
1. 全量序时账先进「样本池-cut off」，一行一笔，看得见每个人怎么被判的；
   再按「是否排除 = 否 + 落在测试期内 + 金额达阈值」挑进 VC&VD.01.4。
2. **结转损益必须先排掉**。费用科目月末结转（借本年利润 / 贷费用）会把当月
   净额整笔挂到贷方，金额往往比任何真实业务都大；不排的话样本池里全是结转，
   截止性测试就成了测结转。识别用「凭证内出现 4103/4104」+「摘要含结转损益」，
   不用裸「结转」二字 ——「结转----确认上市费用」是真实借方费用（V1.2 备忘的坑）。
3. 口径排除（薪酬/股权激励/折旧摊销）的判定顺序固定：股权激励 → 折旧 → 薪酬。
   顺序不能反：股份支付常写成「员工股权激励薪酬」，先判薪酬会归错。
4. 预提冲销按「同主体 + 同科目代码 + 同金额、一正一负」配对。配对只在**未被
   口径排除**的行之间做，否则「借方 1 万 / 结转贷方 1 万」会被当成冲销对，
   一下子吃掉大半样本池。金额下限 100 元，避免 0.04 元这种尾差被标成预提。
"""

import datetime
import re
from collections import defaultdict

POOL_SHEET = "样本池-cut off"
TARGET_SHEET = "VC&VD.01.4 截止性测试"

# 「是否排除」的取值。**一套分类同时服务 01.4 和 TOD**，颗粒度按 TOD 模板的
# 1.2 剔除项定（固定资产折旧费 / 使用权资产折旧费 / 无形资产摊销费 /
# 长期待摊费用摊销费 / 工资薪酬 / 租赁费 / 预提费用 / 负值），
# 这样 TOD 的 SUMIFS 直接按这一列匹配，不用再多一列「TOD 分类」。
KIND_XFER = "结转损益"
KIND_PAY = "工资薪酬"
# 股权激励/股份支付单列一项：它没有现金流出、底稿也不同（股份支付测算），
# 混在「工资薪酬」里，TOD 1.2 的剔除表就看不出这块被排除了多少。
KIND_EQUITY = "股权激励（股份支付）"
KIND_DEP_FIXED = "固定资产折旧费"
KIND_DEP_ROU = "使用权资产折旧费"
KIND_AMORT_IA = "无形资产摊销费"
KIND_AMORT_LT = "长期待摊费用摊销费"
KIND_RENT = "租赁费"
KIND_ACCRUAL = "预提费用（含预提和冲销）"
KIND_NEG = "负值"
KIND_NONE = "否"

# 折旧/摊销四类：TOD 模板分开列，01.4 只看「不是否」，两处共用
DEP_KINDS = (KIND_DEP_FIXED, KIND_DEP_ROU, KIND_AMORT_IA, KIND_AMORT_LT)
ALL_KINDS = (KIND_XFER, KIND_PAY, KIND_EQUITY) + DEP_KINDS + (KIND_RENT, KIND_ACCRUAL,
                                                             KIND_NEG, KIND_NONE)

# 顺义不能乱：股权激励在薪酬之前（股份支付常写「员工股权激励薪酬」）
EQUITY_WORDS = ("股权激励", "股份支付", "限制性股票", "期权", "员工持股", "限制性股份")
DEPR_WORDS = ("折旧", "摊销")
PAY_WORDS = ("薪酬", "工资", "奖金", "提成", "社会保险", "公积金", "福利",
             "职工教育", "工会", "离职补偿", "辞退", "劳务", "董事", "监事",
             "年金", "补贴", "津贴", "社保", "人力成本", "年终奖")

#: 科目名归不出类别时，从**摘要**再认一次的词表。
#: 客户常把摊销/股份支付记在「管理费用_装修费」「管理费用_其他」这类科目下，
#: 光看科目名永远认不出（瑞云 1-9 月：装修费 31 行摘要全是「摊销--N月--长期待摊费用-…」）。
#: 顺序即优先级：先认长期待摊（避免「长期待摊摊销」被当普通摊销），再认股份支付。
SUMMARY_HINTS = (
    ("长期待摊", KIND_AMORT_LT),
    ("装修费摊销", KIND_AMORT_LT),
    ("股份支付", KIND_EQUITY),
    ("股权激励", KIND_EQUITY),
    ("限制性股票", KIND_EQUITY),
    ("员工持股", KIND_EQUITY),
)

XFER_WORDS = ("结转本期损益", "结转损益", "期间损益结转")
PAIR_MIN = 100.0          # 低于这个金额不配「预提冲销」，避免尾差噪声
AMT_EPS = 0.005
GAP_COLOR = "FFFF00"      # 与 BKD 一致的"待人工补"黄

# TOD 池就是在截止性池上再加三列（分类已合到「是否排除」，不再单列）
TOD_HEADERS = ["是否关键项", "抽样ID", "TOD 说明"]


def _f(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _fiscal_start(balance_date, period_from="1"):
    """本期起点 = 资产负债表日所在年的「期间起」月 1 日。解析不出返回 None（不过滤）。"""
    try:
        y = int(str(balance_date)[:4])
    except (TypeError, ValueError):
        return None
    s = str(period_from or "1").strip()
    try:
        m = int(s.split("-")[-1]) if "-" in s else int(s)
    except ValueError:
        m = 1
    m = min(max(m, 1), 12)
    return datetime.date(y, m, 1)


def _core(name):
    """去掉「销售费用_ / 管理费用_」前缀，只看明细科目名。"""
    nm = str(name or "").strip()
    return nm.split("_", 1)[1] if "_" in nm else nm


def cutoff_kind(acct_name):
    """科目名 → 「是否排除」的细分类。归不进返回空串（后面会落「否」或「负值」）。

    颗粒度按 TOD 模板的 1.2 剔除项走，一套分类两边共用：
    股权激励并入「工资薪酬」（本质还是人工成本），折旧摊销按**资产类型**拆四类，
    租赁费单列。这样 TOD 的 SUMIFS 直接吃这一列，不用再维护第二套分类。
    """
    core = _core(acct_name)
    for w in EQUITY_WORDS:
        if w in core:
            return KIND_EQUITY
    for w in DEPR_WORDS:
        if w in core:
            if "使用权" in core:
                return KIND_DEP_ROU
            if "无形资产" in core:
                return KIND_AMORT_IA
            if "长期待摊" in core or "待摊" in core:
                return KIND_AMORT_LT
            return KIND_DEP_FIXED
    for w in PAY_WORDS:
        if w in core:
            return KIND_PAY
    # V1.2.80：**租赁费不再默认排除**。租金在费用里往往是单笔大额、且不必然
    # 在「使用权资产」底稿里测过（短期/低价值租赁直接进费用，没有使用权资产），
    # 一刀切排除会让样本池漏掉一整类高金额业务。归「否」，由抽样决定要不要查。
    return ""


def _is_xfer(voucher_key, summary, xfer_keys):
    if voucher_key in xfer_keys:
        return True
    s = str(summary or "")
    return any(w in s for w in XFER_WORDS)


def read_pool(conn, l1_list=("销售费用", "管理费用")):
    """取销售费用/管理费用的**借方**序时账（不受 BKD 期间限制）。

    金额方向（V1.2.82 起）：**只取借方行，贷方行不进池子**。

    为什么：费用科目的贷方几乎全是月末结转（借本年利润 / 贷费用），借贷两侧一起
    收进来，整列加总必然抵成 0 —— 瑞云 1-9 月实测借方 38,362,151.70 与贷方
    38,362,151.70 完全相等，池子合计 0.00。样本池合计、按金额排序、按阈值挑关键项
    全部建立在这个 0 上，等于没有总体。费用发生的正常方向就是借方，借方发生额
    才是 BKD「本期账面数」的口径。
    负数借方（红字冲销）仍是借方行，照旧保留 —— 负值分析不能因此瞎掉。
    """
    try:
        from . import expense as _EXP0
        cm = _EXP0.load_company_map(conn)
    except Exception:
        cm = {}
    ph = ",".join("'%s'" % x for x in l1_list)
    xfer_keys = set()
    try:
        for r in conn.execute(
                "SELECT DISTINCT voucher_key FROM journal "
                "WHERE acct_code LIKE '4103%' OR acct_code LIKE '4104%'").fetchall():
            if r[0]:
                xfer_keys.add(r[0])
    except Exception:
        xfer_keys = set()

    sql = """
    SELECT company_name, acct_code, acct_name, acct_l1, vdate, period, vno, voucher_key,
           summary, currency,
           debit                                  AS amt,
           debit_fx                               AS amt_fx,
           debit, credit, debit_fx, credit_fx
    FROM journal
    WHERE acct_l1 IN (%s) AND coalesce(debit, 0) <> 0
    ORDER BY vdate, company_name, acct_code, voucher_key
    """ % ph
    out = []
    for r in conn.execute(sql).fetchall():
        (company, code, name, l1, vdate, period, vno, vkey,
         summary, currency, amt, amt_fx, debit, credit, dfx, cfx) = r
        # 折算率：本币 / 原币。只从借方取（贷方行已不在池子里），拿不到就按 1。
        # **必须存全精度**：存成 6 位小数会让「原币 × 折算率」回不到本币，
        # 十万行累计能差出一块多，池子合计与 BKD 就对不上了。显示格式在写表时设。
        rate_base = 1.0
        if dfx is not None and abs(float(dfx)) > 1e-9 and debit is not None:
            rate_base = abs(float(debit) / float(dfx))
        # 折成人民币这一步必须跟 BKD 同源：客户「公司mapping」按公司给的汇率
        # （费用属于损益类 → 损益汇率）。池子「原币 × 折算率」要等于 BKD
        # 「本年账面数（人民币）」，否则 1.1 总体与 BKD 合计对不上。
        try:
            from . import expense as _EXP
            _cmr = _EXP.rate_of(cm, company, name) or _EXP.fx_rate_of(company)
        except Exception:
            _cmr = 1.0
        rate = rate_base * _cmr
        a = _f(amt)
        afx = _f(amt_fx) if dfx is not None else a
        out.append({
            "company": company or "", "code": str(code or ""), "name": name or "",
            "l1": l1 or "", "vdate": vdate, "period": period or "",
            "vno": vno or "", "vkey": vkey or "", "summary": summary or "",
            "currency": currency or "", "amt": a, "amt_fx": afx,
            "rate": rate if rate else 1.0,
            "rate_base": rate_base if rate_base else 1.0,
            "xfer": _is_xfer(vkey, summary, xfer_keys),
            "excl": "", "note": "",
        })
    return out


#: OAR 口径里属于「人工成本」的词。命中即归「工资薪酬」桶。
#: 池子分桶与 BKD 的 OAR 口径必须是同一批科目 —— 否则 check 的差额既可能是
#: 分类差、也可能是金额差，说不清就不能用。
PAY_HINTS = ("薪酬", "工资", "人力", "职工", "奖金", "提成", "社会保险", "社保",
             "公积金", "福利", "津贴", "补贴", "劳务", "董事", "监事", "年金",
             "股权激励", "股份支付", "年终奖")


def is_pay_bucket(bucket):
    """这个 OAR 口径算不算「工资薪酬」。口径名空值一律不算。"""
    b = str(bucket or "").strip()
    return bool(b) and any(w in b for w in PAY_HINTS)


def mark_exclusions(rows, oar_map=None):
    """给每行打「是否排除」。

    顺序（不能乱）：
        结转损益 → 科目性质 → 预提冲销 → 负值 → 否

    **负值排在最后**：先按科目性质把「已在其他底稿测过」的项目挑走，剩下那些
    说不出性质的负数才归「负值」。反过来的话，一笔「折旧 -5 万」会先被负值
    吃掉，a 段的固定资产折旧费就少一块，TOD 的 1.2 剔除表永远对不上。

    V1.2.82：**工资薪酬改听 OAR 口径**。oar_map 是 {(主体, 科目名): 口径}，
    口径属人工成本就进这个桶；OAR 明确给了别的口径，关键词也不许把它判成薪酬
    （「劳务费」「董事津贴」这类名字很容易误伤）。查不到口径的行才退回关键词。
    其余四个折旧摊销桶仍然按科目名关键词 —— 这次的改动只针对工资。
    """
    # 1) 结转损益
    for r in rows:
        if r["xfer"]:
            r["excl"] = KIND_XFER
            r["note"] = "期间损益结转凭证，不是费用发生"

    # 2) 科目性质：工资看 OAR 口径，其余看关键词
    for r in rows:
        b = (oar_map or {}).get((r["company"], r["name"])) or (oar_map or {}).get(
            ("", r["name"]))
        # 口径挂在行上，TOD 的 1.2 核对要用它反过来推「工资类口径有哪些」
        r["oarBucket"] = b or ""
        if r["excl"]:
            continue
        if b and is_pay_bucket(b):
            # 口径里明确是股权激励/股份支付的，单列一项；其余人工成本仍归工资薪酬
            if any(w in str(b) for w in EQUITY_WORDS):
                r["excl"] = KIND_EQUITY
                r["note"] = "股权激励（股份支付）：OAR 口径为「%s」" % b
            else:
                r["excl"] = KIND_PAY
                r["note"] = "工资薪酬：OAR 口径为「%s」" % b
            continue
        k = cutoff_kind(r["name"])
        # OAR 已明确不是人工成本的口径，关键词不得再判成薪酬（含股权激励）
        if k in (KIND_PAY, KIND_EQUITY) and b and not is_pay_bucket(b):
            k = ""
        _src = _core(r["name"])
        if not k:
            # 科目名认不出时看摘要：长期待摊费用摊销、股份支付常常藏在
            # 「装修费」「其他」这类科目下，只有摘要写得清楚。
            _sm = str(r.get("summary") or "")
            for _w, _kk in SUMMARY_HINTS:
                if _w in _sm:
                    # 股份支付若 OAR 已给出非人工成本口径，同样不硬判成薪酬
                    if _kk in (KIND_PAY, KIND_EQUITY) and b and not is_pay_bucket(b):
                        continue
                    k = _kk
                    _src = "摘要「%s」" % _w
                    break
        if k:
            r["excl"] = k
            r["note"] = "%s：已在其他底稿执行程序（%s）" % (k, _src)

    # 3) 预提冲销：同主体 + 同科目代码 + 同金额、一正一负（只在剩余行里配）
    groups = defaultdict(list)
    for i, r in enumerate(rows):
        if r["excl"]:
            continue
        if abs(r["amt"]) < PAIR_MIN:
            continue
        groups[(r["company"], r["code"], round(abs(r["amt"]), 2))].append(i)
    pairs = 0
    for _k, idxs in groups.items():
        pos = [i for i in idxs if rows[i]["amt"] > 0]
        neg = [i for i in idxs if rows[i]["amt"] < 0]
        for a, b in zip(pos, neg):
            for i, j in ((a, b), (b, a)):
                rows[i]["excl"] = KIND_ACCRUAL
                rows[i]["note"] = "%s %s %s 与 %s 的 %s 一正一负" % (
                    "计提" if rows[i]["amt"] > 0 else "冲回",
                    rows[j]["vdate"], ("%+.2f" % rows[j]["amt"]),
                    rows[j]["vdate"], r"同科目同金额")
            pairs += 1

    # 4) 负值：性质都归完之后，剩下的负数才进这个桶
    neg_cnt = 0
    for r in rows:
        if r["excl"]:
            continue
        if r["amt"] < 0:
            r["excl"] = KIND_NEG
            r["note"] = "无对应科目性质，作为负值单独分析（见表2）"
            neg_cnt += 1

    # 5) 其余「否」
    for r in rows:
        if not r["excl"]:
            r["excl"] = KIND_NONE
    return {"pairs": pairs, "negatives": neg_cnt}


# ---------------------------------------------------------------- 写样本池
def _pool_headers():
    return ["序号", "主体", "科目代码", "科目名称", "归属", "日期", "期间", "凭证号",
            "摘要", "币种", "原币金额", "折算率", "折合人民币", "是否排除", "排除说明"]


def write_pool(wb, rows, balance_date="", sheet=None, tod=False):
    """写样本池。tod=True 时多出 4 列（TOD 分类 / 关键项 / 抽样ID / 说明）。

    同一份 rows 写两遍：一遍给截止性（15 列），一遍给 TOD（19 列）。
    这样两份池子的数据必然一致，改口径不会一边变一边不变。
    """
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    sheet = sheet or POOL_SHEET
    if sheet in wb.sheetnames:
        del wb[sheet]
    # 紧挨着 01.4 放，别沉到 Skywind 那些内部表后面
    if tod:
        try:
            pos = wb.sheetnames.index(POOL_SHEET) + 1
        except ValueError:
            pos = len(wb.sheetnames)
    else:
        try:
            pos = wb.sheetnames.index(TARGET_SHEET)
        except ValueError:
            pos = len(wb.sheetnames)
    ws = wb.create_sheet(sheet, pos)

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    white = Font(name="宋体", size=10)
    head_f = Font(name="宋体", size=10, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4472C4")
    title_f = Font(name="宋体", size=14, bold=True)
    sub_f = Font(name="宋体", size=9, color="666666")
    # 颜色按「大类」给，不按细类：工资薪酬和四类折旧摊销都是灰（已在其他底稿测过），
    # 预提费用黄（要看配对），负值橙（要看方向），结转损益蓝（不是费用发生）
    grey = PatternFill("solid", fgColor="EDEDED")
    fills = {KIND_NONE: None}
    for _k in (KIND_PAY,) + DEP_KINDS + (KIND_RENT,):
        fills[_k] = grey
    fills[KIND_ACCRUAL] = PatternFill("solid", fgColor="FFF2CC")
    fills[KIND_NEG] = PatternFill("solid", fgColor="FCE4D6")
    fills[KIND_XFER] = PatternFill("solid", fgColor="DDEBF7")
    ctr = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=False)

    headers = _pool_headers() + (TOD_HEADERS if tod else [])
    ncol = len(headers)
    widths = [6, 28, 12, 30, 10, 11, 9, 22, 52, 8, 15, 8, 15, 10, 46]
    if tod:
        widths += [10, 10, 40]
    for i, w in enumerate(widths[:ncol], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.cell(1, 1, ("样本池-%s　销售费用 + 管理费用（仅借方凭证）"
                   % ("TOD　TOD 年审用（与截止性池同一份底料）" if tod else "cut off"))).font = title_f
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ws.cell(2, 1, "资产负债表日 %s；共 %d 行，只含借方凭证（贷方是结转/冲回，不进池子）。"
                  "是否排除：结转损益 / 工资薪酬 / 折旧摊销 / 预提冲销，其余为「否」。"
                  "折合人民币 = 原币金额 × 折算率（公式，可核对），关键项阈值、1.1 总体、各路核对全部以它为准；"
                  "原币金额仅供查看，不参与任何判定。"
                  "折算率列黄底＝台账没有本币数，需人工填汇率后重算。"
                  "%s"
            % (balance_date or "（未填）", len(rows),
               "TOD 分类对齐 VC&VD.01.2 的 1.2 剔除项：负数一律进「负值」（含预提冲销的负方），"
               "正数按性质分桶；工资薪酬按 OAR 口径判定，其余桶按科目名。抽样ID 覆盖「是否排除=否」"
               "的全部本期样本，按台账顺序 TOD-0001 起连续编号（日后凭证属截止性测试范围，不编号）。"
               if tod else "")).font = sub_f
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)

    for i, h in enumerate(headers, start=1):
        c = ws.cell(3, i, h)
        c.font = head_f
        c.fill = head_fill
        c.border = border
        c.alignment = ctr

    # 「折算率」列写成 XLOOKUP：客户改「公司mapping」里的汇率，池子跟着变。
    try:
        from . import expense as _EXPw
        _cmref = _EXPw.find_company_map_ref(wb)
    except Exception:
        _cmref = None

    r = 4
    fx_need = 0
    for n, row in enumerate(rows, start=1):
        vd = row["vdate"]
        if isinstance(vd, datetime.datetime):
            vd = vd.date()
        vals = [n, row["company"], row["code"], row["name"], row["l1"],
                vd, row["period"], row["vkey"], row["summary"],
                row["currency"], round(row["amt_fx"], 2), row["rate"],
                None, row["excl"], row["note"]]
        if tod:
            # 分类已合进「是否排除」，不再单列一列 TOD 分类
            vals += [row.get("todKey") or "", row.get("sid") or "",
                     row.get("todNote") or ""]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(r, i, v)
            c.font = white
            c.border = border
            if i in (5, 6, 7, 10, 14, 16, 17):
                c.alignment = ctr
                if i == 6:
                    c.number_format = "yyyy-mm-dd"
            elif i == 11 or i == 13:
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = r'_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '
            elif i == 12:
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "General"   # 折算率原样显示，不四舍五入
            else:
                c.alignment = left
        # 折算率 = 本币/原币 × 客户「公司mapping」的损益汇率（XLOOKUP，不写死）
        if _cmref:
            ws.cell(r, 12).value = (
                "=%r*IFERROR(XLOOKUP($B%d,'%s'!$%s$%d:$%s$%d,'%s'!$%s$%d:$%s$%d),1)"
                % (float(row.get("rate_base") or 1.0), r, _cmref["sheet"],
                   _cmref["full"], _cmref["r1"], _cmref["full"], _cmref["r2"],
                   _cmref["sheet"], _cmref["pl"], _cmref["r1"], _cmref["pl"], _cmref["r2"]))
            ws.cell(r, 12).number_format = "General"
        # 折合人民币：写成公式，过程可追溯
        # 不逐行取整：与 BKD「本年账面数（人民币）= 本币×汇率」保持同一算法，
        # 否则一边按凭证取整、一边按科目取整，累计差几分钱会把 TOD 1.1 的 check 弄脏。
        ws.cell(r, 13).value = '=IF(K%d="","",K%d*L%d)' % (r, r, r)
        ws.cell(r, 13).number_format = r'_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '
        f = fills.get(row["excl"])
        if f is not None:
            ws.cell(r, 14).fill = f
        # 真正缺汇率的情况：给了外币金额、但折算率还是 1（说明台账里没有本币数），
        # 此时「折合人民币」其实等于原币，按阈值判会判错 —— 标黄请人工填汇率后重算。
        # 币种串是「人民币(CNY)」这种，不能用 `!= "人民币"` 判，那样十万行人民币全标黄。
        if abs(row["rate"] - 1.0) < 1e-9 and abs(row["amt_fx"] - row["amt"]) > 0.005:
            ws.cell(r, 12).fill = PatternFill("solid", fgColor=GAP_COLOR)
            fx_need += 1
        r += 1

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = "A3:%s%d" % (get_column_letter(ncol), r - 1)
    return {"rows": len(rows), "sheet": sheet, "fxNeedRate": fx_need}


# ---------------------------------------------------------------- 写 01.4
def _insert_rows(ws, at, n):
    """插行并搬走下方合并区与行高（openpyxl 的 insert_rows 两样都不管）。"""
    if n <= 0:
        return 0
    merges, heights = [], []
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row >= at:
            merges.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
            ws.unmerge_cells(str(mc))
    for rr in list(ws.row_dimensions.keys()):
        if rr >= at:
            h = ws.row_dimensions[rr].height
            if h:
                heights.append((rr, h))
    ws.insert_rows(at, n)
    for (r1, r2, c1, c2) in merges:
        ws.merge_cells(start_row=r1 + n, start_column=c1,
                       end_row=r2 + n, end_column=c2)
    for (rr, h) in heights:
        try:
            ws.row_dimensions[rr + n].height = h
        except Exception:
            pass
    return len(merges)


def _copy_style(ws, src_row, dst_row, cols=range(2, 12)):
    from copy import copy
    for c in cols:
        try:
            ws.cell(dst_row, c)._style = copy(ws.cell(src_row, c)._style)
        except Exception:
            pass
    try:
        if ws.row_dimensions[src_row].height:
            ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    except Exception:
        pass


def fill_target(wb, rows, balance_date, threshold, days=30):
    """把「否 + 测试期内 + 金额达阈值」的样本写进 VC&VD.01.4 表3。"""
    from datetime import date, timedelta
    from openpyxl.styles import Alignment

    if TARGET_SHEET not in wb.sheetnames:
        return {"status": "skip", "reason": "模板里没有「%s」" % TARGET_SHEET}

    ws = wb[TARGET_SHEET]
    bal = None
    if balance_date:
        try:
            bal = datetime.date.fromisoformat(str(balance_date)[:10])
        except ValueError:
            bal = None
    if bal is None:
        return {"status": "skip", "reason": "没有资产负债表日，无法确定测试期间"}

    d1, d2 = bal - timedelta(days=days), bal + timedelta(days=days)
    ws["B10"] = days                       # 表1：前后天数（D10/E10 是公式）

    sel = []
    for r in rows:
        if r["excl"] != "否":
            continue
        v = r["vdate"]
        if not isinstance(v, (datetime.date, datetime.datetime)):
            continue
        v = v.date() if isinstance(v, datetime.datetime) else v
        if not (d1 <= v <= d2):
            continue
        if abs(r["amt"]) < threshold:
            continue
        sel.append(r)

    before = [r for r in sel if r["vdate"] <= bal]
    after = [r for r in sel if r["vdate"] > bal]

    # 模板：本期 r25-29（5 行），日后标题 r30，日后 r31-35（5 行）
    R_BEFORE, HAVE_BEFORE, R_AFTER, HAVE_AFTER = 25, 5, 31, 5
    if len(before) > HAVE_BEFORE:
        _insert_rows(ws, R_AFTER - 1, len(before) - HAVE_BEFORE)
        R_AFTER += len(before) - HAVE_BEFORE
    if len(after) > HAVE_AFTER:
        _insert_rows(ws, R_AFTER + HAVE_AFTER, len(after) - HAVE_AFTER)

    def put(items, start, seg):
        for i, it in enumerate(items):
            r = start + i
            if i >= 5:
                _copy_style(ws, start, r)
            ws.cell(r, 2, it["vkey"] or it["vno"])          # 凭证编号
            vd = it["vdate"]
            vd = vd.date() if isinstance(vd, datetime.datetime) else vd
            c = ws.cell(r, 3, vd)
            c.number_format = "yyyy-mm-dd"
            c = ws.cell(r, 4, round(it["amt"], 2))
            c.number_format = r'_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '
            ws.cell(r, 5, it["name"])                        # 明细科目名称
            ws.cell(r, 6, "")                                # 交易凭据名称（人工）
            ws.cell(r, 7, "")                                # 交易凭据编号（人工）
            ws.cell(r, 8, "")                                # 交易发生日期（人工）
            for cidx in (2, 3, 4, 5, 6, 7, 8):
                ws.cell(r, cidx).alignment = Alignment(
                    horizontal="center" if cidx in (2, 3, 8) else "left",
                    vertical="center")
            if seg == "before":
                ws.cell(r, 9, "√")
                ws.cell(r, 9).alignment = Alignment(horizontal="center")
            else:
                ws.cell(r, 10, "√")
                ws.cell(r, 10).alignment = Alignment(horizontal="center")
            ws.cell(r, 11, "")
            ws.cell(r, 12, it["summary"][:80])

    put(before, R_BEFORE, "before")
    put(after, R_AFTER, "after")

    # ---- 表2：样本策略说明（数字取**实际写进表3**的样本数）----
    # 模板这块是空的（B13:E15 合并框），项目组每次都要手写一遍；
    # 数字直接引用本次实际抽出的条数，避免"策略里写 5 个、表3 有 7 行"。
    try:
        txt = ("经分析，剔除薪酬、折旧等已在其他底稿测试的费用明细后，我们基于定量和定性两个维度确定关键样本。"
               "定量方面，我们选择完整性和发生认定的测试阈值中的较低值作为关键样本的定量标准，"
               "定性方面，我们选择性质特殊的项目作为关键样本的定性标准。\n"
               "按照如上标准，我们在资产负债表日前识别到 %d 个测试样本，"
               "在资产负债表日后识别到 %d 个测试样本。" % (len(before), len(after)))
        c = ws.cell(13, 2, txt)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        # 合并区 B13:E15 只有 3 行高，长文本会溢出被盖住 —— 把这三行撑开
        for rr in (13, 14, 15):
            ws.row_dimensions[rr].height = 30
    except Exception:
        pass

    return {
        "status": "ok", "sheet": TARGET_SHEET,
        "days": days, "from": str(d1), "to": str(d2),
        "threshold": round(threshold, 2),
        "before": len(before), "after": len(after), "total": len(sel),
        "beforeRows": [R_BEFORE, R_BEFORE + max(0, len(before) - 1)],
        "afterRows": [R_AFTER, R_AFTER + max(0, len(after) - 1)],
        "strategyFilled": True,
    }


def drop_before_period(rows, balance_date, period_from="1"):
    """掐掉「本期起点」之前的行，返回 (保留的行, 掐掉的条数)。

    账套里常带上一年 12 月的计提+结转：结转那一半会被判成「结转损益」剔掉，
    计提那一半留在池子里，TOD 的 1.1 总体就会比 BKD 合计多出一块。
    瑞云这笔正好 80,378.56（销售 25,477.87 / 管理 54,899.69）。
    """
    d_from = _fiscal_start(balance_date, period_from)
    if not d_from:
        return rows, 0
    keep, dropped = [], 0
    for r in rows:
        v = r.get("vdate")
        if isinstance(v, datetime.datetime):
            v = v.date()
        if isinstance(v, datetime.date) and v < d_from:
            dropped += 1
            continue
        keep.append(r)
    return keep, dropped


def build(wb, conn, balance_date="", threshold=275000.0, days=30,
          stats=None, tpl_dir=None, pre_end=None, period_from="1",
          prior_tod_bytes=None, prior_tod_sheet=None, oar_map=None):
    """总入口：建样本池 + 填 01.4 + TOD 年审与预审+剩余期间（同一份 rows 多个用途）。

    stats 是 BKD 各页的写入统计（含 sum_row），TOD 的 1.1 总体要链到 BKD 合计行。
    pre_end 是预审段结束月份；给了它，另一版 TOD 页才拆得出「剩余期间」。
    period_from 用来掐掉**上一年度**的凭证（口径见 drop_before_period）。
    oar_map 是 {(主体, 科目名): OAR 口径}，另加 ("", 科目名) 同名兜底键；
    工资薪酬桶按它判定，与 BKD 的 OAR 口径同一批科目。
    """
    try:
        rows = read_pool(conn)
    except Exception as e:
        return {"status": "error", "reason": "%s: %s" % (type(e).__name__, str(e)[:160])}
    if not rows:
        return {"status": "skip", "reason": "序时账里没有销售费用/管理费用行"}

    rows, dropped = drop_before_period(rows, balance_date, period_from)
    if not rows:
        return {"status": "skip", "reason": "资产负债表日/期间设置下没有本期费用行"}

    pinfo = mark_exclusions(rows, oar_map)
    pool = write_pool(wb, rows, balance_date)
    try:
        tinfo = fill_target(wb, rows, balance_date, threshold, days)
    except Exception as e:
        tinfo = {"status": "error", "reason": "%s: %s" % (type(e).__name__, str(e)[:160])}

    # ---- TOD 年审：先打 TOD 标记（池子要用），再复制模板页填数，最后写池 ----
    tod_info = {}
    try:
        from . import expense_tod as ETOD
        tod_info = ETOD.build(wb, rows, stats=stats, balance_date=balance_date,
                              threshold=threshold, tpl_dir=tpl_dir, pre_end=pre_end,
                              prior_tod_bytes=prior_tod_bytes,
                              prior_tod_sheet=prior_tod_sheet,
                              period_from=period_from)
        if tod_info.get("status") == "ok":
            tod_info["pool"] = write_pool(wb, rows, balance_date,
                                          sheet=ETOD.TOD_POOL, tod=True)
    except Exception as e:
        tod_info = {"status": "error",
                    "reason": "%s: %s" % (type(e).__name__, str(e)[:160])}

    from collections import Counter
    cnt = Counter(r["excl"] for r in rows)
    out = {
        "status": "ok" if tinfo.get("status") == "ok" else tinfo.get("status", "partial"),
        "pool": pool, "poolRows": len(rows),
        # 按「是否排除」的细分类统计（与 TOD 1.2 剔除项一一对应）
        "excluded": {k: cnt.get(k, 0) for k in ALL_KINDS},
        "pairCount": pinfo.get("pairs", 0),
        # 上一年度的计提行被掐掉的条数（不掐会让 TOD 总体比 BKD 合计多出一块）
        "beforePeriod": dropped,
        "target": tinfo,
        "tod": tod_info,
    }
    if pool.get("fxNeedRate"):
        out["warning"] = ("有 %d 行是外币主体（记账本位币不是人民币），折算率按 1 计，"
                          "「折合人民币」实际是外币金额。样本池折算率列已标黄，"
                          "填好期末汇率后按原币金额×汇率重算，再判断是否达阈值。"
                          % pool["fxNeedRate"])
    return out
