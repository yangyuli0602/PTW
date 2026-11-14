# -*- coding: utf-8 -*-
"""口径 Mapping 表：OAR 口径 + 公司简称，两张表放同一个 sheet 的不同列。

为什么要它
----------
以前 BKD 的「OAR 口径」列是 `oar_bucket()` 当场算出来的定死文本：同一个科目
在不同年份、不同客户之间口径不连续，审计问「去年这个科目算在哪个口径」时
答不上来，也没法人工干预。

改成：先由**上年 BKD 的主体+科目**建一份 mapping，本年 BKD 用 XLOOKUP 去匹配。
匹配不上的（新主体、新科目）交给**平台 AI 做语义匹配**：导出待判断清单 →
AI 按该客户已有口径归类并给出依据与置信度 → 回填 oar_map_ai.json。
口径一旦确认就沉淀下来，明年同一客户直接命中（V1.2.80 起不再出现「未匹配-待判断」）。

布局（同一 sheet，两段互不干扰）
---------------------------------
A~F：OAR 口径 mapping
    A 主体 / B 科目名称 / C OAR 口径 / D 匹配键(主体|科目) / E 来源 / F 备注
H~J：公司 mapping
    H 主体全称 / I 简称 / J 备注

BKD 里 A 列（简称）按 H 列查、R 列（OAR 口径）按 D 列查，互不影响。
"""

import io
import re

SHEET = "口径Mapping"
HDR_OAR = 5          # OAR 段的表头行
HDR_COMPANY = 5      # 公司段的表头行
DATA_ROW = 6         # 数据起始行

# 「来源」列的三种取值。一眼能看出这一行是谁定的、要不要复核：
#   由AI判断         —— 平台 AI 做过语义匹配，备注里有判断依据 + 置信度
#   规则推导-待AI复核 —— 只按科目名关键词推的，没经过 AI，需要人看一眼
#   上年底稿          —— 有上期依据（上年口径Mapping 页 / 上年 BKD 口径列 / 上年同名科目），不需复核
SRC_PRIOR = "上年底稿"
SRC_AI = "由AI判断"
SRC_RULE = "规则推导-待AI复核"

# 上年 BKD 里口径列的表头候选（不同年份/模板写法不一）
#   客户自己的 WP 里这一列常写「for OAR」（小写 f），不能只按大写匹配 ——
#   V1.2.87 之前就是这么漏掉客户的 for OAR 列、退化到规则推导的。
OAR_HDR_WORDS = ("OAR 口径", "OAR口径", "FOR OAR", "FOR OAR 口径", "口径")


def is_oar_header(text):
    """这个表头是不是「OAR 口径」列（大小写、空格都不敏感）。"""
    u = _T(text).upper().replace(" ", "")
    if not u:
        return False
    return u in ("OAR口径", "FOROAR", "FOROAR口径", "口径") or u.startswith("FOROAR")


SHORT_HDR_WORDS = ("简称", "公司简称")

# 简称自动提取：先摘掉地域，再摘掉公司性质后缀，剩下字号
REGION_WORDS = ("深圳市", "深圳", "上海市", "上海", "北京市", "北京", "广州市", "广州",
                "重庆市", "重庆", "长沙市", "长沙", "杭州市", "杭州", "成都市", "成都",
                "西安市", "西安", "盐城市", "盐城", "贵安新区", "江西", "江苏省", "江苏",
                "无锡市", "无锡", "湖北省", "湖北", "东莞市", "东莞", "苏州市", "苏州",
                "香港", "新加坡")
TAIL_WORDS = ("股份有限责任公司", "股份有限公司", "有限责任公司", "集团有限公司",
              "科技有限公司", "有限公司", "集团", "公司")


def _T(v):
    if v is None:
        return ""
    return str(v).replace("\u00a0", " ").strip()


def guess_short(full):
    """从主体全称猜简称。只做**默认值**，人可以改，改完明年照用。

    「深圳市瑞云科技股份有限公司」→ 瑞云
    「深圳市瑞影科技有限公司广州分公司」→ 瑞影（先剥分公司再剥地域/后缀）
    「【瑞云】帐簿(001)」→ 瑞云
    「RAYWING PTE.LTD.」→ RAYWING
    """
    s = _T(full)
    if not s:
        return ""
    m = re.search(r"[【\[]([^】\]]+)[】\]]", s)      # 【瑞云】帐簿
    if m and _T(m.group(1)):
        return _T(m.group(1))[:12]
    s = re.sub(r"[（(][^）)]*[）)]", "", s)           # 瑞雲科技（香港）
    # 英文公司后缀先去掉，否则 "RAYWING PTE.LTD." 会截成 "RAYWING PTE."
    s = re.sub(r"\b(PTE\.?|LTD\.?|LIMITED|INC\.?|CO\.?|CORP\.?|B\.?V\.?|GMBH|LLC)\b",
               " ", s, flags=re.I)
    # 分公司：先剥掉「XX分公司」，露出的才是母体字号
    s = re.sub(r"(广州|北京|上海|深圳|杭州|成都|西安|苏州|东莞|武汉|南京)?分公司$", "", s)
    for w in REGION_WORDS:
        if s.startswith(w):
            s = s[len(w):]
            break
    for w in ("集团股份有限公司", "股份有限公司", "有限责任公司", "科技有限公司",
              "有限公司", "集团", "公司"):
        if s.endswith(w):
            s = s[:-len(w)]
            break
    s = s.replace("科技", "").replace("创新", "").strip(" 　-_.|")
    return (s or _T(full))[:10]


def _norm_key(company, name):
    return ("%s|%s" % (_T(company), _T(name))).lower()


# ---------------------------------------------------------------- 读上年
def read_prior(prior_bytes):
    """上年费用底稿 → ({(主体,科目): (口径, 来源)}, {主体: 简称}, 说明)。

    两级读，缺一不可（V1.2.85）：

    ① **上年自己那张「口径Mapping」页** —— 上年已经定论过的口径。主体已经是公司
       全称、口径是静态文本（不是公式），是全套数据里最干净的一层；
    ② **上年 BKD 的「OAR 口径」列** —— 键是账簿名（要 book_map 折算），值可能是
       XLOOKUP 公式（文件没被 Excel 保存过就读成空），所以只能兜底。

    两层的来源统一写「上年底稿」：复核的人只关心"这条口径去年就有依据"，
    不关心它来自哪张表。上年 mapping 里自己都还标着「规则推导-待AI复核」的行
    **跳过不读** —— 那不是依据，是去年的猜测，沿用它等于把猜测传下去。

    两层都取不到的，按 oar_bucket() 推一份并标「规则推导-待AI复核」（黄底，
    并入待 AI 清单）。来源不写清楚，明年没人知道这条口径从哪来。
    """
    from openpyxl import load_workbook
    out, shorts, notes = {}, {}, []
    if not prior_bytes:
        return out, shorts, "未提供上年费用底稿，mapping 只能按本年科目新建"
    try:
        wb = load_workbook(io.BytesIO(prior_bytes), data_only=True)
    except Exception as e:
        return out, shorts, "上年费用底稿打不开：%s" % e

    from .expense import oar_bucket

    try:
        n_map = _read_prior_mapping(wb, out, notes)
        raw, n_bkd = _read_prior_bkd(wb, shorts, notes)
    finally:
        try:
            wb.close()
        except Exception:
            pass

    # BKD 层只补 mapping 层没覆盖的键：上年定论过的口径优先，不被 BKD 的公式列盖掉
    tot = {}
    for (_bk, name), (_b, amt) in raw.items():
        tot[name] = tot.get(name, 0.0) + amt
    n_fill = n_rule = 0
    for (book, name), (bucket, amt) in raw.items():
        if (book, name) in out:
            continue
        if bucket:
            out[(book, name)] = (bucket, SRC_PRIOR)
            n_fill += 1
        else:
            # 按**科目合计**推：oar_bucket 的大额单列是按科目总额判断的，
            # 逐行判断会让「中介费」这种拆成 8 个账簿的科目全部落进「其他」。
            out[(book, name)] = (oar_bucket(name, tot.get(name, amt)), SRC_RULE)
            n_rule += 1
    if n_map:
        notes.append("上年「口径Mapping」页沿用 %d 行" % n_map)
    if n_fill:
        notes.append("上年 BKD 口径列补 %d 行" % n_fill)
    if n_rule:
        notes.append("上年 BKD 无口径（或口径列无缓存值）按规则推 %d 行" % n_rule)
    if not n_bkd and not n_map:
        notes.append("上年底稿里没认出 BKD 明细页或口径Mapping 页（表头要有「科目名称」）")
    return out, shorts, "；".join(notes)


def _read_prior_mapping(wb, out, notes):
    """读上年底稿里的「口径Mapping」页，写进 out。

    识别靠表头内容而不是页名：一行里同时有「科目名称」「OAR 口径」，
    且**没有**「账套/账簿」—— BKD 页也有前两者，靠「账套/账簿」把两者分开。
    """
    n = 0
    for ws in wb.worksheets:
        max_c = min(ws.max_column or 0, 20)
        hdr = None
        for r in range(1, min(ws.max_row or 0, 30) + 1):
            texts = [_T(ws.cell(r, c).value) for c in range(1, max_c + 1)]
            if not texts:
                continue
            # 必须**整格等于**表头词：说明行（第 2 行那种一句话图例）里也含「科目名称」
            # 「OAR 口径」，用包含匹配会把说明行当表头，整页被跳过。
            has_name = any(t == "科目名称" or t.startswith("科目名称") for t in texts)
            has_oar = any(is_oar_header(t) for t in texts)
            has_comp = any(t in ("主体", "主体全称", "公司名称", "公司全称", "公司")
                           for t in texts)
            has_book = any(("账套" in t or "账簿" in t) for t in texts)
            if has_name and has_oar and has_comp and not has_book:
                hdr = r
                break
        if not hdr:
            continue
        cols = {}
        for c in range(1, max_c + 1):
            t = _T(ws.cell(hdr, c).value)
            if not t:
                continue
            if t.startswith("科目名称") and "name" not in cols:
                cols["name"] = c
            elif t in ("主体", "主体全称", "公司名称", "公司全称", "公司") and "comp" not in cols:
                cols["comp"] = c
            elif is_oar_header(t) and "oar" not in cols:
                cols["oar"] = c
            elif "来源" in t and "src" not in cols:
                cols["src"] = c
            elif "备注" in t and "note" not in cols:
                cols["note"] = c
        if not all(k in cols for k in ("name", "comp", "oar")):
            continue
        for r in range(hdr + 1, min(ws.max_row or 0, hdr + 5000) + 1):
            comp = _T(ws.cell(r, cols["comp"]).value)
            name = _T(ws.cell(r, cols["name"]).value)
            bucket = _T(ws.cell(r, cols["oar"]).value)
            if not comp or not name or not bucket:
                continue
            if name in ("合计", "总计", "小计"):
                continue
            if cols.get("src"):
                src = _T(ws.cell(r, cols["src"]).value)
                if src == SRC_RULE:
                    # 上年自己都没定论的行，不能当今年的依据
                    continue
            k = (comp, name)
            if k in out:
                continue
            out[k] = (bucket, SRC_PRIOR)
            n += 1
    return n


def _read_prior_bkd(wb, shorts, notes):
    """上年 BKD 明细页 → ({(账簿,科目): (口径, 本期账面数)}, 命中页数)。

    行列都按表头文字现场定位，不写死位置：各年模板列位不同，
    锁死列号会在换客户时整体错位而**不报错**（读到别列的数，比读不到更危险）。
    """
    out, n_sheets = {}, 0
    for ws in wb.worksheets:
        # 找表头行：同时有「科目名称」和（账簿/账套）——这是 BKD 明细页，不是 mapping 页
        hdr = None
        for r in range(1, min(ws.max_row or 0, 60) + 1):
            texts = [_T(ws.cell(r, c).value) for c in range(1, min(ws.max_column or 0, 30) + 1)]
            if any("科目名称" in t for t in texts) and any(
                    ("账套" in t or "账簿" in t) for t in texts):
                hdr = r
                break
        if not hdr:
            continue
        cols = {}
        oar_col = short_col = None
        for c in range(1, min(ws.max_column or 0, 40) + 1):
            t = _T(ws.cell(hdr, c).value)
            if not t:
                continue
            if "科目名称" in t and "name" not in cols:
                cols["name"] = c
            elif ("账套" in t or "账簿" in t) and "book" not in cols:
                cols["book"] = c
            if is_oar_header(t) and oar_col is None:
                oar_col = c
            if any(w in t for w in SHORT_HDR_WORDS) and short_col is None:
                short_col = c
        if "name" not in cols or "book" not in cols:
            continue
        # 必须像 BKD 明细表：要有「本期账面数 / 本期未审数 / 本期审定数」这类列。
        # 不加这道，序时账页（账簿 + 科目编码 + 摘要）也会被当 BKD 读进来，
        # mapping 里会混进「资本公积」「年初未分配利润」这种根本不是费用的假科目。
        _sig = ("本期账面数", "本期未审数", "本期审定数", "本期发生额")
        if not any(any(k in _T(ws.cell(hdr, c).value) for k in _sig)
                   for c in range(1, min(ws.max_column or 0, 40) + 1)):
            continue
        n_sheets += 1
        if not oar_col:
            notes.append("「%s」没有 OAR 口径列，按规则推导" % ws.title)
        amt_col = None
        for c in range(1, min(ws.max_column or 0, 40) + 1):
            t = _T(ws.cell(hdr, c).value)
            if t and ("本期账面数" in t or "本期发生额" in t):
                amt_col = c
                break
        _blank = 0
        for r in range(hdr + 1, min(ws.max_row or 0, hdr + 3000) + 1):
            name = _T(ws.cell(r, cols["name"]).value)
            book = _T(ws.cell(r, cols["book"]).value)
            # 有的客户底稿（VD 那张）连「合计」行都没有，表1 下面直接接
            # 波动说明 / 审计调整表。连续 5 行空就当作数据区结束，
            # 否则整张调整表会被当成「科目」读进 mapping（资本公积、年初未分配利润…）。
            if not name and not book:
                _blank += 1
                if _blank >= 5:
                    break
                continue
            _blank = 0
            if name in ("合计", "总计", "小计"):
                # 表1 数据区到「合计」为止。再往下是 A3 / check / 口径汇总区，
                # 那些行的 B、D 列会被当成「账套 + 科目」读进来，
                # 结果 mapping 里混进 "A3|20832935.73"、"OAR 口径|上期审定数" 这种假科目。
                break
            if not name or not book:
                continue
            # 汇总行的金额/编号也会落到科目列，纯数字的丢掉
            if re.fullmatch(r"[\d,.\-+%()\s]+", name):
                continue
            if book in ("A3", "check", "Diff", "OAR 口径", "上期审定数", "本期审定数",
                        "调整类型", "调整", "合计"):
                continue
            # 调整表的行号（61 / 86 / #ADJ_6）也会落在账套列，不像账套名的丢掉
            if re.fullmatch(r"[\d#\s].{0,8}", book) and not re.search(r"[\u4e00-\u9fa5A-Za-z]{2}", book):
                continue
            bucket = _T(ws.cell(r, oar_col).value) if oar_col else ""
            amt = 0.0
            if amt_col:
                v = ws.cell(r, amt_col).value
                if isinstance(v, (int, float)):
                    amt = float(v)
            out[(_T(book), name)] = (bucket, amt)
            if short_col:
                sh = _T(ws.cell(r, short_col).value)
                if sh:
                    shorts.setdefault(book, sh)
    return out, n_sheets


# ---------------------------------------------------------------- 建 mapping
def build(wb, prior, prior_shorts, cur_rows, ai_map=None, book_map=None,
          company_shorts=None, company_rates=None):
    """写「口径Mapping」sheet。

    prior        {(主体,科目): (口径, 来源)}   —— 主体可能是**账簿名**，这里统一转公司全称
    prior_shorts {主体: 简称}   —— 上年底稿里带的简称（兜底用）
    company_shorts {主体全称: 简称} —— **客户 PBC「公司mapping」表**里的简称，优先用这个
    company_rates  {主体全称: {"bal": 余额汇率, "pl": 损益汇率}} —— 同表的汇率列。
                   公司段就把这两列写出来，BKD 的折算汇率与两个样本池的折算率
                   都 XLOOKUP 到这里；不再单独建一张「公司mapping」页。
    cur_rows     本年 BKD 行（含 company / name / amount）
    ai_map       {匹配键小写: (口径, 说明)} —— AI 或人工确认过的结果，优先级最高
    book_map     {账簿: 公司全称}，把上年账簿名折成本年 BKD 用的公司全称

    返回 info（含待判断清单，供导出给 AI）
    """
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    ai_map = ai_map or {}

    # 上年 BKD 的「账套名称」是账簿名（RAYWING PTE.LTF.美金 / 【云语】帐簿），
    # 本年 BKD 写的是公司全称（RAYWING PTE.LTD. / 深圳市云语科技有限公司）。
    # 不折一次，两边的键永远对不上，本年会全部落成「待判断」。
    _resolver = None
    if book_map:
        try:
            from .expense import resolve_company as _rc
            _resolver = _rc
        except Exception:
            _resolver = None

    def _comp(x):
        s = _T(x)
        if not s or _resolver is None:
            return s
        try:
            c, _how = _resolver(s, book_map)
            return _T(c) or s
        except Exception:
            return s

    prior_norm, prior_short_norm = {}, {}
    for (company, name), (bucket, src) in (prior or {}).items():
        k = (_comp(company), _T(name))
        if k[1] and k not in prior_norm:
            prior_norm[k] = (bucket, src)
    for company, sh in (prior_shorts or {}).items():
        c = _comp(company)
        if c and c not in prior_short_norm:
            prior_short_norm[c] = _T(sh)
    prior = prior_norm
    prior_shorts = prior_short_norm

    if SHEET in wb.sheetnames:
        del wb[SHEET]
    pos = 6 if len(wb.sheetnames) > 6 else len(wb.sheetnames)
    ws = wb.create_sheet(SHEET, pos)

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    t14 = Font(name="宋体", size=14, bold=True)
    t9 = Font(name="宋体", size=9, color="666666")
    sec = Font(name="宋体", size=11, bold=True, color="1F3864")
    secf = PatternFill("solid", fgColor="F2F2F2")
    head = Font(name="宋体", size=10, bold=True, color="FFFFFF")
    headf = PatternFill("solid", fgColor="4472C4")
    t10 = Font(name="宋体", size=10)
    pend = PatternFill("solid", fgColor="FFFF00")     # 未过 AI，待语义匹配
    newf = PatternFill("solid", fgColor="FFF2CC")     # 本年新增
    aifill = PatternFill("solid", fgColor="E2EFDA")   # 由 AI 判断，待人复核

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 46
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 30
    ws.column_dimensions["G"].width = 3
    ws.column_dimensions["H"].width = 32
    ws.column_dimensions["I"].width = 14
    ws.column_dimensions["J"].width = 24
    ws.column_dimensions["K"].width = 22   # 余额汇率
    ws.column_dimensions["L"].width = 22   # 损益汇率

    ws.cell(1, 1, "口径 Mapping（OAR 口径 + 公司简称）").font = t14
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    ws.cell(2, 1, "口径来源优先级：由 AI 判断 > 上年底稿（上年口径Mapping 页 > 上年 BKD 口径列）> 规则推导（待 AI 复核）。"
                  "BKD 的「OAR 口径」列用 XLOOKUP 按 D 列（匹配键）取 C 列，取不到再按 B 列（科目名称）兜底；"
                  "浅绿＝由 AI 判断（备注列写了判断依据与置信度，需人工复核）；"
                  "黄底＝只按科目名关键词推导、还没做语义匹配，见 bkd_work/ai_tasks/oar_bucket_pending.json；"
                  "备注列写「上年同名科目沿用」的行，口径来自上年底稿里同名科目（跨主体），请确认主体是否适用。").font = t9
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)

    # ---------- OAR 段 ----------
    ws.cell(4, 1, "一、OAR 口径 Mapping").font = sec
    ws.cell(4, 1).fill = secf
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=6)

    for i, h in enumerate(["主体", "科目名称", "OAR 口径", "匹配键（主体|科目）", "来源", "备注"], start=1):
        c = ws.cell(HDR_OAR, i, h)
        c.font = head
        c.fill = headf
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")

    rows_oar = []          # [(主体, 科目, 口径, 来源, 备注, 是否待判断)]
    seen = set()

    # 先放上年（保证「上年有、本年没有」的科目也在表里，口径有连续性）
    # 上年也认不出的（按科目名规则推的）照样进表，但标黄并计入待复核 ——
    # 那是去年的猜测，不能因为"在去年的文件里"就当成依据。
    for (company, name), (bucket, src) in sorted(prior.items()):
        k = _norm_key(company, name)
        if k in seen:
            continue
        seen.add(k)
        _is_rule = (src == SRC_RULE)
        # 口径值与 AI 已确认的一致时，来源仍写「由AI判断」并带上依据与置信度 ——
        # 值没变，可信度信息不该因为"上年 mapping 里也有这一行"就丢掉。
        _ai = ai_map.get(k)
        if _ai and _T(_ai[0]) and _T(_ai[0]) == _T(bucket) and not _is_rule:
            rows_oar.append((company, name, bucket, SRC_AI,
                             _T(_ai[1]) or "AI 语义匹配，未写依据", False))
            continue
        rows_oar.append((company, name, bucket, src,
                         "上年底稿里没有口径，按科目名规则推导；待 AI 判定：bkd_work/ai_tasks/oar_bucket_pending.json" if _is_rule else "",
                         _is_rule))

    # 再放本年：能匹配就沿用，匹配不到就待判断
    from .expense import oar_bucket as _ob
    tot_by_name = {}
    for row in cur_rows:
        _n = _T(row.get("name"))
        if _n:
            tot_by_name[_n] = tot_by_name.get(_n, 0.0) + float(row.get("amount") or 0)

    pending = []
    for row in cur_rows:
        company = _T(row.get("company"))
        name = _T(row.get("name"))
        if not company or not name:
            continue
        k = _norm_key(company, name)
        if k in seen:
            continue
        # 1) 平台 AI 语义匹配过（ai_map 由 oarmap-fill 写入，跨年沿用）
        ai = ai_map.get(k)
        if ai:
            seen.add(k)
            rows_oar.append((company, name, ai[0], SRC_AI,
                             _T(ai[1]) or "AI 语义匹配，未写依据", False))
            continue
        # 2) 上年同主体同科目（上年自己也是规则猜的就不算「匹配到」，落到第 4 步等 AI）
        pv = prior.get((company, name))
        if pv and pv[1] != SRC_RULE:
            seen.add(k)
            rows_oar.append((company, name, pv[0], pv[1], "", False))
            continue
        # 3) 仅科目名兜底（跨主体沿用）：来源同样压成「上年底稿」，
        #    「是同名沿用」这条线索留在备注列，不占来源列。
        same_name = [(b, s, cc) for (cc, n), (b, s) in prior.items()
                     if n == name and b and s != SRC_RULE]
        if same_name:
            seen.add(k)
            rows_oar.append((company, name, same_name[0][0], SRC_PRIOR,
                             "上年同名科目沿用（当时主体：%s）" % same_name[0][2], False))
            continue
        # 4) 上年无对应项：V1.2.80 起不再落「未匹配-待判断」。
        #    两种结果，C 列都必须有值（留空会让 BKD 的 XLOOKUP 取回空白，
        #    OAR 按口径 SUMIF 静默漏掉这一块，实测漏过 76 万）：
        #      · 平台 AI 已语义匹配（ai_map 有键，但走到这里说明键不同）→ 规则建议兜底
        #      · 从未做过 AI 匹配 → 规则建议兜底，来源标「规则推导-待AI复核」
        #    规则建议只是「先有个值」，来源列一眼能看出没经过 AI。
        seen.add(k)
        _sg = _suggest(name)
        rows_oar.append((company, name, _sg, SRC_RULE, "上年无对应科目，按名称规则推导；待 AI 判定：bkd_work/ai_tasks/oar_bucket_pending.json",
                         True))
        pending.append({"company": company, "name": name,
                        "amount": round(float(row.get("amount") or 0), 2),
                        "suggest": _ob(name, tot_by_name.get(name, 0)),
                        "ruleSuggest": _sg})

    r = DATA_ROW
    for (company, name, bucket, src, note, is_pend) in rows_oar:
        ws.cell(r, 1, company).font = t10
        ws.cell(r, 2, name).font = t10
        c = ws.cell(r, 3, bucket)
        c.font = t10
        ws.cell(r, 4, "%s|%s" % (company, name)).font = t10
        ws.cell(r, 5, src).font = t10
        ws.cell(r, 6, note).font = t10
        for cc in range(1, 7):
            ws.cell(r, cc).border = border
        if is_pend:
            # 未过 AI 的行：整行黄底，项目组一眼扫到「这几行还没语义匹配」
            ws.cell(r, 3).fill = pend
            ws.cell(r, 5).fill = pend
        elif src == SRC_AI:
            # AI 判过但需复核：只给「来源」和「备注」轻着色，不抢眼
            ws.cell(r, 5).fill = aifill
        elif note:
            # 上年底稿有依据、但带一句需要看一眼的说明（如跨主体同名沿用）
            ws.cell(r, 6).fill = newf
        r += 1
    oar_last = r - 1

    # ---------- 公司段 ----------
    ws.cell(4, 8, "二、公司 Mapping（简称 + 汇率）").font = sec
    ws.cell(4, 8).fill = secf
    ws.merge_cells(start_row=4, start_column=8, end_row=4, end_column=12)

    # 客户 PBC 的「公司mapping」本来就带两列汇率（余额汇率 / 损益汇率），
    # 这一段就是那张表的主体维度部分：简称 + 两列汇率都写在这里，
    # BKD 的折算汇率列、两个样本池的折算率列都 XLOOKUP 指向这一段。
    for i, h in enumerate(["主体全称", "简称", "备注",
                           "余额汇率（本币-人民币）", "损益汇率（本币-人民币）"], start=8):
        c = ws.cell(HDR_COMPANY, i, h)
        c.font = head
        c.fill = headf
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")

    companies = []
    seen_c = set()
    for row in cur_rows:
        c = _T(row.get("company"))
        if c and c not in seen_c:
            seen_c.add(c)
            companies.append(c)
    for c in prior_shorts:
        if c and c not in seen_c:
            seen_c.add(c)
            companies.append(c)

    r = DATA_ROW
    company_shorts = company_shorts or {}
    company_rates = company_rates or {}
    for c in companies:
        # 优先级：客户 PBC「公司mapping」的简称 > 上年底稿带的 > 按字号临时猜
        # 客户自己有一张 mapping 表，简称就用他们的，别再自己造一套：
        # 简称同时是调整表的 SUMIFS 匹配键，两边不一致调整数就会取不到。
        if company_shorts.get(c):
            sh, why = company_shorts[c], "PBC 公司mapping"
        elif prior_shorts.get(c):
            sh, why = prior_shorts[c], "上年底稿"
        else:
            sh, why = guess_short(c), "自动提取，可改"
        ws.cell(r, 8, c).font = t10
        ws.cell(r, 9, sh).font = t10
        ws.cell(r, 10, why).font = t10
        # 汇率两列：客户表有就照抄，没有就别留空 —— XLOOKUP 取到空会把金额算成 0，
        # 比填错更难发现。缺的一律给 1（人民币主体本币即人民币）并在备注里标明。
        _pair = company_rates.get(c) or {}
        _bal = _pair.get("bal")
        _plv = _pair.get("pl")
        if _bal is None and _plv is None:
            _bal = _plv = 1.0
            ws.cell(r, 10, why + "；mapping 未给汇率，按 1").font = t10
        _bal = _bal if _bal is not None else _plv
        _plv = _plv if _plv is not None else _bal
        for _i, _v in enumerate((_bal, _plv)):
            _c = ws.cell(r, 11 + _i, _v)
            _c.font = t10
            # 汇率不四舍五入（用户口径）
            _c.number_format = "General"
            _c.alignment = Alignment(horizontal="right", vertical="center")
        for cc in (8, 9, 10, 11, 12):
            ws.cell(r, cc).border = border
        r += 1
    comp_last = r - 1

    ws.freeze_panes = "A6"
    # Python 侧再算一份「主体|科目 → 口径」：BKD 的 R 列现在是 XLOOKUP 公式，
    # openpyxl 读回来是字符串，OAR 页要靠这份才能知道有哪些口径。
    resolved = {}
    samples = {}
    for (company, name, bucket, src, note, is_pend) in rows_oar:
        resolved[(_T(company), _T(name))] = bucket or SRC_RULE
        # 参照系只收「已定论的」行：待匹配行的口径本身是规则猜的，
        # 拿它当例子给 AI 看，等于让 AI 参照自己的猜测。
        if bucket and not is_pend:
            lst = samples.setdefault(bucket, [])
            if name not in lst and len(lst) < 8:
                lst.append(name)

    return {
        "sheet": SHEET,
        "oarRows": len(rows_oar), "oarLast": oar_last,
        "companyRows": len(companies), "companyLast": comp_last,
        # 公司段的表头行与数据范围：BKD/样本池的折算汇率 XLOOKUP 这一段
        "companyHeaderRow": HDR_COMPANY, "companyFirst": DATA_ROW,
        "pending": pending,
        "pendingCount": len(pending),
        # 有多少行是 AI 判的、多少行还只是规则推导 —— 生成报告的「待复核」数字
        "aiRows": sum(1 for x in rows_oar if x[3] == SRC_AI),
        "ruleRows": sum(1 for x in rows_oar if x[3] == SRC_RULE),
        # 已有口径 + 每类 8 个典型科目名：AI 做语义匹配时的参照系
        "bucketSamples": samples,
        "resolved": resolved,
    }


_SUGGEST_RULES = [
    ("股权激励", ("股权激励", "股份支付", "限制性股票", "期权", "员工持股")),
    ("折旧摊销", ("折旧", "摊销")),
    ("广告费", ("广告", "宣传", "推广", "展会")),
    ("薪酬", ("薪酬", "工资", "奖金", "提成", "社会保险", "公积金", "福利",
              "职工教育", "工会", "离职补偿", "辞退", "劳务", "董事", "监事",
              "年金", "补贴", "津贴", "社保", "人力成本", "年终奖")),
]


def _suggest(name):
    """给 AI 的建议口径（只是建议，最终以 AI/人工判断为准）。"""
    s = str(name or "")
    core = s.split("_", 1)[1] if "_" in s else s
    for label, words in _SUGGEST_RULES:
        if any(w in core for w in words):
            return label
    return "其他"


# ---------------------------------------------------------------- AI 通道
def export_pending(info, meta, path):
    """把「上年无对应项」的科目导出成 JSON，交给平台 AI 做语义匹配。

    为什么要带上 existingBuckets：AI 不是凭常识分类，而是把新科目**映射到这家客户
    已经用了一年的口径**上（薪酬 / 折旧摊销 / 广告费 / 股权激励 / 单列大额 / 其他）。
    没有参照系，同一个「技术服务费」在两家客户会被分到不同桶，跨年就不可比。
    """
    import json
    payload = {
        "task": "oar_bucket",
        "instruction": "为每个科目判断 OAR 口径：只能取下面 choices 里的值，或套用 existingBuckets 里"
                       "已存在的口径名。判断依据是该科目在序时账里实际核算的业务实质（优先看科目名"
                       "核心词，其次看金额规模）。必须给 reason（判断依据，一句话说明为什么落到这个"
                       "口径，可引用已有的同义科目）与 confidence（0~1 的小数，或「高/中/低」）。"
                       "同科目名在不同主体下应给同一口径，除非业务实质明显不同。",
        "meta": meta,
        # 这家客户已经用了一整年的口径 + 典型科目名 —— AI 的参照系
        "existingBuckets": info.get("bucketSamples") or {},
        "choices": ["薪酬", "折旧摊销", "广告费", "股权激励", "其他"],
        "outputFormat": {"items": [{"company": "主体全称", "name": "科目名称",
                                    "bucket": "口径", "reason": "判断依据",
                                    "confidence": 0.9}]},
        "items": info.get("pending") or [],
    }
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)
    return path


def _conf_txt(v):
    """置信度统一成「92%（高）」这种写法：数字给百分比，文字原样保留。"""
    if v is None or v == "":
        return ""
    if isinstance(v, (int, float)):
        f = float(v)
        if f <= 1.0:
            f *= 100
        lvl = "高" if f >= 85 else ("中" if f >= 60 else "低")
        return "%.0f%%（%s）" % (f, lvl)
    s = _T(v)
    return s


def load_ai_map(path):
    """读 AI 判断结果 → {匹配键小写: (口径, 备注)}。

    备注拼成「依据：…；置信度 92%（高）」，直接落到 mapping 的备注列，
    复核的人不用再翻原始 JSON 就知道 AI 凭什么这么判、有多大把握。
    两种格式都收：{"items":[{"company","name","bucket","reason","confidence"}]}
    或 {"公司|科目": "口径"}。AI 输出经常前后加说明文字，所以做宽松解析。
    """
    import json
    import os
    if not path or not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            d = json.load(fh)
    except Exception:
        return {}
    out = {}
    items = d.get("items") if isinstance(d, dict) else d
    if isinstance(items, list):
        for it in items:
            if not isinstance(it, dict):
                continue
            comp = _T(it.get("company"))
            name = _T(it.get("name"))
            b = _T(it.get("bucket") or it.get("oar") or it.get("口径"))
            if comp and name and b:
                out[_norm_key(comp, name)] = (b, _ai_note(it))
    elif isinstance(d, dict):
        for k, v in d.items():
            if isinstance(v, str) and "|" in k:
                out[k.strip().lower()] = (_T(v), "依据：AI 按科目名语义归类（未给置信度）")
    return out


def _ai_note(it):
    """AI 结论 → mapping 备注列的一句话。"""
    reason = _T(it.get("reason") or it.get("依据") or "")
    conf = _conf_txt(it.get("confidence", it.get("置信度")))
    parts = []
    if reason:
        parts.append("依据：" + reason)
    if conf:
        parts.append("置信度 " + conf)
    return "；".join(parts) or "AI 语义匹配（未给依据与置信度，建议复核）"


def ai_map_path():
    """AI/人工确认过的口径，落在 <data>/bkd_work/oar_map_ai.json（跨年沿用）。

    跟着 KZ_DATA_DIR 走：测试用 --data-dir 指到临时目录时，不能把结果写进用户真实库。
    """
    import os
    base = os.environ.get("KZ_DATA_DIR")
    if base:
        return os.path.join(base, "bkd_work", "oar_map_ai.json")
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(os.path.normpath(os.path.join(here, "..", "..", "..",
                                                      "data", "bkd_work")), "oar_map_ai.json")


def read_items(path):
    """读一份 AI 结果文件，返回 [{company,name,bucket,reason,confidence}]（宽松解析）。"""
    import json
    import os
    if not path or not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as fh:
            d = json.load(fh)
    except Exception:
        return []
    items = d.get("items") if isinstance(d, dict) else d
    out = []
    if isinstance(items, list):
        for it in items:
            if isinstance(it, dict) and _T(it.get("company")) and _T(it.get("name")):
                b = _T(it.get("bucket") or it.get("oar") or it.get("口径"))
                if b:
                    out.append({"company": _T(it.get("company")), "name": _T(it.get("name")),
                                "bucket": b, "reason": _T(it.get("reason") or ""),
                                "confidence": it.get("confidence", it.get("置信度")) or ""})
    elif isinstance(d, dict):
        for k, v in d.items():
            if isinstance(v, str) and "|" in k:
                comp, name = k.split("|", 1)
                out.append({"company": comp.strip(), "name": name.strip(),
                            "bucket": _T(v), "reason": "", "confidence": ""})
    return out


def load_ai_map_all(path=None):
    """正本 + 结果侧文件一起读，侧文件优先（它是最新一次 AI 判断）。

    任务卡让外部 AI 只写 `oar_map_ai_result.json`，不动正本 —— 正本里可能有
    上一年人工确认过、明年还要复用的行，不能被一次自动回填冲掉。
    """
    import os
    path = path or ai_map_path()
    out = load_ai_map(path)
    try:
        from .tasks import result_path
        _rp = result_path(path)
    except Exception:
        _rp = ""
    if _rp and os.path.exists(_rp):
        try:
            out.update(load_ai_map(_rp))
        except Exception:
            pass
    return out


def save_ai_map(raw_items, path=None):
    """把 AI/人工确认的口径合并进 oar_map_ai.json（同键覆盖）。返回总条数。

    依据与置信度一起落盘 —— 备注列的「依据：…；置信度 92%（高）」就是从这儿来的，
    明年复用这条口径时，复核的人仍看得到当初凭什么这么判。
    """
    import json
    import os
    path = path or ai_map_path()
    old = read_items(path)
    idx = {}
    for it in old:
        idx[_norm_key(it["company"], it["name"])] = it
    n_new = 0
    for it in raw_items or []:
        k = _norm_key(it.get("company"), it.get("name"))
        if not k or not _T(it.get("bucket")):
            continue
        if k not in idx:
            n_new += 1
        idx[k] = {"company": _T(it.get("company")), "name": _T(it.get("name")),
                  "bucket": _T(it.get("bucket")),
                  "reason": _T(it.get("reason") or "") or "人工指定",
                  "confidence": it.get("confidence") or ""}
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump({"items": list(idx.values())}, fh, ensure_ascii=False, indent=1)
    return {"total": len(idx), "added": n_new, "path": path}
