# -*- coding: utf-8 -*-
"""费用底稿上数：VC.00 销售费用 BKD / VD.00 管理费用 BKD。

设计约束（来自 03_文档/审计底稿上数逻辑说明_BKD_OAR_TOD.md v2 与 V1.2 备忘）：
- 只写 E 列（本期账面数）和 B/C/D 描述列；G/H/I/J 等一律是模板公式，不动。
- 列位、表头行、合计行全部按表头文字动态定位，不写死列字母（空白模板与项目版列不同）。
- 取不到的数据标黄 FFFF00 并进待确认清单；绝不填 0 冒充，也不给「建议值」。
"""

import os
import re
import io

EXP_TPL_NAME = "U_exp SWP VC&VD 模板.xlsx"
SHEET_VC = "VC.00 销售费用BKD"
SHEET_VD = "VD.00 管理费用BKD"

# 表头正则（顺序即优先级，单值字段匹配到第一列就停）
SPEC = [
    ("book",   re.compile(r"账套|账簿")),
    ("code",   re.compile(r"总账科目编码|科目编码|^编码$")),
    ("name",   re.compile(r"科目名称|^科目$")),
    ("amount", re.compile(r"本期账面数|本期发生额|本期账面|本年账面数")),
    ("prior",  re.compile(r"上期审定数|上年审定")),
]

# 会计专用（无币种）：两位小数、千分位、0 显示成「-」，负数带括号。
# 用户 2026-11-12 口径：本平台产出的所有数字格统一这个格式（分析文本不动）。
ACC_FMT = r'_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '
MONEY_FMT = ACC_FMT
GAP_COLOR = "FFFF00"

#: 「For X-ref」索引：这些费用已被别的底稿测过，BKD 这里只做交叉索引、不重复测。
#: 顺序有讲究 —— 先认「无形资产/长期待摊/使用权资产」这些具体资产，
#: 再落到「固定资产」，最后才是「应付职工薪酬」；反了会把无形资产摊销标成固定资产。
XREF_RULES = [
    (("无形资产", "软件服务费摊销", "软件摊销"), "无形资产"),
    (("长期待摊", "待摊费用"), "长期待摊费用"),
    (("使用权资产",), "使用权资产"),
    (("累计折旧", "折旧费", "固定资产"), "固定资产"),
    (("股权激励", "股份支付"), "应付职工薪酬"),
    (("薪酬", "工资", "奖金", "提成", "社会保险", "公积金", "福利", "职工教育",
      "工会", "离职补偿", "辞退", "劳务", "董事", "监事", "年金", "补贴", "津贴",
      "社保", "人力成本", "年终奖"), "应付职工薪酬"),
]


def xref_of(acct_name, balance_date=""):
    """科目名 → For X-ref 索引，如「应付职工薪酬20260930」。日期跟着资产负债表日走。"""
    core = str(acct_name or "")
    core = core.split("_", 1)[1] if "_" in core else core
    tag = ""
    for words, label in XREF_RULES:
        if any(w in core for w in words):
            tag = label
            break
    if not tag:
        return None
    d = str(balance_date or "").replace("-", "").replace("/", "").strip()
    return tag + d if d else tag

# ---------------------------------------------------------------- OAR 口径
# 分析性复核要按「业务性质」看，不是按科目名看。78 个科目逐行列出来，
# 复核人看不出「薪酬一共多少、比去年多了多少」——那是 OAR 要回答的问题。
# 口径只做归类、不改金额：同一个科目只会落进一个口径，SUMIF 汇总后总额与 BKD 合计相等。
OAR_BUCKET_RULES = [
    ("薪酬", ("薪酬", "工资", "奖金", "提成", "社会保险", "公积金", "福利",
             "职工教育", "工会", "离职补偿", "辞退", "劳务", "董事", "监事",
             "年金", "补贴", "津贴", "社保")),
    ("折旧摊销", ("折旧", "摊销")),
    ("广告费", ("广告", "宣传", "推广", "展会", "市场推广", "市场费用")),
]
OAR_BUCKET_OTHER = "其他"
OAR_BIG = 1_000_000.0        # 单科目达到这个金额才单列（约 2×TE）


def oar_bucket(name, amount=0.0, big=None):
    """科目名 → OAR 口径。

    归不进的科目按金额处理：达到 big 的单独成一类（用去掉「销售费用_」前缀的
    科目名），否则归「其他」。客户每年总有一两个"说不清但金额巨大"的科目
    （技术服务费、中介费、股权激励），混进「其他」会被淹没，单列才看得见。
    """
    nm = str(name or "").strip()
    core = nm.split("_", 1)[1] if "_" in nm else nm
    for label, words in OAR_BUCKET_RULES:
        for w in words:
            if w in core:
                return label
    lim = OAR_BIG if big is None else big
    if abs(float(amount or 0)) >= lim and core:
        return core[:12]
    return OAR_BUCKET_OTHER


def _T(v):
    if v is None:
        return ""
    return str(v).replace("\u00a0", " ").strip()


def _r2(v):
    return round(float(v or 0), 2)


# ---------------------------------------------------------------- 分段
def read_expense_by_segment(conn, period_from="1", period_to="6", pre_end="6"):
    """按预审段 / 年审段拆分费用发生额，返回 {pre, final, total, preEnd}。

    为什么要拆：TOD 的样本池和预提冲销配对都按段做（Skill §5.3-5.4），
    跨段配对会把一段的净额挪到另一段。费用 BKD 本身不拆（账面数是全年口径），
    所以这里只提供金额，不改 BKD 输出。

    外币折算不在这里做：BKD 取的是序时账本币，汇率只影响 TOD 池的
    「未审本币」列（预审段与年审段可能用不同汇率，见 Skill §5.3）。
    """
    from core import recon
    try:
        pf, pt, pe = int(period_from), int(period_to), int(pre_end or period_to)
    except (TypeError, ValueError):
        return {"error": "期间必须是数字月份"}
    pe = max(pf, min(pe, pt))          # 预审段不能超出总期间

    def _amt(a, b):
        if b < a:
            return 0.0
        try:
            cond = recon._period_cond(conn, str(a), str(b), ["period"])
        except Exception as e:
            raise RuntimeError("期间条件构建失败：%s" % e)
        where = (" AND " + cond) if cond else ""
        sql = """
            SELECT COALESCE(SUM(
                     COALESCE(debit, 0)
                     - (COALESCE(credit, 0)
                        - CASE WHEN COALESCE(summary, '') LIKE '%%结转%%'
                               THEN COALESCE(credit, 0) ELSE 0 END)
                   ), 0)
            FROM journal
            WHERE 1=1 %s
              AND (
                    COALESCE(acct_l1, '') LIKE '%%销售费用%%'
                 OR COALESCE(acct_l1, '') LIKE '%%管理费用%%'
                 OR acct_name LIKE '销售费用%%'
                 OR acct_name LIKE '管理费用%%'
                 OR acct_name LIKE '%%_销售费用%%'
                 OR acct_name LIKE '%%_管理费用%%'
              )
        """ % where
        return _r2(conn.execute(sql).fetchone()[0] or 0)

    try:
        pre = _amt(pf, pe)
        fin = _amt(pe + 1, pt)
    except Exception as e:
        return {"error": str(e)}
    return {"preEnd": pe, "pre": pre, "final": fin, "total": _r2(pre + fin),
            "period": "%s-%s" % (pf, pt)}


# ---------------------------------------------------------------- 取数
def read_expense(conn, period_from="1", period_to="6"):
    """从 journal（序时账）按「账簿 × 科目编码」汇总费用发生额。

    返回 (rows, meta)。rows 每项：
      {book, code, name, l1, amount, dr, cr}
    amount = 借方 − 贷方（费用科目正常只有借方；贷方是红字冲销/转出，要抵掉）。
    """
    from core import recon   # 复用期间归一逻辑（数字 / YYYY-MM / 2026年1期 都吃得下）

    try:
        cond = recon._period_cond(conn, period_from, period_to, ["period"])
    except Exception as e:
        return [], {"error": "期间条件构建失败：%s" % e}

    where_period = (" AND " + cond) if cond else ""
    sql = """
        SELECT
            COALESCE(NULLIF(trim(book_name), ''), '(未标注账簿)') AS book,
            COALESCE(NULLIF(trim(acct_code), ''), '')            AS code,
            COALESCE(NULLIF(trim(acct_name), ''), '(未命名科目)') AS name,
            COALESCE(acct_l1, '')                                AS l1,
            SUM(COALESCE(debit, 0))                              AS dr,
            SUM(COALESCE(credit, 0))                             AS cr,
            SUM(CASE WHEN COALESCE(summary, '') LIKE '%%结转%%'
                     THEN COALESCE(credit, 0) ELSE 0 END)        AS cr_close
        FROM journal
        WHERE 1=1
          %s
          AND (
                COALESCE(acct_l1, '') LIKE '%%销售费用%%'
             OR COALESCE(acct_l1, '') LIKE '%%管理费用%%'
             OR acct_name LIKE '销售费用%%'
             OR acct_name LIKE '管理费用%%'
             OR acct_name LIKE '%%_销售费用%%'
             OR acct_name LIKE '%%_管理费用%%'
          )
        GROUP BY 1, 2, 3, 4
        ORDER BY 1, 2
    """ % where_period
    try:
        raw = conn.execute(sql).fetchall()
    except Exception as e:
        return [], {"error": "从序时账取费用数据失败：%s" % e}

    rows = []
    close_total = 0.0
    for book, code, name, l1, dr, cr, cr_close in raw:
        dr, cr, cr_close = _r2(dr), _r2(cr), _r2(cr_close)
        # 本期账面数 = 借方 − （贷方 − 结转损益贷方）
        # 关键：月末「结转本期损益」会把费用科目贷方冲平，若不剔除，每个科目都会算成 0。
        # 真实红字冲销仍照扣（它们通常体现在贷方且摘要不含"结转"）。
        close_total += cr_close
        rows.append({
            "book": _T(book), "code": _T(code), "name": _T(name), "l1": _T(l1),
            "dr": dr, "cr": cr, "cr_close": cr_close,
            "amount": _r2(dr - (cr - cr_close)),
        })
    # 净额为 0 的行不进底稿（本期无发生额）
    rows = [r for r in rows if abs(r["amount"]) > 0.004]

    meta = {
        "rows": len(rows),
        "sum": _r2(sum(r["amount"] for r in rows)),
        "books": len({r["book"] for r in rows}),
        "period": "%s-%s" % (period_from, period_to),
        "close_out": _r2(close_total),
    }
    return rows, meta


def split_vc_vd(rows):
    """按一级科目把费用行分成销售费用(VC)与管理费用(VD)。"""
    vc, vd, other = [], [], []
    for r in rows:
        tag = r["l1"] or r["name"]
        if "销售费用" in tag or r["name"].startswith("销售费用"):
            vc.append(r)
        elif "管理费用" in tag or r["name"].startswith("管理费用"):
            vd.append(r)
        else:
            other.append(r)
    return vc, vd, other


def read_book_map(conn, extra=None):
    """账簿 → 公司 的映射。优先级：本次传入 > 导入时的主体 mapping（_books）> 历史对照表 > tb。

    顺序很重要：科目余额表通常没有「账簿名称」列（只有合并方案层级），
    tb 往往给不出映射。但**导入阶段已经把主体 mapping 落进 `_books`**（账簿→公司全称），
    正常情况下这里直接就能读到，项目组不用再传一份「账簿↔公司对照表」。
    历史对照表只作兜底：换客户时残留的旧映射不能压过本次导入的数据。
    """
    out = {}
    # 0) 本次显式传入（接口/命令行给的，最高优先）
    if extra:
        out.update({_T(k): _T(v) for k, v in extra.items() if _T(k) and _T(v)})

    try:
        tabs = [x[0] for x in conn.execute("SHOW TABLES").fetchall()]
    except Exception:
        tabs = []

    # 1) 导入时的主体 mapping：与本次账套同源，最权威
    if "_books" in tabs:
        try:
            for book, company in conn.execute(
                    "SELECT book_name, company_name FROM _books "
                    "WHERE book_name IS NOT NULL AND trim(book_name) <> ''").fetchall():
                b, c = _T(book), _T(company)
                if b and c:
                    out.setdefault(b, c)
        except Exception:
            pass

    # 2) 历史上传过的对照表（data/bkd_work/book_map.json）
    try:
        p = _book_map_path()
        if os.path.exists(p):
            import json as _json
            with open(p, "r", encoding="utf-8") as fh:
                d = _json.load(fh)
            for k, v in (d or {}).items():
                if _T(k) and _T(v):
                    out.setdefault(_T(k), _T(v))
    except Exception:
        pass
    # 3) tb 里已经映射好的
    try:
        if "tb" in tabs:
            for book, company in conn.execute(
                    "SELECT DISTINCT book_name, company_name FROM tb "
                    "WHERE book_name IS NOT NULL AND trim(book_name) <> ''").fetchall():
                b, c = _T(book), _T(company)
                if b and c:
                    out.setdefault(b, c)
    except Exception:
        pass
    return out


def _book_map_path():
    here = os.path.dirname(os.path.abspath(__file__))     # modules/bkd
    data_dir = os.path.normpath(os.path.join(here, "..", "..", "..", "data", "bkd_work"))
    return os.path.join(data_dir, "book_map.json")


def save_book_map(pairs):
    """保存账簿→公司对照（覆盖）。返回条数。"""
    import json as _json
    p = _book_map_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    clean = {_T(k): _T(v) for k, v in (pairs or {}).items() if _T(k) and _T(v)}
    old = {}
    try:
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as fh:
                old = _json.load(fh) or {}
    except Exception:
        old = {}
    old.update(clean)
    with open(p, "w", encoding="utf-8") as fh:
        _json.dump(old, fh, ensure_ascii=False, indent=1)
    return len(old)


def parse_book_map_xlsx(data):
    """解析「主体 mapping 表」：找含 账簿/账套 与 公司 的两列，返回 {账簿: 公司}。"""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    pairs = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(max_row=min(ws.max_row or 1, 3000), values_only=True))
        if not rows:
            continue
        ci_book = ci_co = None
        head_i = 0
        for i, r in enumerate(rows[:12]):
            for j, v in enumerate(r):
                t = _T(v)
                if not t:
                    continue
                if ci_book is None and re.search(r"账簿|账套", t):
                    ci_book = j
                elif ci_co is None and re.search(r"公司|主体|单位", t):
                    ci_co = j
            if ci_book is not None and ci_co is not None:
                head_i = i
                break
        if ci_book is None or ci_co is None:
            continue
        for r in rows[head_i + 1:]:
            if not r:
                continue
            b = _T(r[ci_book]) if ci_book < len(r) else ""
            c = _T(r[ci_co]) if ci_co < len(r) else ""
            if b and c and b != c:
                pairs[b] = c
        if pairs:
            break
    wb.close()
    return pairs


def norm_book_key(s):
    """账簿名归一：去空格、全半角、去掉尾部 (101) 这类编码后缀。"""
    t = _T(s)
    if not t:
        return ""
    t = t.replace("（", "(").replace("）", ")")
    t = re.sub(r"\(\s*\d{2,4}\s*\)\s*$", "", t)
    return re.sub(r"\s+", "", t).upper()


def resolve_company(book, book_map):
    """两级匹配：先精确，再归一。返回 (公司名, 命中方式)。"""
    if book in book_map:
        return book_map[book], "精确"
    k = norm_book_key(book)
    for b, c in book_map.items():
        if norm_book_key(b) == k:
            return c, "归一"
    return "", "未命中"


# ---------------------------------------------------------------- 写模板
def _find_tpl(tpl_dir):
    for name in (EXP_TPL_NAME,):
        p = os.path.join(tpl_dir, name)
        if os.path.exists(p):
            return p
    for f in sorted(os.listdir(tpl_dir)) if os.path.isdir(tpl_dir) else []:
        if f.lower().endswith((".xlsx", ".xlsm")) and ("exp" in f.lower() or "费用" in f):
            return os.path.join(tpl_dir, f)
    return None


def _locate(ws):
    """定位表头行与列。返回 {字段: 列号, "_header_row": 行号}；关键字段缺失返回 None。"""
    best, best_score = None, 0
    for r in range(1, min(ws.max_row, 60) + 1):
        hit, sc = set(), 0
        for c in range(1, min(ws.max_column, 30) + 1):
            t = _T(ws.cell(r, c).value)
            if not t or len(t) > 40:
                continue
            for key, rx in SPEC:
                if key in hit or not rx.search(t):
                    continue
                hit.add(key)
                sc += 1
        if sc > best_score:
            best, best_score = r, sc
    if best_score < 3:
        return None
    col = {"_header_row": best}
    taken = set()
    for key, rx in SPEC:
        for c in range(1, min(ws.max_column, 30) + 1):
            if c in taken:
                continue
            t = _T(ws.cell(best, c).value)
            if t and len(t) <= 40 and rx.search(t):
                col[key] = c
                taken.add(c)
                break
    if not col.get("book") or not col.get("name") or not col.get("amount"):
        return None
    return col


def _find_sum_row(ws, data_start, col):
    for r in range(data_start, min(ws.max_row, data_start + 400) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().upper().startswith("=SUM("):
            return r
    return None


PY_NEW_COLS = 2     # 「上期未审数」「上期调整数」两列

#: 外币主体折算（香港 / 新加坡两个美金账簿）
#: 插在「本年账面数」后面两列：汇率、本年账面数（人民币）
FX_SHIFT = 2
FX_RATE = 6.8454
#: 账簿名出现这些字样就按外币折（这两个主体的记账本位币是美元）
FX_BOOK_HINTS = ("香港", "新加坡", "RAYWING", "PTE", "HK")
#: 公司维度（简称 + 余额/损益汇率）在底稿里就落在「口径Mapping」的公司段，
#: 不单独建一张 sheet：口径Mapping 本身已经按主体列了简称，再加两列汇率即可。
CM_SHEET = "口径Mapping"


def fx_rate_of(book):
    """这个账簿的折算汇率：外币主体给 FX_RATE，人民币主体给 1。"""
    b = str(book or "").upper()
    return FX_RATE if any(h.upper() in b for h in FX_BOOK_HINTS) else 1.0


def load_company_map(conn):
    """读客户 PBC 的「公司mapping」表 → {"shorts": {全称: 简称}, "books": {账簿名: 全称}}。

    为什么要读客户的：
    1) 简称同时是「调整汇总表」的 SUMIFS 匹配键，自己按字号猜的（瑞云/瑞影）
       跟客户那张表对不上，调整数会 SUMIFS 取空、BKD 简称列还会显示「未匹配」；
    2) 上年底稿的 B 列是**2025 年的账簿名**（「瑞影科技」「瑞影广州分公司」），
       用 2026 的账簿映射折不动，补行就会把这些旧账簿名直接写进本年 B 列。
       客户 mapping 里 2025/2026 两列账簿名都有，正好用它折。
    """
    out = {"shorts": {}, "books": {}, "compacts": {}, "table": None, "rates": {}}
    try:
        cols = [d[0] for d in conn.execute("describe _files").fetchall()]
    except Exception:
        return out
    c_name = next((c for c in cols if "name" in str(c).lower()), None)
    c_path = next((c for c in cols if str(c).lower() in ("path", "file_path", "storage_path")), None)
    if not c_name or not c_path:
        return out
    try:
        files = conn.execute('select "%s", "%s" from _files' % (c_name, c_path)).fetchall()
    except Exception:
        return out
    path = None
    for fn, fp in files or []:
        low = str(fn or "").lower()
        if "mapping" in low or "对照" in low:
            path = fp
            break
    if not path:
        return out
    try:
        cur = conn.execute("select * from read_xlsx(?, header=true, all_varchar=true)", [path])
        hdr = [d[0] for d in cur.description]
        rows = cur.fetchall()
    except Exception:
        return out

    i_short = i_full = None
    i_bal = i_pl = None
    book_cols = []
    for i, c in enumerate(hdr):
        t = str(c or "").strip()
        if i_short is None and "简称" in t:
            i_short = i
        if i_full is None and ("全称" in t or "公司名称" in t or "主体名称" in t):
            i_full = i
        if "账簿" in t or "帐簿" in t:
            book_cols.append(i)
        # 按公司维度给的两列汇率：余额类科目用「余额汇率」，损益类用「损益汇率」
        if "汇率" in t:
            if "余额" in t and i_bal is None:
                i_bal = i
            elif "损益" in t and i_pl is None:
                i_pl = i
    if i_full is None:
        return out
    out["table"] = {"headers": hdr, "rows": rows, "i_full": i_full,
                    "i_bal": i_bal, "i_pl": i_pl}
    for r in rows:
        full = str(r[i_full] or "").strip()
        if not full:
            continue
        if i_short is not None:
            sh = str(r[i_short] or "").strip()
            if sh:
                out["shorts"][full] = sh
        for bc in book_cols:
            bk = str(r[bc] or "").strip()
            if bk:
                out["books"][bk] = full
                out["books"][norm_book_key(bk)] = full
    # 汇率：按主体收两类；空或 0 一律当「没填」（0 会把金额吃成 0，比空更危险）
    if i_bal is not None or i_pl is not None:
        for r in rows:
            full = str(r[i_full] or "").strip()
            if not full:
                continue
            pair = {}
            for tag, idx in (("bal", i_bal), ("pl", i_pl)):
                if idx is None:
                    continue
                try:
                    val = float(str(r[idx]).replace(",", "").strip())
                except (TypeError, ValueError):
                    val = 0.0
                pair[tag] = val if abs(val) > 0 else None
            if pair:
                out["rates"][full] = pair

    # 紧凑键：全称、各账簿名、简称都压一遍；同一键对应多个主体时**丢弃**（宁可不折，不能折错）
    cand = {}
    for full in set(list(out["shorts"].keys()) + [v for v in out["books"].values()]):
        keys = {compact_name(full)}
        for bk, f2 in out["books"].items():
            if f2 == full:
                keys.add(compact_name(bk))
        sh = out["shorts"].get(full)
        if sh:
            keys.add(compact_name(sh))
        for k in keys:
            if not k:
                continue
            if k in cand and cand[k] != full:
                cand[k] = None          # 冲突，标记为不可用
            else:
                cand.setdefault(k, full)
    out["compacts"] = {k: v for k, v in cand.items() if v}
    return out


def load_company_short(conn):
    """只要简称时用这个（兼容旧调用）。"""
    return load_company_map(conn).get("shorts") or {}


#: 上期底稿常用「字号 + 分支」当主体名（「瑞影科技」「瑞影广州分公司」），
#: 不是全称也不是客户 mapping 的简称。做一把「去掉地域/公司性质/编号」的紧凑键，
#: 两边用同一套规则压一次再比，就能对上。
_NOISE_WORDS = ("股份有限公司", "有限责任公司", "有限公司", "小合并", "公司",
                "科技", "技术", "信息技术", "信息", "网络", "实业", "集团", "控股",
                "计算机", "无限", "文化", "传媒", "服务", "投资", "管理")
_PLACES = ("深圳市", "深圳", "上海市", "上海", "北京市", "北京", "广州市", "广州",
           "杭州市", "杭州", "成都市", "成都", "西安市", "西安", "苏州市", "苏州",
           "东莞市", "东莞", "武汉市", "武汉", "南京市", "南京", "长沙市", "长沙",
           "重庆市", "重庆", "盐城市", "盐城", "贵安新区", "贵安", "呼和浩特市",
           "呼和浩特", "鄂尔多斯市", "鄂尔多斯", "江西省", "江西", "香港", "新加坡")


def compact_name(s):
    """把主体名压成可比的核心键：去地域、去公司性质、去编号与标点。"""
    t = _T(s)
    if not t:
        return ""
    t = re.sub(r"^[\d\-－—\.\s、]+", "", t)          # 开头的 10-1 / 4-2 编号
    m = re.search(r"[【\[]([^】\]]+)[】\]]", t)          # 【瑞影】账簿(021) → 瑞影
    if m and _T(m.group(1)):
        t = _T(m.group(1))
    t = re.sub(r"[（(][^）)]*[）)]", "", t)              # 括号内容
    for w in _PLACES + _NOISE_WORDS:
        t = t.replace(w, "")
    return re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9]", "", t).upper()


#: 「余额汇率」适用的科目（资产/负债类）；费用 BKD 基本都是损益类，走「损益汇率」。
#: 这条规则放在一处，将来同一张表混进余额类科目时不用满地改。
RATE_BAL_HINTS = ("存款", "借款", "应收", "应付", "预付", "预收", "其他应收", "其他应付",
                  "货币资金", "短期投资", "长期投资", "存货")


def rate_kind_of(subject):
    """这个科目该用哪一列汇率：余额汇率(bal) 还是 损益汇率(pl)。"""
    nm = _T(subject)
    for w in RATE_BAL_HINTS:
        if w in nm:
            return "bal"
    return "pl"


def find_company_map_ref(wb, name=CM_SHEET):
    """在「口径Mapping」的公司段里定位表头 → 供公式与取值用的引用信息。

    公司段不是从第 1 行开始的（上面是口径明细），所以要按表头文字找那一行：
    同时含「主体全称」与「汇率」。返回 {"sheet","full","bal","pl","r1","r2"}；
    找不到（没有这张表或客户没给汇率列）返回 None。
    """
    from openpyxl.utils import get_column_letter

    if name not in wb.sheetnames:
        return None
    ws = wb[name]
    hdr = None
    for r in range(1, min(ws.max_row or 0, 60) + 1):
        texts = [_T(ws.cell(r, c).value) for c in range(1, min(ws.max_column or 0, 30) + 1)]
        if any("主体全称" in t for t in texts) and any("汇率" in t for t in texts):
            hdr = r
            break
    if not hdr:
        return None
    i_full = i_bal = i_pl = None
    for c in range(1, min(ws.max_column or 0, 30) + 1):
        t = _T(ws.cell(hdr, c).value)
        if not t:
            continue
        if i_full is None and ("全称" in t or "公司名称" in t or "主体名称" in t):
            i_full = c
        if "汇率" in t:
            if "余额" in t and i_bal is None:
                i_bal = c
            elif "损益" in t and i_pl is None:
                i_pl = c
    if not i_full or not (i_bal or i_pl):
        return None
    # 数据区：表头下一行起，连续有「主体全称」的最后一行
    r1 = hdr + 1
    r2 = r1 - 1
    r = r1
    while r <= (ws.max_row or 0):
        if not _T(ws.cell(r, i_full).value):
            break
        r2 = r
        r += 1
    if r2 < r1:
        return None
    return {"sheet": name,
            "full": get_column_letter(i_full),
            "bal": get_column_letter(i_bal) if i_bal else get_column_letter(i_pl),
            "pl": get_column_letter(i_pl) if i_pl else get_column_letter(i_bal),
            "has_bal": bool(i_bal), "has_pl": bool(i_pl),
            "r1": r1, "r2": r2}


def rate_formula(cm_ref, row, subject):
    """折算汇率列的公式：XLOOKUP 到「口径Mapping」的公司段，取该科目该用的那一列。

    找不到主体时回退 1（人民币主体本币即人民币），并在单元格批注里说明，
    不静默写 6.8454 —— 那是旧版写死的数，换客户就是错的。
    """
    if not cm_ref:
        return None
    col = cm_ref["bal"] if rate_kind_of(subject) == "bal" else cm_ref["pl"]
    sh = cm_ref["sheet"]
    return ("=IFERROR(XLOOKUP($B%d,'%s'!$%s$%d:$%s$%d,'%s'!$%s$%d:$%s$%d),1)"
            % (row, sh, cm_ref["full"], cm_ref["r1"], cm_ref["full"], cm_ref["r2"],
               sh, col, cm_ref["r1"], col, cm_ref["r2"]))


def rate_of(cm, company, subject):
    """Python 侧的同一个汇率：公式读不到值，汇总（表2/OAR）要用它。"""
    kind = rate_kind_of(subject)
    pair = ((cm or {}).get("rates") or {}).get(_T(company)) or {}
    val = pair.get(kind)
    if val is None and kind == "bal":
        val = pair.get("pl")            # 没填余额汇率时退回损益汇率
    if val is None and kind == "pl":
        val = pair.get("bal")
    return val


def _augment_prior_only(wb, prior_bkd, vc, vd, book_map):
    """上年有、本年没数的 (主体, 科目) 补一行（账面 0），保证两张表上期合计一致。

    不补的话，K 列的 SUMIFS 只能匹配到「本年存在的行」，上年那块会被漏掉：
    瑞云实测 VC 漏 39 行、VD 漏 28 行，上期合计与上年底稿差一截。
    补进来的行：本币/人民币/未审/审定都是 0，上期未审由 SUMIFS 取上年数，
    口径沿用上年（走口径Mapping）。
    """
    added = {}
    for key, rows in (("VC", vc), ("VD", vd)):
        pb = (prior_bkd or {}).get(key) or {}
        pg = pb.get("sheet")
        if not pg or pg not in wb.sheetnames:
            continue
        have = set()
        for r in rows:
            comp, _how = resolve_company(r.get("book"), book_map)
            have.add((_T(comp), _T(r.get("name"))))
        ws = wb[pg]
        try:
            last = int(pb.get("last") or 0)
        except (TypeError, ValueError):
            last = 0
        if last <= 0:
            last = ws.max_row or 0
        n = 0
        for r in range(30, last + 1):
            comp = _T(ws.cell(r, 2).value)
            nm = _T(ws.cell(r, 4).value)
            v = ws.cell(r, 7).value
            if not comp or not nm or not isinstance(v, (int, float)):
                continue
            # 上期页里本年未审数为 0 的行不补：补进来就是一行「0 / 0」，
            # 既没有分析价值，还会被当成"取数漏了"的疑点（实测 VC 38 行、VD 34 行）
            if abs(float(v)) < 0.005:
                continue
            if (comp, nm) in have:
                continue
            have.add((comp, nm))
            rows.append({"book": comp, "code": "", "name": nm,
                         "l1": "销售费用" if key == "VC" else "管理费用",
                         "dr": 0.0, "cr": 0.0, "cr_close": 0.0,
                         "amount": 0.0, "priorOnly": True})
            n += 1
        added[key] = n
    return added


def _drop_zero_rows(prior_bkd, rows, key, book_map):
    """本年金额为 0 且上年也为 0 的科目不列示。返回 (保留的行, 丢弃行数)。

    表上留一行「0 / 0」毫无信息量，还会让人怀疑是不是取数漏了；
    上期有数、本年没数的（补行）当然要留 —— 那正是变动分析的对象。
    """
    pb = (prior_bkd or {}).get(key) or {}
    vals = pb.get("values") or {}
    kept, dropped = [], 0
    for r in rows:
        try:
            amt = float(r.get("amount") or 0)
        except (TypeError, ValueError):
            amt = 0.0
        comp = _T(r.get("book"))
        if not r.get("priorOnly"):
            try:
                comp = _T(resolve_company(r.get("book"), book_map)[0]) or comp
            except Exception:
                pass
        pv = vals.get((comp, _T(r.get("name"))))
        try:
            pvf = float(pv) if pv is not None else 0.0
        except (TypeError, ValueError):
            pvf = 0.0
        if abs(amt) < 0.005 and abs(pvf) < 0.005:
            dropped += 1
            continue
        kept.append(r)
    return kept, dropped


def _add_fx_cols(ws, col):
    """在「本年账面数」后插两列：汇率、本年账面数（人民币）。

    同 `_add_py_cols` 的理由：openpyxl 的 insert_cols 不搬合并区、不搬列宽、
    不改公式，三样都得自己补。插完 col 里所有绝对列号（prior 等）要 +2。
    """
    from openpyxl.utils import get_column_letter
    from copy import copy as _cp

    at = (col.get("amount") or 5) + 1
    maxc = ws.max_column or (at + 12)
    old_w = {}
    for c in range(at, maxc + 1):
        d = ws.column_dimensions.get(get_column_letter(c))
        old_w[c] = d.width if d and d.width else None

    recs = []
    for mc in list(ws.merged_cells.ranges):
        if mc.max_col < at:
            continue
        recs.append((mc.min_row, mc.min_col, mc.max_row, mc.max_col,
                     ws.cell(mc.min_row, mc.min_col).value))
    for r1, c1, r2, c2, _v in recs:
        try:
            ws.unmerge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        except Exception:
            pass

    ws.insert_cols(at, FX_SHIFT)

    w0 = max(old_w.get(at) or 12.0, 12.0)
    for c in range(maxc + FX_SHIFT, at + FX_SHIFT - 1, -1):
        w = old_w.get(c - FX_SHIFT)
        if w:
            ws.column_dimensions[get_column_letter(c)].width = w
    for i in range(FX_SHIFT):
        ws.column_dimensions[get_column_letter(at + i)].width = w0

    for r1, c1, r2, c2, v in recs:
        nc1 = c1 + FX_SHIFT if c1 >= at else c1
        nc2 = c2 + FX_SHIFT
        try:
            ws.merge_cells(start_row=r1, start_column=nc1, end_row=r2, end_column=nc2)
            if v is not None:
                ws.cell(r1, nc1).value = v
        except Exception:
            pass

    hdr = col.get("_header_row") or 1
    # 样式要从**第 1 行**开始整列跟，不能只刷表头以下 ——
    # insert_cols 只把右边的单元格整体平移，新插的两列本身是白纸：
    # 模板顶部那几条黑/灰分区色带（封面标题带）会在这里断开，
    # 打开底稿第一眼就是「H 列黑、F/G 白」的破窗。
    try:
        for r in range(1, (ws.max_row or 0) + 1):
            src = ws.cell(r, col["amount"])
            for i in range(FX_SHIFT):
                cell = ws.cell(r, at + i)
                if src.has_style:
                    cell.font = _cp(src.font)
                    cell.fill = _cp(src.fill)
                    cell.border = _cp(src.border)
                    cell.alignment = _cp(src.alignment)
                cell.number_format = src.number_format
    except Exception:
        pass
    ws.cell(hdr, col["amount"]).value = "本年账面数（本币）"
    for i, nm in enumerate(("汇率", "本年账面数（人民币）")):
        ws.cell(hdr, at + i, nm)
    try:
        for r in range(hdr + 1, (ws.max_row or 0) + 1):
            ws.cell(r, at).number_format = "General"
            ws.cell(r, at + 1).number_format = MONEY_FMT
    except Exception:
        pass

    # 绝对列号全部右移
    for k in ("prior", "prior_unaud", "unadj", "adj"):
        if isinstance(col.get(k), int):
            col[k] = col[k] + FX_SHIFT
    return col


def _add_py_cols(ws, col):
    """在「上期审定数」前面插两列：上期未审数、上期调整数。

    为什么要真插列而不是把值挤到右边：用户要的是「未审 + 调整 = 审定」三段式，
    右边 变动金额/变动%/For X-ref/波动判断/定性/Notes/OAR 口径 六列得整体右移，
    模板里这六列的公式引用也要跟着改。

    openpyxl 的 insert_cols 只搬值和样式，**不动合并区、不动列宽、不改公式**，
    所以三样都得自己补。返回更新后的 col（prior 指向新的上期审定数列）。
    """
    from openpyxl.utils import get_column_letter

    at = col.get("prior")
    if not at:
        return col                       # 模板没有「上期审定数」列，不动

    maxc = ws.max_column or (at + 10)
    old_w = {}
    for c in range(at, maxc + 1):
        d = ws.column_dimensions.get(get_column_letter(c))
        old_w[c] = d.width if d and d.width else None

    # 合并区先记账（insert_cols 不会跟着搬）
    recs = []
    for mc in list(ws.merged_cells.ranges):
        if mc.max_col < at:
            continue
        recs.append((mc.min_row, mc.min_col, mc.max_row, mc.max_col,
                     ws.cell(mc.min_row, mc.min_col).value))

    # 插列前先把「PY 段」两个表头抓下来：它们在 K 列（原上期审定列），
    # insert_cols 会把它们推到 M，而 M 可能正好落在模板的合并区里（读出来是 MergedCell），
    # 事后再找就找不到了。
    from copy import copy as _cp1
    py_heads = []
    for _r in range(1, min(ws.max_row or 0, 40) + 1):
        _v = ws.cell(_r, at).value
        if isinstance(_v, str) and ("PY" in _v or "DATE(YEAR(" in _v.upper()):
            _src = ws.cell(_r, at)
            py_heads.append((_r, _v, _cp1(_src.font), _cp1(_src.alignment),
                             _cp1(_src.border), _src.number_format, _cp1(_src.fill)))

    # **解除合并必须在插列之前**：insert_cols 只搬格子不搬合并区，
    # 插完再按旧坐标 unmerge，解到的是"旧合并区、新内容"——openpyxl 的
    # _clean_merge_range 会把成员格的值和样式一起清成 Normal。
    # 实测踩过：模板 N28:N29 / O28:O29 是纵向合并（「本期和上期发生额」下面那两格），
    # 插列后 L29『变动金额』M29『变动%』刚好落进这两格，随后被解合并清空，
    # 表头上就出现两个没字、没底色的空格。
    for r1, c1, r2, c2, _v in recs:
        try:
            ws.unmerge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        except Exception:
            pass

    ws.insert_cols(at, PY_NEW_COLS)

    # 1) 列宽：右移的列各回原位，新两列沿用原「上期审定数」的宽度
    # 上期三列都是十几位金额（51,660,694.97），沿用模板 14.6 会挤成一片，统一放到 18
    w0 = max(old_w.get(at) or 14.6, 18.0)
    for c in range(maxc + PY_NEW_COLS, at + PY_NEW_COLS - 1, -1):
        w = old_w.get(c - PY_NEW_COLS)
        if w:
            ws.column_dimensions[get_column_letter(c)].width = w
    for i in range(PY_NEW_COLS):
        ws.column_dimensions[get_column_letter(at + i)].width = w0

    # 2) 合并区按新坐标重设（旧合并已在插列前解除）
    for r1, c1, r2, c2, v in recs:
        nc1 = c1 + PY_NEW_COLS if c1 >= at else c1
        nc2 = c2 + PY_NEW_COLS
        try:
            ws.merge_cells(start_row=r1, start_column=nc1, end_row=r2, end_column=nc2)
            if v is not None:
                ws.cell(r1, nc1).value = v
        except Exception:
            pass

    # 3) 表头：两列新名。样式从右邻格抄，保证字体/边框/底色一致
    hdr = col.get("_header_row") or 1
    for i, nm in enumerate(("上期未审数", "上期调整数")):
        c = at + i
        cell = ws.cell(hdr, c, nm)
        try:
            from copy import copy as _cp
            src = ws.cell(hdr, at + PY_NEW_COLS)
            if src.has_style:
                cell.font = _cp(src.font)
                cell.fill = _cp(src.fill)
                cell.border = _cp(src.border)
                cell.alignment = _cp(src.alignment)
        except Exception:
            pass

    # 4) 新两列在数据区的边框/字体，跟着上期审定数列走
    try:
        from copy import copy as _cp
        for r in range(hdr + 1, (ws.max_row or 0) + 1):
            src = ws.cell(r, at + PY_NEW_COLS)
            for i in range(PY_NEW_COLS):
                cell = ws.cell(r, at + i)
                if src.has_style:
                    cell.font = _cp(src.font)
                    cell.border = _cp(src.border)
                    cell.alignment = _cp(src.alignment)
                    cell.number_format = src.number_format
    except Exception:
        pass

    # 表头区（第 1 行到表头行）新插的两列是"裸格"：insert_cols 只搬已有的格子，
    # 不会给新列造样式。结果是「货币/单位」那条灰底带断成两截、PY 段右边空出两格，
    # 看着像表头被切掉一块。这里按右侧邻列（原 K 搬过去的位置）逐格复制。
    for _r in range(1, (col.get("_header_row") or 1) + 1):
        _src = ws.cell(_r, at + PY_NEW_COLS)
        for _i in range(PY_NEW_COLS):
            _cell = ws.cell(_r, at + _i)
            try:
                if _src.has_style:
                    _cell.font = _cp1(_src.font)
                    _cell.fill = _cp1(_src.fill)
                    _cell.border = _cp1(_src.border)
                    _cell.alignment = _cp1(_src.alignment)
                    _cell.number_format = _src.number_format
            except Exception:
                pass

    # PY 段表头重建到 K 列并横向合并三列：它们要盖住
    # K 上期未审 / L 上期调整 / M 上期审定，留在 M 只会盖住最后一列
    for _r, _v, _f, _al, _bd, _nf, _fl in py_heads:
        _dst = ws.cell(_r, at)
        _dst.value = _v
        try:
            _dst.font = _f
            # 对齐跟左边「本期日期」那一格走：两个日期都居中才对称
            _dst.alignment = _cp1(ws.cell(_r, col.get("amount") or 5).alignment)
            _dst.border = _bd
            _dst.number_format = _nf
            _dst.fill = _fl          # 底纹不能丢：丢了日期那格就是白底，和左边不对称
        except Exception:
            pass
        try:
            ws.merge_cells(start_row=_r, start_column=at, end_row=_r,
                           end_column=at + PY_NEW_COLS)
        except Exception:
            pass

    new = dict(col)
    new["prior"] = at + PY_NEW_COLS
    new["prior_unaud"] = at
    new["prior_adj"] = at + 1
    return new


def _write_formulas(ws, r, sum_row, col):
    """按模板模式重写行内公式（插行后原公式不会自动跟，只能统一重写）。

    列序（插完 PY 两列 + 汇率两列后，相对「本年账面数」的偏移）：
      0 本币 / 1 汇率 / 2 人民币 / 3 账表调整 / 4 未审 / 5 结构比 /
      6 审计调整 / 7 本期审定 / 8 上期未审 / 9 上期调整 / 10 上期审定 /
      11 变动额 / 12 变动率 / 13 X-ref / 14 波动判断 / 15 定性 / 16 Notes / 17 OAR 口径
    """
    from openpyxl.utils import get_column_letter
    a = col.get("amount", 5)
    L = lambda n: get_column_letter(a + n)
    rmb, adj, unadj = L(2), L(3), L(4)
    aud, cur = L(6), L(7)
    ku, ka, py = L(8), L(9), L(10)
    var, pct = L(11), L(12)
    # 上期三段式：上期未审 + 上期调整 = 上期审定；未审空则留空，不能算成 0
    ws.cell(r, a + 10).value = '=IF(%s%d="","",%s%d+%s%d)' % (ku, r, ku, r, ka, r)
    # 未审 = 人民币账面 + 账表调整（人民币列本身就是 本币 × 汇率）
    ws.cell(r, a + 4).value = "=%s%d+%s%d" % (rmb, r, adj, r)
    # 结构比 = 本行未审 / 合计行未审
    ws.cell(r, a + 5).value = '=IF($%s$%d<>0,%s%d/$%s$%d,"")' % (
        unadj, sum_row, unadj, r, unadj, sum_row)
    # 本期审定 = 未审 + 审计调整
    ws.cell(r, a + 7).value = "=%s%d+%s%d" % (unadj, r, aud, r)
    # 上期为空时变动额/变动率留空，不能算成"全是本期新增"
    ws.cell(r, a + 11).value = '=IF(%s%d="","",%s%d-%s%d)' % (py, r, cur, r, py, r)
    ws.cell(r, a + 12).value = '=IF(OR(%s%d="",%s%d=0),0,%s%d/%s%d)' % (
        py, r, py, r, var, r, py, r)
    o = a + 14
    if o <= ws.max_column:
        ws.cell(r, o).value = (
            '=IF(OR(%s%d="",%s%d=""),IF(ABS(%s%d)>=$C$22,"是（本期达阈值）","否"),'
            'IF(AND(ABS(%s%d)>=$C$22,ABS(%s%d)>=$C$23),"是","否"))'
            % (var, r, pct, r, rmb, r, var, r, pct, r))
    # 定性列是职业判断，留给人工，不写公式。


_TOTALS_WORDS = ("合计", "总计", "小计", "本页合计")


def probe_prior_sheets(data):
    """逐 sheet 探测上年费用底稿：能否定位表头、有没有上期审定数、是 VC 还是 VD。

    为什么要探：上年底稿可能把销售/管理费用分成两张表、也可能合在一张，
    表名各家写法也不同（VC.00 / 销售费用BKD / 底稿1）。**按内容判而不是只看名字**：
    科目名称列里「销售费用」多的算 VC，「管理费用」多的算 VD，两类都有算 both。

    返回 {"sheets":[...], "recVC":i, "recVD":i}，i 为下标，-1 表示没找到。
    """
    from openpyxl import load_workbook
    if not data:
        return {"sheets": [], "recVC": -1, "recVD": -1, "error": "未提供文件"}
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        return {"sheets": [], "recVC": -1, "recVD": -1, "error": "打不开：%s" % e}
    try:
        return _probe_sheets_wb(wb)
    finally:
        try:
            wb.close()
        except Exception:
            pass


def probe_prior_all(data):
    """一份文件解析一次，同时给出「上年 BKD 页」和「上一期 TOD 页」两类探测结果。

    客户把上一期 TOD 和上年费用 BKD 放在同一个底稿里（本来就是一册 WP），
    界面上不该让同一个人把同一个文件上传两次。所以由这里一次读完、
    两类结果一起回给前端，前端只留一个上传口。
    """
    from openpyxl import load_workbook
    empty_tod = {"sheets": [], "recommended": -1}
    if not data:
        return {"sheets": [], "recVC": -1, "recVD": -1, "tod": empty_tod,
                "error": "未提供文件"}
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        return {"sheets": [], "recVC": -1, "recVD": -1, "tod": empty_tod,
                "error": "打不开：%s" % e}
    try:
        base = _probe_sheets_wb(wb)
        from . import expense_tod as ETOD
        base["tod"] = ETOD.probe_tod_sheets_wb(wb)
        return base
    finally:
        try:
            wb.close()
        except Exception:
            pass


def _probe_sheets_wb(wb):
    """probe_prior_sheets 的实体：在已打开的 workbook 上做，便于一次解析多处复用。"""
    out = []
    for ws in wb.worksheets:
        item = {"sheet": ws.title, "headerRow": 0, "rows": 0, "salesRows": 0,
                "adminRows": 0, "hasPrior": False, "usable": False, "kind": ""}
        col = _locate(ws)
        if not col or not col.get("prior"):
            out.append(item)
            continue
        hdr = col.get("_header_row") or 1
        n = nvc = nvd = 0
        for r in range(hdr + 1, min(ws.max_row or 0, hdr + 3000) + 1):
            name = _T(ws.cell(r, col["name"]).value)
            v = ws.cell(r, col["prior"]).value
            if not name or not isinstance(v, (int, float)):
                continue
            if name in _TOTALS_WORDS:
                continue
            n += 1
            if "销售费用" in name:
                nvc += 1
            elif "管理费用" in name:
                nvd += 1
        nm = ws.title
        if "VC" in nm.upper() or "销售" in nm:
            kind = "vc"
        elif "VD" in nm.upper() or "管理" in nm:
            kind = "vd"
        elif nvc and nvd:
            kind = "both"
        elif nvc:
            kind = "vc"
        elif nvd:
            kind = "vd"
        else:
            kind = ""
        item.update({"headerRow": hdr, "rows": n, "salesRows": nvc, "adminRows": nvd,
                     "hasPrior": True, "usable": n > 0, "kind": kind})
        out.append(item)

    def _pick(kind, alt):
        for i, s in enumerate(out):                 # 1) 类型对得上且真有数
            if s["kind"] == kind and s["usable"]:
                return i
        for i, s in enumerate(out):                 # 2) 两类合在一张
            if s["kind"] == "both" and s["usable"]:
                return i
        for i, s in enumerate(out):                 # 3) 名字里带关键字
            if s["usable"] and alt in s["sheet"]:
                return i
        return -1

    return {"sheets": out, "recVC": _pick("vc", "销售"), "recVD": _pick("vd", "管理")}


def parse_prior_expense(data, only_sheets=None):
    """从上年费用底稿抽 (账簿, 科目名) -> 上期审定数。

    用 _locate 定位列，不猜位置。同时存一份只按科目名的兜底键：
    第二年账簿名可能改（""期初"→"期初数"），但科目名一般稳定。

    only_sheets 给定时只读这些工作表（前端两个下拉选定的 VC / VD 各一张）；
    不给则读全部能识别的表（命令行/旧行为）。
    """
    from openpyxl import load_workbook
    if not data:
        return {}, "未提供上年费用底稿"
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        return {}, "无法打开上年费用底稿：%s" % e
    want = set(_T(s) for s in only_sheets if _T(s)) if only_sheets else None
    out, n = {}, 0
    try:
        for ws in wb.worksheets:
            if want is not None and ws.title not in want:
                continue
            col = _locate(ws)
            if not col or not col.get("prior"):
                continue
            hdr = col.get("_header_row") or 1
            for r in range(hdr + 1, min(ws.max_row or 0, hdr + 3000) + 1):
                book = _T(ws.cell(r, col["book"]).value)
                name = _T(ws.cell(r, col["name"]).value)
                v = ws.cell(r, col["prior"]).value
                if not name or not isinstance(v, (int, float)):
                    continue
                if str(name).strip() in _TOTALS_WORDS:
                    continue
                pv = _r2(v)
                out[(norm_book_key(book), name)] = pv
                if ("", name) not in out:
                    out[("", name)] = pv
                n += 1
    finally:
        try:
            wb.close()
        except Exception:
            pass
    if not n:
        if want:
            return {}, "选定工作表里没找到「上期审定数」列或有数值的行：%s" % "、".join(sorted(want))
        return {}, "上年费用底稿里没找到「上期审定数」列，或该列没有数值"
    return out, ""


def aud_amount(ws, r, col, cm=None):
    """一行的「本期审定数」实体值（本币 × 汇率 + 账表调整 + 审计调整）。

    为什么不能直接读审定列：「本年账面数（人民币）」「本期未审数」「本期审定数」
    全是公式，openpyxl 读回来是字符串。表2 与 OAR 的口径汇总原来一个读本币、
    一个读人民币，导致同口径两边差一个汇率倍数（实测宣传推广费 1,731 万 vs 2,157 万）。

    V1.2.110 起「折算汇率」列也是 XLOOKUP 公式，读回来同样是字符串，
    所以要从客户 mapping（cm）按「主体 + 科目」再取一次，跟公式同源。
    """
    a = col.get("amount") or 5
    b0 = ws.cell(r, a).value
    rt = ws.cell(r, a + 1).value
    if not isinstance(rt, (int, float)) and cm:
        rt = rate_of(cm, _T(ws.cell(r, col.get("book") or 2).value),
                     _T(ws.cell(r, col.get("name") or 4).value)) or 1.0
    t = 0.0
    if isinstance(b0, (int, float)):
        t = float(b0) * (float(rt) if isinstance(rt, (int, float)) else 1.0)
    for off in (1, 4):                  # 账表调整 / 审计调整
        v = ws.cell(r, a + FX_SHIFT + off).value
        if isinstance(v, (int, float)):
            t += float(v)
    return t


def _write_variance_block(ws, col, data_start, last_data, oar_col, bucket_by_row,
                          groups=None, threshold=0.0, a3_row=0, prior_map=None,
                          prior_vals=None, cm=None):
    """在「波动和特殊性质费用说明」标题下，插一张按 OAR 口径汇总的联动表。

    为什么要它：BKD 明细动辄一两百行，审计真正要看的是「广告费今年少了 1,300 万」
    这种**口径级**波动，不是逐科目。表里每个金额都是 SUMIF 链接到明细的
    （R 列口径 + J/K 金额），改数据自动跟着变，不是硬值快照。

    「是否需调查」直接引用 C22 / C23（波动幅度、波动率），与表头上方的阈值单元格
    永远保持一致 —— 用户改 TE，这里跟着变，不用重跑。

    判定用 **OR**（金额**或**比率任一超阈值即查）：与明细行 O 列的 AND 不同，
    明细行是"两个条件都满足才标"，口径级汇总金额大、笔数多，任一超标就值得看。

    表2 的**业务描述**跟 OAR 页同一套（`expense_oar._narrate`）：从序时账摘要里
    提炼"钱花在哪类业务、单家集中度"，金额一律按**千元**给，不写"占六成"这种模糊话。
    """
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import FormulaRule

    groups = groups or {}
    try:
        from . import expense_oar as EOAR
    except ImportError:
        EOAR = None

    agg = {}
    for r, b in (bucket_by_row or {}).items():
        if not b:
            continue
        agg[b] = agg.get(b, 0.0) + aud_amount(ws, r, col, cm)
    if not agg:
        return {"status": "skip", "reason": "没有可汇总的口径"}
    buckets = sorted(agg, key=lambda k: -abs(agg[k]))
    n = len(buckets)

    # 口径 → 科目列表（_narrate 靠它去 groups 里取证据）
    subj_by, prior_by = {}, {}
    for r, b in (bucket_by_row or {}).items():
        if not b:
            continue
        nm = ws.cell(r, col["name"]).value
        if nm:
            subj_by.setdefault(b, []).append(nm)
        # 读「上期未审数」（K）而不是「上期审定数」（M）：M 是 =K+L 公式，
        # openpyxl 拿到的是字符串，isinstance 判断会全部落空 → NB 编号一个都编不出。
        # 调整表为空时 L=0，K 就等于 M，口径一致。
        _pc = col.get("prior_unaud") or col.get("prior")
        pv = ws.cell(r, _pc).value if _pc else None
        if isinstance(pv, (int, float)):
            prior_by[b] = prior_by.get(b, 0.0) + float(pv)
        else:
            # K 列此时是「SUMIFS 链接上期 BKD 页」的公式，openpyxl 读回来是字符串。
            # 回退只用「上期 BKD 页的 G 列」（与表2 D 列的 SUMIF 完全同源：
            # 同样按 B 列账套 + D 列科目名称去对）。
            # 绝不用解析上年底稿得来的「上期审定数」兜底 —— 那是更早一年的数，
            # 混进来之后导出的「上期/变动」就和表上算的不是一回事（实测差过 1,014 千）。
            _cp = _T(ws.cell(r, col["book"]).value)
            _cd = _T(ws.cell(r, col["name"]).value)
            _p2 = (prior_vals or {}).get((_cp, _cd))
            if _p2 is not None:
                prior_by[b] = prior_by.get(b, 0.0) + float(_p2)

    def _desc(b):
        """从序时账证据写一句业务描述；写不出就如实说，不编。"""
        if EOAR is None:
            return ""
        prior = prior_by.get(b)
        now = agg.get(b, 0.0)
        pct = ((now - prior) / prior) if (prior not in (None, 0)) else None
        c = {"name": b, "now": now, "prior": prior, "pct": pct,
             "chg": ((now - prior) if prior is not None else None),
             "subjects": subj_by.get(b, [])}
        try:
            note, ok = EOAR._narrate(c, groups)
        except Exception:
            note, ok = "", False
        return note if ok else ""

    trow = None
    for r in range(1, min(ws.max_row or 0, 600) + 1):
        if "波动和特殊性质费用说明" in str(ws.cell(r, 2).value or ""):
            trow = r
            break
    if trow is None:
        return {"status": "skip", "reason": "找不到「波动和特殊性质费用说明」"}

    # 小标题1 + 表头1 + 口径N + 合计1 + 空1 + 分析标题1 + 分析N + 空1
    need = 2 * n + 6
    _insert_rows_keep_layout(ws, trow + 1, need)

    J = get_column_letter(col["amount"] + FX_SHIFT + 5)      # 本期审定数
    K = get_column_letter(col["amount"] + FX_SHIFT + 8)      # 上期审定数（插两列后从 K 变 M）
    R = get_column_letter(oar_col)                # OAR 口径
    d1, d2 = data_start, last_data

    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hfont = Font(name="Arial", size=10, bold=True)
    tfont = Font(name="Arial", size=10)
    bfont = Font(name="Arial", size=10, bold=True)
    hfill = PatternFill("solid", fgColor="D9E2F3")
    sfill = PatternFill("solid", fgColor="FFF2CC")
    ctr = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    nbfont = Font(name="Arial", size=10, bold=True, color="FF0000")

    # ---- NB 编号：表2 要分析的口径各给一个号，明细行 Notes 同步标 ----
    # 判据与表1 的 G 列一致（OR：金额或比率任一超阈值）。Python 侧先算一遍，
    # 因为 openpyxl 读不到公式结果 —— 但两边必须同口径，否则「表2 说 NB1、
    # 明细行没标」或者反过来，审计抽样时对不上。
    nb_map, _seq = {}, 0
    thr_v = float(threshold or 0)
    for b in buckets:
        now_v = agg.get(b, 0.0)
        prior_v = prior_by.get(b)
        need = False
        if prior_v is not None:
            chg = now_v - prior_v
            if thr_v and abs(chg) >= thr_v:
                need = True
            elif prior_v and abs(chg / prior_v) >= 0.10:
                need = True
        if need:
            _seq += 1
            nb_map[b] = "NB%d" % _seq

    # 模板在「波动说明」下面留着一个 [A][B] 占位框（合并区 B:F），新版用表1/表2
    # 取代它了。不解除的话，表2 一长就压上去，merge_cells 抛异常 → 后面整段静默不写。
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row >= trow:
            v = str(ws.cell(mc.min_row, mc.min_col).value or "")
            if "[A]" in v or "[B]" in v:
                try:
                    ws.unmerge_cells(str(mc))
                    ws.cell(mc.min_row, mc.min_col).value = None
                except Exception:
                    pass

    # 表1 整块重画前先清边框：模板在这一带留了零散的横线/竖线，不清掉就会
    # 在表外「叉出」一根线，看着像表格没闭合。
    from openpyxl.styles.borders import Border as _Bd
    for _r in range(trow, min(trow + 120, (ws.max_row or 0) + 1) + 1):
        for _c in range(2, 10):
            ws.cell(_r, _c).border = _Bd()

    r = trow + 1
    c = ws.cell(r, 2, "表1  按 OAR 口径汇总（本期 vs 上期；金额链接上方明细，阈值取 C22 / C23）")
    c.font = Font(name="Arial", size=10, bold=True, color="C00000")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)

    r += 1
    # 注意：B 列必须保持**纯口径名**，SUMIF 的条件就是 $B{r}，加了前后缀就汇总不到。
    # 编号另起一列（H），表2 和明细行 Notes 用同一个号，三处对得上。
    # 表头到「是否需调查」为止；编号（NB1…）写在框外右侧，不占表头列
    for i, h in enumerate(["OAR 口径", "本期审定数", "上期审定数", "变动金额",
                           "变动率", "是否需调查"], start=2):
        c = ws.cell(r, i, h)
        c.font = hfont
        c.fill = hfill
        c.border = border
        c.alignment = ctr

    first = r + 1
    for k, b in enumerate(buckets):
        r = first + k
        ws.cell(r, 2, b).font = tfont
        ws.cell(r, 2).alignment = left
        # 本期：SUMIF 按口径汇总审定数
        ws.cell(r, 3, '=SUMIF($%s$%d:$%s$%d,$B%d,$%s$%d:$%s$%d)'
                % (R, d1, R, d2, r, J, d1, J, d2))
        # 上期：K 列整列没数字时留空，而不是显示 0（0 会被误读成"上年为 0"）
        ws.cell(r, 4, '=IF(COUNT($%s$%d:$%s$%d)=0,"",SUMIF($%s$%d:$%s$%d,$B%d,$%s$%d:$%s$%d))'
                % (K, d1, K, d2, R, d1, R, d2, r, K, d1, K, d2))
        ws.cell(r, 5, '=IF(OR($C%d="",$D%d=""),"",$C%d-$D%d)' % (r, r, r, r))
        ws.cell(r, 6, '=IF(OR($D%d="",$D%d=0,$E%d=""),"",$E%d/$D%d)' % (r, r, r, r, r))
        # 是否需调查：一律用嵌套 IF 逐层判空，**不用 OR/AND 包住 ABS()**。
        # Excel 的 OR/AND 不短路：上期为 0 时 F（变动率）返回 ""，OR(ABS(E)>=C22,
        # AND(F<>"",ABS(F)>=C23)) 里那个 ABS("") 仍会被算，整格直接 #VALUE!。
        ws.cell(r, 7, '=IF($E%d="","",IF(ABS($E%d)>=$C$22,"是",'
                      'IF($F%d="","否",IF(ABS($F%d)>=$C$23,"是","否"))))'
                % (r, r, r, r))
        # 数据行只画横线 + 左右外框（内部不竖分栏）：照 EY 底稿的写法，
        # 满格网线像台账，横线一行一条才像分析表
        for cc in range(2, 9):
            cell = ws.cell(r, cc)
            cell.font = tfont
            # 数据行之间不画线，只有左右外框（照 EY 底稿：一行一条太像台账）
            cell.border = Border(
                left=thin if cc == 2 else None,
                right=thin if cc == 7 else None)
            if cc in (3, 4, 5):
                cell.number_format = MONEY_FMT
                cell.alignment = right
            elif cc == 6:
                cell.number_format = "0.0%"
                cell.alignment = right
            elif cc == 7:
                cell.alignment = ctr
            elif cc == 8:
                cell.alignment = ctr
        code = nb_map.get(b)
        if code:
            c8 = ws.cell(r, 8, code)
            c8.font = nbfont
            c8.alignment = ctr

    last = first + n - 1
    r = last + 1
    ws.cell(r, 2, "合计").font = bfont
    for cc, formula in ((3, "=SUM(C%d:C%d)" % (first, last)),
                        (4, '=IF(COUNT(D%d:D%d)=0,"",SUM(D%d:D%d))' % (first, last, first, last)),
                        (5, '=IF(OR(C%d="",D%d=""),"",C%d-D%d)' % (r, r, r, r)),
                        (6, '=IF(OR(D%d="",D%d=0,E%d=""),"",E%d/D%d)' % (r, r, r, r, r))):
        ws.cell(r, cc, formula)
    for cc in range(2, 9):
        cell = ws.cell(r, cc)
        cell.font = bfont
        # 合计行不加上下边框（用户明确要求）：合计和明细同处一个框内，
        # 加线反而把表切成两截
        cell.border = Border(
            left=thin if cc == 2 else None,
            right=thin if cc == 7 else None)
        if cc in (3, 4, 5):
            cell.number_format = MONEY_FMT
            cell.alignment = right
        elif cc == 6:
            cell.number_format = "0.0%"
            cell.alignment = right

    # ---- A3 / check 两行（EY 底稿的勾稽标记）----
    # 主表 A3（明细下方那行）= 审计确认数，**留空由人填**（见上）；本表 A3 只做转引，
    # check = 本表合计 − 本表 A3。A3 没填时 check 留空，不拿 0 冒充「已勾稽」；
    # 填了才比，比不平才亮差额。
    red_f = Font(name="Arial", size=10, bold=True, color="C00000")
    for _off, _label in ((1, "A3"), (2, "check")):
        rr = r + _off
        lc = ws.cell(rr, 2, _label)
        lc.font = red_f
        lc.alignment = left
        if _label == "A3":
            # 引用**主表** A3：主表是审计确认数，可能被手工改过，表1 必须跟着它走。
            # 主表没填时这里也是空（不退回合计——退回等于自己给自己签字）。
            if a3_row:
                ws.cell(rr, 3, '=IF($J$%d="","",$J$%d)' % (a3_row, a3_row))
                ws.cell(rr, 4, '=IF($M$%d="","",$M$%d)' % (a3_row, a3_row))
        else:
            # 本期、上期各一个 check：本表合计 − 本表 A3，都应为 0
            ws.cell(rr, 3, '=IF(OR(C%d="",C%d=""),"",C%d-C%d)' % (r, rr - 1, r, rr - 1))
            ws.cell(rr, 4, '=IF(OR(D%d="",D%d=""),"",D%d-D%d)' % (r, rr - 1, r, rr - 1))
        for cc in (3, 4):
            cell = ws.cell(rr, cc)
            cell.number_format = MONEY_FMT
            cell.alignment = right
        for cc in range(2, 9):
            ws.cell(rr, cc).border = Border(
                left=thin if cc == 2 else None,
                right=thin if cc == 7 else None,
                bottom=thin if (_label == "check" and cc <= 7) else None)

    # ---- 分析段：超阈值的逐条说明（全用公式拼，数据变了文字跟着变）----
    ar = r + 5
    c = ws.cell(ar, 2, "表2  变动分析（金额 或 变动率超过波动范围的口径）")
    c.font = hfont
    ws.merge_cells(start_row=ar, start_column=2, end_row=ar, end_column=7)

    # 导出给 AI 深度分析用的证据包（口径级金额 + 明细构成 + 对手方）
    # 为什么不在生成时就把结论写死：结论是判断，程序只能给事实。
    analysis_items = []
    # AI 写过的结论：重新生成时优先用，别被默认描述顶掉
    _saved_t2 = load_table2_concl()

    for k in range(n):
        rr = ar + 1 + k
        src = first + k
        # 业务描述从序时账证据来，金额一律千元；引号/花括号要清掉，否则公式报废
        _saved = (_saved_t2.get("%s|%s" % (ws.title, buckets[k]))
                  or _saved_t2.get(buckets[k]) or "")
        desc = (_saved or _desc(buckets[k]))
        # AI 写的文字也要过一遍清洗：半角双引号会截断公式字符串，
        # 花括号会被 Excel 当成数组常量，两者都会让整格变成 #VALUE! 或文字被吞。
        desc = (desc.replace('"', "'").replace("{", "（").replace("}", "）")
                .replace("\n", " ").replace("\r", " "))
        # 编号放到文字前面（NB1：「广告费」…），框外右侧同时保留编号，
        # 和 EY 底稿一致：左侧读结论，右侧扫编号
        _code0 = nb_map.get(buckets[k])
        _pfx = (_code0 + "：") if _code0 else "·"
        f = ('=IF($G{src}="否","{pfx}——「"&$B{src}&"」未超波动范围（"'
         '&TEXT($C$22/1000,"#,##0")&" 千元 且 "&TEXT($C$23,"0.0%")&"），无需说明。",'
         '"{pfx}「"&$B{src}&"」本期 "&TEXT($C{src}/1000,"#,##0")&" 千元"'
         '&IF($F{src}="","（上期无数据或为 0，无法计算变动率）",'
         '"，上期 "&TEXT($D{src}/1000,"#,##0")&" 千元，变动 "&TEXT($E{src}/1000,"#,##0")'
         '&" 千元（"&TEXT($F{src},"0.0%")&"）")'
         '&"，超过波动范围（"&TEXT($C$22/1000,"#,##0")&" 千元 或 "&TEXT($C$23,"0.0%")'
         '&"）。"&"{desc}")').format(src=src, desc=desc, pfx=_pfx)
        c = ws.cell(rr, 2, f)
        c.font = tfont
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        if not _saved:
            # 还没 AI 结论：把任务卡挂在这一格，打开底稿就能看到"这段还没人写"
            try:
                from openpyxl.comments import Comment
                _cm = Comment(_t2_cell_note(buckets[k], ws.title, n), "PTW")
                try:
                    _cm.width, _cm.height = 460, 320
                except Exception:
                    pass
                c.comment = _cm
            except Exception:
                pass
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=7)
        # 表2 内外都要框线：每段一条底边（段间分隔 + 末段封底）、首段补顶边，
        # 左右两条竖线贯通到底 —— 只有外框的话，读的人分不清哪句对应哪个口径
        for cc in range(2, 8):
            cell = ws.cell(rr, cc)
            cell.border = Border(
                left=thin if cc == 2 else None,
                right=thin if cc == 7 else None,
                top=thin if k == 0 else None,
                bottom=thin)
        try:
            # B:G 合并宽约 90 字符，按每行 46 字估；AI 写的结论比默认描述长，
            # 固定 30 会切掉后半句（截断的文字在底稿里看不出是被切了）
            ws.row_dimensions[rr].height = max(30, min(96, 15 * (len(desc) // 46 + 2)))
        except Exception:
            pass
        code = nb_map.get(buckets[k])
        if code:
            c8 = ws.cell(rr, 8, code)
            c8.font = nbfont
            c8.alignment = ctr

        # ---- 证据包：给 AI 写变动结论用 ----
        # 只放事实（金额、构成、对手方、笔数），判断留给 AI；
        # `row` 是回填时的锚点，apply 时按它定位、再用 bucket 复核，不靠行号硬猜。
        try:
            _b = buckets[k]
            _now = float(agg.get(_b, 0.0))
            _pri = prior_by.get(_b)
            _c = {"name": _b, "now": _now, "prior": _pri,
                  "subjects": subj_by.get(_b, [])}
            _ev = EOAR._evidence_of(_c, groups, limit=8) if EOAR else []
        except Exception:
            _b, _now, _pri, _ev = buckets[k], 0.0, None, []
        analysis_items.append({
            "sheet": ws.title, "row": rr, "nb": nb_map.get(_b) or "",
            "bucket": _b,
            "now": round(_now, 2),
            "prior": (round(float(_pri), 2) if isinstance(_pri, (int, float)) else None),
            "chg": (round(_now - float(_pri), 2) if isinstance(_pri, (int, float)) else None),
            "pct": (round((_now - float(_pri)) / float(_pri), 4)
                    if isinstance(_pri, (int, float)) and float(_pri) else None),
            "subjects": subj_by.get(_b, []),
            "evidence": [{"name": _T(x.get("key")), "nature": (EOAR._nature_of(x) if EOAR else ""),
                          "amt": round(float(x.get("amt") or 0), 2),
                          "sample": _T(x.get("sample"))[:60]} for x in _ev],
        })

    # ---- 明细行 Notes 列同步标同一个编号（红色加粗）----
    # 表2 说「NB1 广告费有问题」，明细里就必须能看到 NB1，否则抽样时对不上。
    for rr, b in (bucket_by_row or {}).items():
        code = nb_map.get(b)
        if not code:
            continue
        c = ws.cell(rr, col["amount"] + FX_SHIFT + 14)          # S = Notes
        c.value = code
        c.font = nbfont

    # 需调查的整行标浅黄（条件格式，改数据自动跟着变）
    try:
        ws.conditional_formatting.add(
            "B%d:G%d" % (first, last),
            FormulaRule(formula=['$G%d="是"' % first],
                        fill=PatternFill("solid", fgColor="FFF2CC")))
    except Exception:
        pass

    return {"status": "ok", "rows": n, "buckets": buckets, "nb": nb_map,
            "table": [first, last], "analysis": [ar + 1, ar + n], "inserted": need,
            "analysisItems": analysis_items}


#: 表2 每行公式的结构是「金额/比例（活的公式） + 最后一段静态文字」。
#: 回填只替换最后那段，前面的金额与比例仍然是引用，改数据自动跟着变 ——
#: 这是「结论可以写死、数字不能写死」的分界线。
_T2_TAIL = re.compile(r'&"([^"]*)"\)$')


def _t2_concl_path():
    """AI 写过的表2 结论落在这里，下次重新生成自动带上。

    为什么必须落盘：不落的话，只要重新生成一次（改 TE、补数据、换参数），
    AI 写的整段分析就被默认描述顶掉了，项目组还以为是 AI 没写。
    """
    base = os.environ.get("KZ_DATA_DIR")
    if base:
        return os.path.join(base, "bkd_work", "table2_concl.json")
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(os.path.normpath(os.path.join(here, "..", "..", "..",
                                                      "data", "bkd_work")), "table2_concl.json")


def _ai_result_paths():
    """三份 AI 结果侧文件的路径（AI 按任务卡只写这几个，不动正本）。"""
    from . import oar_map as OM
    from . import expense_oar as EOAR
    from .tasks import result_path
    return {"bucket": result_path(OM.ai_map_path()),
            "table2": result_path(_t2_concl_path()),
            "oar": result_path(EOAR._concl_path())}


def _task_prompt(title, target, n_items, read_path, save_path, fmt, rules,
                 choices="", extra=""):
    """拼一段自包含的任务说明（照货币资金「汇率截图任务卡」的写法）。

    为什么写这么细：分发给客户的包里没有模型、也不联网，这一步天生要外部 AI 来做；
    换一台电脑的 AI 没有任何项目上下文，卡里不写清"写哪个文件、什么字段、
    写完怎么算完成"，它就只能猜。猜错的成本是复核人重做一遍。
    """
    lines = [
        "【%s · 待补任务】" % title,
        "对象：%s；本次 %d 项。" % (target, n_items),
        "现状：金额与比例已由工具算好，判断部分留空，需由使用 PTW 的 AI（或人工）完成。",
        "平台自身不调用任何模型、也不联网，只负责把事实摊平、把结论写回底稿。",
        "",
        "一、读什么",
        "  %s" % read_path,
        "  每个 item 都带说明与证据字段（金额、构成、对手方），逐条处理，不要漏。",
        "",
        "二、写什么（路径与格式都不能改）",
        "  %s" % save_path,
        "  格式：%s" % fmt,
    ]
    if choices:
        lines.append("  允许取值：%s" % choices)
    lines += [
        "",
        "三、判断要求",
    ]
    for i, r in enumerate(rules, start=1):
        lines.append("  %d) %s" % (i, r))
    lines += [
        "",
        "四、写完自检（全过才算完成）",
        "  1) 文件能被 JSON 解析；",
        "  2) 条数与待办条数一致（漏掉的必须写进 skipped 并说明原因）；",
        "  3) 取值都在允许范围内；",
        "  4) 文字里不要出现半角双引号与花括号——它们会截断底稿公式。",
        "",
        "五、收尾（这一步不做等于没做）",
        "  重新生成一次费用底稿（命令行 ptw.py exp ...，或界面点「生成」）。",
        "  程序会自动读取上面的结果文件并写进底稿，不需要手工改 Excel、也不需要再跑别的命令。",
        "",
        "  总规程见包内「文档/00_总则.md」（阶段状态机）与「文档/10_费用.md」（本册闭环），"
        "不用另找其他文件。",
        "",
        "做不到时：在该文件里加 \"skipped\": [{\"name\": \"…\", \"why\": \"…\"}]，",
        "不要编一个值填上——留空能被发现，编错会被当成人已经判断过了。",
    ]
    if extra:
        lines += ["", extra]
    return "\n".join(lines)


def _t2_cell_note(bucket, sheet, n_items):
    """表2 段落还没 AI 结论时，挂在该格批注上的任务卡。

    写在格子上而不是只写进 JSON：项目组打开底稿就要能看见"这段还没人写"，
    点一下批注就知道该写哪个文件、什么格式、写完怎么生效。
    """
    from .tasks import result_path
    p = _t2_concl_path()
    task_json = os.path.join(os.path.dirname(p), "ai_tasks", "table2_pending.json")
    res = result_path(p)
    return "\n".join([
        "【BKD 表2 变动分析 · 待补任务】",
        "口径：%s（%s）" % (bucket, sheet),
        "",
        "一、读什么",
        "  %s" % task_json,
        "  本次共 %d 个口径待写，每个 item 带金额、构成、对手方。" % n_items,
        "",
        "二、写什么（路径与格式都不能改）",
        "  %s" % res,
        '  {"items":[{"sheet":"%s","bucket":"%s","text":"结论文字"}]}' % (sheet, bucket),
        "",
        "三、怎么写",
        "  50~120 字，只写判断：变动由什么驱动、是否一次性、是否需追加程序；",
        "  不要复述金额与比例（它们在底稿里是公式）；只依据清单里的证据。",
        "  证据不足就写「需补充…」并说明缺什么，不要编业务背景。",
        "",
        "四、收尾",
        "  重新生成一次费用底稿，本段会自动换成你写的结论；不要手改 Excel。",
    ])


def _export_ai_tasks(oar_info, map_info, t2_items, customer="", balance_date="", period=""):
    """把三件 AI 待办落成固定路径的清单，并登记 tasks[]（V1.2.87）。

    返回 (tasks, warnings)。三件事都不阻断出稿：缺 AI 时底稿照样能用，
    只是文字退化成规则描述、口径留黄底待复核。
    """
    bag, warns = {}, []          # add_task 要一个 dict 容器，最后取出 tasks 列表
    try:
        from . import tasks as TK
        from . import oar_map as OM
        from . import expense_oar as EOAR
    except Exception:
        return [], warns
    data_dir = os.path.dirname(os.path.dirname(_t2_concl_path()))
    meta = {"customer": customer, "balanceDate": balance_date, "period": period}
    paths = _ai_result_paths()

    # ① OAR 口径待判定
    pending = (map_info or {}).get("pending") or []
    if pending:
        try:
            p = TK.task_file(data_dir, "oar_bucket_pending.json", {
                "task": "oar_bucket",
                "meta": meta,
                "instruction": "为每个科目判断 OAR 口径：只能取下面 choices 的值，或套用 "
                               "existingBuckets 里该客户已有的口径名。依据是该科目在序时账里"
                               "实际核算的业务实质（先看科目名核心词，再看金额规模）。",
                "existingBuckets": (map_info or {}).get("bucketSamples") or {},
                "choices": ["薪酬", "折旧摊销", "广告费", "股权激励", "其他"],
                "outputFormat": {"items": [{"company": "主体全称", "name": "科目名称",
                                            "bucket": "口径", "reason": "判断依据",
                                            "confidence": 0.9}]},
                "resultFile": paths["bucket"],
                "items": pending,
            })
        except Exception:
            p = ""
        if p:
            prompt = _task_prompt(
                "OAR 口径判定", "口径Mapping（%d 个科目上年无对应口径）" % len(pending),
                len(pending), p, paths["bucket"],
                '{"items":[{"company":"主体全称","name":"科目名称","bucket":"口径",'
                '"reason":"为什么落这个口径","confidence":0.9}]}',
                ["bucket 只能取 薪酬 / 折旧摊销 / 广告费 / 股权激励 / 其他，"
                 "或 existingBuckets 里已存在的口径名；",
                 "同科目名在不同主体下给同一口径，除非业务实质明显不同；",
                 "每项必须给 reason（一句话依据）与 confidence（0~1 或 高/中/低）。"],
                choices="、".join(["薪酬", "折旧摊销", "广告费", "股权激励", "其他"]))
            TK.add_task(bag, id="exp-oar-bucket", type="classify",
                        title="OAR 口径待判定（%d 项）" % len(pending),
                        target="口径Mapping", saveTo=paths["bucket"], prompt=prompt,
                        blocking=False, readPath=p, expectCount=len(pending),
                        afterDone="重新生成一次费用底稿，黄底行会变成「由AI判断」")
            warns.append("OAR 口径待判定 %d 项：清单 bkd_work/ai_tasks/oar_bucket_pending.json，"
                         "AI 判定后写 %s 并重新生成即生效"
                         % (len(pending), os.path.basename(paths["bucket"])))

    # ② BKD 表2 变动分析结论（只列还没 AI 结论的口径）
    _t2_saved = load_table2_concl()
    t2_todo = [it for it in t2_items
               if not (_t2_saved.get("%s|%s" % (_T(it.get("sheet")), _T(it.get("bucket"))))
                       or _t2_saved.get(_T(it.get("bucket"))))]
    if t2_todo:
        try:
            p = TK.task_file(data_dir, "table2_pending.json", {
                "task": "table2_concl",
                "meta": meta,
                "instruction": "对金额或变动率超过波动范围的口径，结合 item 里的明细构成与"
                               "对手方，写一段审计师视角的变动分析（50~120 字）：说清变动由"
                               "什么驱动、是否一次性、是否需追加程序。只写判断，不要重复数字"
                               "——数字由公式给。",
                "outputFormat": {"items": [{"sheet": "VC.00 销售费用BKD",
                                            "bucket": "广告费", "text": "结论文字"}]},
                "resultFile": paths["table2"],
                "items": t2_todo,
            })
        except Exception:
            p = ""
        if p:
            prompt = _task_prompt(
                "BKD 表2 变动分析", "两张 BKD 的「表2 变动分析」段（%d 个口径）" % len(t2_todo),
                len(t2_todo), p, paths["table2"],
                '{"items":[{"sheet":"VC.00 销售费用BKD","bucket":"广告费","text":"结论文字"}]}',
                ["只写判断（驱动因素 / 是否一次性 / 是否需追加程序），不要复述金额与比例；",
                 "依据只取 item 里的证据，证据不足就直说「需补充…」，不要编业务背景；",
                 "sheet 必须与 item 里的 sheet 一致，bucket 原样抄，不要改写。"])
            TK.add_task(bag, id="exp-table2-concl", type="write",
                        title="BKD 表2 变动分析待撰写（%d 个口径）" % len(t2_todo),
                        target="BKD 表2", saveTo=paths["table2"], prompt=prompt,
                        blocking=False, readPath=p, expectCount=len(t2_todo),
                        afterDone="重新生成一次费用底稿，表2 段落会换成 AI 写的结论")
            warns.append("BKD 表2 待分析 %d 个口径：清单 bkd_work/ai_tasks/table2_pending.json，"
                         "结果写 %s 并重新生成即生效"
                         % (len(t2_todo), os.path.basename(paths["table2"])))

    # ③ OAR 变动分析文字
    try:
        reason_items = EOAR.ai_reason_items(oar_info or {}) if oar_info else []
    except Exception:
        reason_items = []
    # 已经有 AI 文字的段落不再进清单，否则任务卡写着"待撰写"，底稿里其实已经写好了
    _oa_saved = EOAR.load_conclusions()
    reason_items = [it for it in reason_items
                    if not (_oa_saved.get("%s·%s" % (_T(it.get("label")), _T(it.get("bucket"))))
                            or _oa_saved.get(_T(it.get("bucket"))))]
    if reason_items:
        try:
            p = TK.task_file(data_dir, "oar_reason_pending.json", {
                "task": "oar_reason",
                "meta": meta,
                "instruction": "为每个 OAR 口径写一段变动原因（40~80 字）：指出由哪类业务、"
                               "哪家公司推动，是否一次性；只陈述 evidence 支持的事实，"
                               "不要重复金额与比例——它们在底稿里是公式。",
                "outputFormat": {"items": [{"sheet": "VC&VD OAR 销售费用",
                                            "bucket": "广告费", "text": "结论文字"}]},
                "resultFile": paths["oar"],
                "items": reason_items,
            })
        except Exception:
            p = ""
        if p:
            prompt = _task_prompt(
                "OAR 变动原因", "两张 OAR 页的「二、变动分析」段（%d 个口径）" % len(reason_items),
                len(reason_items), p, paths["oar"],
                '{"items":[{"sheet":"VC&VD OAR 销售费用","bucket":"广告费","text":"结论文字"}]}',
                ["只写原因与性质判断，金额、比例、是否需调查都由公式给，不要写；",
                 "evidence 里没有的事实不要补；写不出就写「需向客户了解…」并说明缺什么；",
                 "sheet 与 bucket 原样抄 item 里的值，不要改写。"])
            TK.add_task(bag, id="exp-oar-reason", type="write",
                        title="OAR 变动分析待撰写（%d 个口径）" % len(reason_items),
                        target="OAR 页", saveTo=paths["oar"], prompt=prompt,
                        blocking=False, readPath=p, expectCount=len(reason_items),
                        afterDone="重新生成一次费用底稿，分析段会换成 AI 写的文字")
            warns.append("OAR 变动分析待撰写 %d 个口径：清单 bkd_work/ai_tasks/oar_reason_pending.json，"
                         "结果写 %s 并重新生成即生效"
                         % (len(reason_items), os.path.basename(paths["oar"])))
    return bag.get("tasks") or [], warns


def load_table2_concl():
    """读回：「页|口径」→ 结论文字。

    同时合并 `table2_concl_result.json`：任务卡让外部的 AI 只写结果侧文件，
    不动正本（正本里可能有上一期人工改过的段落）。
    """
    import json
    p = _t2_concl_path()
    docs = []
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as fh:
                docs.append(json.load(fh))
        except Exception:
            pass
    try:
        from .tasks import result_path
        _rp = result_path(p)
    except Exception:
        _rp = ""
    if _rp and os.path.exists(_rp):
        try:
            with open(_rp, encoding="utf-8") as fh:
                docs.append(json.load(fh))
        except Exception:
            pass
    out = {}
    for d in docs:
        for it in (d.get("items") or []):
            txt = _T(it.get("text"))
            b = _T(it.get("bucket"))
            sh = _T(it.get("sheet"))
            if not txt or not b:
                continue
            if sh:
                out["%s|%s" % (sh, b)] = txt
            out.setdefault(b, txt)
    return out


def save_table2_concl(raw_items):
    """把 AI 结论合并落盘（同键覆盖）。"""
    import json
    p = _t2_concl_path()
    old = load_table2_concl()
    items = []
    for k, v in old.items():
        if "|" in k:
            sh, b = k.split("|", 1)
            items.append({"sheet": sh, "bucket": b, "text": v})
    idx = {(it["sheet"], it["bucket"]): it for it in items}
    for it in raw_items or []:
        b = _T(it.get("bucket"))
        txt = _T(it.get("text") or it.get("conclusion") or it.get("结论"))
        if not b or not txt:
            continue
        idx[(_T(it.get("sheet")), b)] = {"sheet": _T(it.get("sheet")), "bucket": b, "text": txt}
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as fh:
        json.dump({"items": list(idx.values())}, fh, ensure_ascii=False, indent=1)
    return {"total": len(idx), "path": p}


def _t2_key(sheet, bucket):
    """结论的匹配键：优先「页|口径」，其次只给口径名（销售/管理各对一次）。"""
    return ("%s|%s" % (sheet, bucket), bucket)


def apply_table2(data, conclusions, mark_white=True):
    """把 AI 写的变动结论回填进 BKD 表2（VC/VD 各一次）。

    conclusions 支持三种写法，越具体越优先：
        {"VC.00 销售费用BKD|广告费": "文字"}
        {"VC.00 销售费用BKD": {"广告费": "文字"}}
        {"广告费": "文字"}                 ← 两张表同名口径共用一段

    只认「表2 变动分析」标题后的那些行，且有 NB 编号（或 B 列公式以「」开头），
    不会误伤表1 与调整汇总表。
    """
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill

    concl = conclusions or {}
    if isinstance(concl, dict) and "items" in concl and isinstance(concl["items"], list):
        flat = {}
        for it in concl["items"]:
            if not isinstance(it, dict):
                continue
            txt = _T(it.get("text") or it.get("conclusion") or it.get("结论"))
            if not txt:
                continue
            b = _T(it.get("bucket") or it.get("口径"))
            sh = _T(it.get("sheet"))
            if sh and b:
                flat["%s|%s" % (sh, b)] = txt
            elif b:
                flat[b] = txt
        concl = flat

    wb = load_workbook(io.BytesIO(data))
    applied, missing = [], []
    for sname in (SHEET_VC, SHEET_VD):
        if sname not in wb.sheetnames:
            continue
        ws = wb[sname]
        for r in range(1, ws.max_row + 1):
            f = ws.cell(r, 2).value
            if not isinstance(f, str) or not f.startswith("="):
                continue
            if "超过波动范围" not in f:
                continue
            m = _T2_TAIL.search(f)
            if not m:
                continue
            # 口径名是**单元格引用**（$B224 指向表2 汇总行），不是公式里的字面量。
            # 从公式里 split「」拿到的是 `"&$B224&"`，永远匹配不上结论的键。
            m_src = re.search(r'\$B(\d+)', f)
            bucket = _T(ws.cell(int(m_src.group(1)), 2).value) if m_src else ""
            txt = None
            for k in _t2_key(sname, bucket):
                if k in concl and _T(concl[k]):
                    txt = _T(concl[k])
                    break
            if not txt:
                continue
            safe = (txt.replace('"', "'").replace("{", "（").replace("}", "）"))
            new = f[:m.start()] + '&"%s")' % safe
            ws.cell(r, 2).value = new
            if mark_white:
                for cc in range(2, 8):
                    try:
                        ws.cell(r, cc).fill = PatternFill(fill_type=None)
                    except Exception:
                        pass
            # 文字比原来长，行高按每 52 字一行估；只放大不缩小
            try:
                need = 15 * max(2, (len(safe) // 46) + 1)
                cur = ws.row_dimensions[r].height or 30
                ws.row_dimensions[r].height = max(cur, min(120, need))
            except Exception:
                pass
            applied.append({"sheet": sname, "row": r, "bucket": bucket})
    for k in concl:
        if not any(("%s|%s" % (a["sheet"], a["bucket"])) == k or a["bucket"] == k
                   for a in applied):
            missing.append(k)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), {"applied": len(applied), "rows": applied,
                            "missing": missing}


def _setup_adjust_block(ws, span=300, body_rows=20):
    """在 BKD 下方的「调整汇总表」补列头 + 画完整网格，返回可引用的数据区范围。

    模板只给了标题和一个「调整类型」，其余列头是空的 —— 项目组不知道往哪填，
    审计调整数永远对不上。这里把列头一次性铺好：

        B 调整类型 | C 调整说明 | D 简称 | E 科目名称 | F 20260930 | G 20251230 | H FOR BKD

    上面的「审计调整数」列用 SUMIFS(求和=F列, 条件1=D列简称, 条件2=H列FOR BKD) 取数，
    所以 FOR BKD 必须填成与 BKD 里「科目名称」完全一致的值（如「管理费用_股权激励」）。

    **框线必须自己画**：模板在这块只给了 B 列的左边线和末行底线，表头却有 7 列宽，
    视觉上就是「表头一条长、下面一个光框」，填进去的调整分录看不出行。
    所以这里按 表头 7 列 × body_rows 行铺满细网格。
    """
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    title_row = None
    for r in range(1, min(ws.max_row or 0, 500) + 1):
        if "调整汇总表" in str(ws.cell(r, 2).value or ""):
            title_row = r
            break
    if title_row is None:
        return None

    # 与表1/表2 连号：模板原文只写「调整汇总表」，这里补成「表3  调整汇总表」
    _tv = _T(ws.cell(title_row, 2).value)
    if _tv and not _tv.startswith("表"):
        ws.cell(title_row, 2).value = "表3  " + _tv

    head_row = title_row + 1
    # 两个金额列改成口径名，跟 BKD 的「审计调整数（本期）」「上期调整数」对上，
    # 项目组填的时候不用再想哪个日期对应哪一期
    # 简称紧跟调整类型：它是 SUMIFS 的匹配键，放在说明前面填的时候更顺
    heads = ["调整类型", "简称", "调整说明", "科目名称",
             "本期调整数", "上期调整数", "FOR BKD"]
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    f = Font(name="Arial", size=10, bold=True)
    body_f = Font(name="Arial", size=10)
    fill = PatternFill("solid", fgColor="D9E2F3")
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    money = MONEY_FMT

    for i, h in enumerate(heads, start=2):          # B 列起
        c = ws.cell(head_row, i)
        if not _T(c.value):
            c.value = h
        c.font = f
        c.fill = fill
        c.border = border
        c.alignment = ctr
        # 列宽：模板在 B~D 是窄的，调整说明/科目名称要宽一点
        L = get_column_letter(i)
        cur = ws.column_dimensions[L].width or 0
        need = {"B": 12, "C": 10, "D": 26, "E": 26, "F": 14, "G": 14, "H": 22}.get(L, 12)
        if cur < need:
            ws.column_dimensions[L].width = need

    try:
        ws.row_dimensions[head_row].height = 20
    except Exception:
        pass

    # ---- 空间不够就**整行插入**（不是插单元格）----
    # 模板在调整表下面埋了一条深蓝横带（整行，页面的分隔线）。原来的做法是在
    # 现有行上硬画 20 行网格，等于把蓝带「穿过」了表格 —— 蓝带被拦腰截断，很难看。
    # 正确做法：算出可用行数，不够就在蓝带**上面整行插入**，让蓝带整体下移，
    # 停在表格下面。
    data_start = head_row + 1
    bottom = _page_bottom(ws, data_start)
    avail = (bottom - data_start) if bottom else body_rows
    if avail < body_rows and bottom:
        # **整行插入**（不是插单元格）：分隔带及其下面的所有内容整体下移，
        # 让带子停在表格下方，而不是被表格穿过
        _insert_rows_keep_layout(ws, bottom, body_rows - avail)
        avail = body_rows

    # ---- 数据区：内部不打线，只给外框（左右两列 + 末行底线）----
    # 审计填调整分录时是一行一条，内部满格网线反而像台账；外框圈住就行。
    last = data_start + body_rows - 1
    for r in range(data_start, last + 1):
        try:
            ws.row_dimensions[r].height = 18
        except Exception:
            pass
        for i in range(2, 9):
            c = ws.cell(r, i)
            # 先把模板残留的边框清干净：模板在这一带留了一条横线，
            # 不清掉的话「内部不打线」就做不到，那条线会横穿表格中间
            c.border = Border()
            if c.font.name != "Arial" or (c.font.sz or 0) != 10:
                c.font = body_f
            if i in (6, 7):
                c.alignment = right
                c.number_format = money
            elif c.alignment.horizontal not in ("left",):
                c.alignment = left
        # 只留左右外边线
        ws.cell(r, 2).border = Border(left=thin)
        ws.cell(r, 8).border = Border(right=thin)
    # 末行底线
    for i in range(2, 9):
        cb = ws.cell(last, i).border
        ws.cell(last, i).border = Border(left=cb.left, right=cb.right,
                                         top=cb.top, bottom=thin)

    return {"title": title_row, "head": head_row,
            "data_start": data_start,
            # SUMIFS 的取值范围仍按 span（项目组真填到第 30 行也不漏取），
            # 网格只画 body_rows 行。
            "data_end": data_start + max(span, body_rows) - 1}


def _page_bottom(ws, start):
    """模板在调整表下方画了一条**横贯 B~H 的深蓝分隔带**，它标记了这页的底边。

    调整表要占的行超过它，就在**它上面整行插入**，让分隔带整体下移、停在表格
    下面 —— 否则表格会穿过这条带子，看起来被拦腰截断。

    只认「B~H 至少 5 格都是深蓝」的行：S 列整列也是深蓝（页面右侧的方向标记），
    按单元格匹配会把第 1 行当成带子，一插就是上百行，整页错位。
    """
    band = ("FF333399", "00333399", "333399")
    for r in range(start, min(ws.max_row or 0, start + 300) + 1):
        hit = 0
        for c in range(2, 9):
            try:
                if str(ws.cell(r, c).fill.fgColor.rgb or "") in band:
                    hit += 1
            except Exception:
                continue
        if hit >= 5:
            return r
    return None


def _fill_sheet(ws, rows, book_map, gaps, sheet_label, prior_map=None, map_info=None,
                groups=None, balance_date="", threshold=0.0, assurance=None,
                prior_bkd=None, cm=None, cm_ref=None):
    """把费用行写进一张 BKD 表。返回统计。

    prior_bkd 给了上期 BKD 页（{"VC": {sheet,last}, "VD": {...}}）时，
    「上期未审数」写成链接它的 SUMIFS，而不是 prior_map 里的静态数字 ——
    项目组能点过去核对，上期底稿改了重算也跟着变。
    """
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    col = _locate(ws)
    if not col:
        raise RuntimeError("%s 表头识别失败：找不到「账套 / 科目名称 / 本期账面数」三列。"
                           "请确认模板是 U_exp SWP VC&VD 底稿。" % sheet_label)

    # 上期三段式：在「上期审定数」前插「上期未审数」「上期调整数」两列。
    # 必须在写数据、找合计行之前做 —— 右边六列的偏移全部 +2。
    col = _add_py_cols(ws, col)
    # 外币主体：在「本年账面数」后插「汇率」「本年账面数（人民币）」两列。
    # 之后所有列偏移都要 +FX_SHIFT（下面统一用常量表达）。
    col = _add_fx_cols(ws, col)

    # 认定风险等级（C8:C11）：完整性 / 发生 / 计量 / 列报和披露。
    # 默认全 Minimal（EY GAM 最低档，模板 VC 页原本也是这么填的）。
    # 空着会 → D8:D11 全空 → E8=0 → C22=0 → O 列「波动幅度」判定恒为"是"，
    # 等于整张表都在"需进一步调查"，抽样门槛等于没有。所以四个格必须有值。
    _lv = [str(x or "").strip() for x in (assurance or [])] + [""] * 4
    _OK = ("Minimal", "Low", "Moderate", "High")
    for _i, _r in enumerate(range(8, 12)):
        ws.cell(_r, 3).value = _lv[_i] if _lv[_i] in _OK else "Minimal"

    hdr = col["_header_row"]
    data_start = hdr + 1
    sum_row = _find_sum_row(ws, data_start, col["amount"])
    if not sum_row:
        raise RuntimeError("%s 找不到「本期账面数」列的合计行（=SUM 公式），模板可能被改过。" % sheet_label)

    preset = max(1, sum_row - data_start)
    need = len(rows)
    n_insert = max(0, need - preset)
    if n_insert:
        # 不能用裸的 ws.insert_rows：它不搬合并区，模板的说明输入框会压进数据区
        _insert_rows_keep_layout(ws, sum_row, n_insert)
    new_sum = sum_row + n_insert
    last_data = data_start + need - 1
    # OAR 口径列 = Notes(Q) 右边一列；表头写上，样式跟明细行一起刷
    oar_col = col["amount"] + FX_SHIFT + 15
    if ws.cell(hdr, oar_col).value in (None, ""):
        ws.cell(hdr, oar_col).value = "OAR 口径"

    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    gap_fill = PatternFill("solid", fgColor=GAP_COLOR)
    txt = Font(name="Arial", size=10)
    left = Alignment(horizontal="left", vertical="center")
    right = Alignment(horizontal="right", vertical="center")

    # ---- 简称列（账套名称左边一列）：XLOOKUP 到「口径Mapping」的公司段 ----
    # 简称是审计调整 SUMIFS 的匹配键之一，先把它打通。
    short_col = max(1, col["book"] - 1)
    ws.cell(hdr, short_col).value = "简称"
    if ws.column_dimensions[get_column_letter(short_col)].width in (None, 0):
        ws.column_dimensions[get_column_letter(short_col)].width = 13

    # ---- 调整汇总表：BKD 下方的空表，补列头，I 列 SUMIFS 指过来 ----
    # 注意：这里不能先建 —— 下面「波动和特殊性质费用说明」还要插行，
    # 一插行调整表就下移，I 列公式里写死的行号会全部指错。
    # 所以放到数据写完、插行做完之后再建。
    adj = None

    mi = map_info or {}
    ol = int(mi.get("oarLast") or 0)
    cl = int(mi.get("companyLast") or 0)
    resolved = mi.get("resolved") or {}
    bucket_by_row = {}
    MAP = "口径Mapping"

    # 口径按**科目合计**判断，不按单行：同一个科目会按账簿拆成多行，
    # 逐行判断会让「中介费」这种 1304 万但在 8 个账簿各 160 万的科目
    # 全部落进「其他」——那正是要单列出来的东西。
    tot_by_name = {}
    for row in rows:
        k = row["name"]
        tot_by_name[k] = tot_by_name.get(k, 0.0) + float(row["amount"] or 0)

    for idx, row in enumerate(rows):
        r = data_start + idx
        if row.get("priorOnly"):
            # 上年有、本年没数的科目：账簿名已经折过公司全称，直接用
            company, how = _T(row.get("book")), "上期补行"
        else:
            company, how = resolve_company(row["book"], book_map)
        gap_cells = set()

        ws.cell(r, col["book"]).value = company or None
        if not company:
            gap_cells.add(col["book"])
            gaps.append({
                "sheet": sheet_label, "row": r,
                "col": col["book"], "field": "账套名称",
                "company": "", "subject": row["name"], "acct": row["book"],
                "reason": "账簿「%s」在科目余额表里找不到对应公司，请补 mapping 表" % row["book"],
            })
        if col.get("code"):
            ws.cell(r, col["code"]).value = row["code"] or None
        ws.cell(r, col["name"]).value = row["name"] or None
        ws.cell(r, col["amount"]).value = row["amount"]
        # 汇率 + 本年账面数（人民币）：人民币主体汇率为 1，公式仍保留，
        # 以后改汇率只要改这一格，未审/审定/结构比整条链跟着变。
        # 折算汇率：客户「公司mapping」按公司给了两列（余额汇率 / 损益汇率），
        # 这里按科目判定该用哪列，写成 XLOOKUP —— 客户改汇率，整表跟着变。
        # 补行（上期有、本年无）金额为 0，汇率填 1 不影响合计，也别去查 mapping。
        _comp = _T(company)
        _frm = None if row.get("priorOnly") else rate_formula(cm_ref, r, row.get("name"))
        if _frm:
            _rc = ws.cell(r, col["amount"] + 1, _frm)
        else:
            _py_rate = rate_of(cm, _comp, row.get("name"))
            if _py_rate is None:
                _py_rate = fx_rate_of(row.get("book")) if not row.get("priorOnly") else 1.0
            _rc = ws.cell(r, col["amount"] + 1, _py_rate)
        # 汇率原样显示，不四舍五入（用户口径）
        _rc.number_format = "General"
        # 不做逐行 ROUND：样本池的「折合人民币」也是不取整的乘积，
        # 两边都取整、但一边按凭证一边按科目汇总，四舍五入方向不同会累计出几分钱，
        # TOD 1.1 的 check 就会显示 -0.07 这种"看起来没勾平"的差额。
        _rmb = ws.cell(r, col["amount"] + 2, "=%s%d*%s%d" % (
            get_column_letter(col["amount"]), r, get_column_letter(col["amount"] + 1), r))
        _rmb.number_format = MONEY_FMT

        # 账表调整数：模板留空会让 G = E + F 算不出未审数，默认填 0 有调整再改
        ws.cell(r, col["amount"] + FX_SHIFT + 1).value = 0
        # 进一步调查基于定性考虑：默认「否」；需要单独判断的由项目组改
        ws.cell(r, col["amount"] + FX_SHIFT + 13).value = "否"
        # For X-ref：已被其他底稿覆盖的科目给交叉索引（应付职工薪酬20260930 之类），
        # 免得审计以为这张表漏测了薪酬。日期跟着资产负债表日走。
        # 认不出的科目直接打「/」：这列是「有没有其他底稿覆盖」的索引，
        # 没有就是没有，写「待人工确认」既占地方又不像审计底稿的写法。
        _xr = xref_of(row["name"], balance_date) or "/"
        ws.cell(r, col["amount"] + FX_SHIFT + 11).value = _xr

        # 简称：XLOOKUP 公司 mapping（A 列）
        if cl:
            ws.cell(r, short_col).value = (
                '=XLOOKUP($B%d,%s!$H$%d:$H$%d,%s!$I$%d:$I$%d,"未匹配")'
                % (r, MAP, 6, cl, MAP, 6, cl))
        # OAR 口径：XLOOKUP 口径 mapping（先按「主体全称|科目」精确，再按科目名兜底）。
        # 键必须用 B 列（账套名称=公司全称），不能用 A 列简称 —— mapping 的键是全称，
        # 拿简称去拼会全部落空。
        # 兜底文字写「见口径Mapping」而不是「未匹配-待判断」：mapping 的 C 列现在永远
        # 有值（AI 判断或规则推导），走到兜底只可能是 mapping 缺行，得让人去看表。
        if ol:
            ws.cell(r, oar_col).value = (
                '=XLOOKUP($B%d&"|"&$D%d,%s!$D$6:$D$%d,%s!$C$6:$C$%d,'
                'XLOOKUP($D%d,%s!$B$6:$B$%d,%s!$C$6:$C$%d,"见口径Mapping"))'
                % (r, r, MAP, ol, MAP, ol, r, MAP, ol, MAP, ol))
            bucket_by_row[r] = (resolved.get((company, row["name"]))
                                or oar_bucket(row["name"], tot_by_name.get(row["name"], row["amount"])))
        else:
            ws.cell(r, oar_col).value = oar_bucket(
                row["name"], tot_by_name.get(row["name"], row["amount"]))
            bucket_by_row[r] = ws.cell(r, oar_col).value

        # 审计调整数：SUMIFS 按「简称 + FOR BKD」从下方调整汇总表取本年调金额。
        # 这里只占位，实际公式在数据写完、插行做完后统一写（见函数末尾）。

        # 上期审定数：能从上年费用底稿回捞就填，捞不到才留空标黄。
        # 捞到的值是「上年审定数」，不是本年发生额——两者差一个审计调整，不能混。
        if col.get("prior"):
            pv = None
            if prior_map:
                pv = prior_map.get((norm_book_key(row["book"]), row["name"]))
                if pv is None:
                    pv = prior_map.get(("", row["name"]))
            _pu = col.get("prior_unaud") or col["prior"]
            _bk = "VC" if "VC" in sheet_label else "VD"
            _pb = (prior_bkd or {}).get(_bk) or {}
            if _pb.get("sheet") and _pb.get("last"):
                # 链接上期 BKD：按「账套名称 + 科目名称」匹配，取它的「本期未审数」G 列。
                # 按科目名称而不是科目编码：客户底稿的编码口径不一定跟本年一致
                # （有的年份改过科目编码、有的按核算维度拆行），名称是人写的、跨年稳定。
                # 就用一条 SUMIFS，不套 IFERROR / COUNTA —— 匹配不到得 0，能直接看出来。
                _s, _l = _pb["sheet"], int(_pb["last"])
                ws.cell(r, _pu).value = (
                    "=SUMIFS('%s'!$G$30:$G$%d,'%s'!$D$30:$D$%d,$D%d,"
                    "'%s'!$B$30:$B$%d,$B%d)"
                    % (_s, _l, _s, _l, r, _s, _l, r))
            elif pv is None:
                ws.cell(r, _pu).value = None
                gap_cells.add(_pu)
                gaps.append({
                    "sheet": sheet_label, "row": r,
                    "col": _pu, "field": "上期未审数",
                    "company": company, "subject": row["name"], "acct": row["book"],
                    "reason": ("未提供上年费用底稿" if not prior_map
                               else "该科目在上年费用底稿里没匹配到"),
                })
            else:
                # 上年审定数作为「上期未审数」的起点；上期调整数从下方调整表取，
                # 表里没有上年调整时它等于 0，M 就等于这里的值。
                ws.cell(r, _pu).value = pv

        _write_formulas(ws, r, new_sum, col)

        for c in range(2, min(ws.max_column, oar_col) + 1):
            cell = ws.cell(r, c)
            cell.font = txt
            cell.border = border
            # 格式与对齐统一由 _kind_of 判定，别在这里再维护一张偏移表 ——
            # 原来这里只列了 9 个偏移，变动金额/变动%（P/Q）漏在外面走了 else 分支，
            # 变成「金额右对齐、变动金额左对齐」的阴阳脸，数字格式也对不上邻居。
            kind = _kind_of(col, c)
            if kind in ("money", "pct", "rate"):
                cell.number_format = _KIND_STYLE[kind][0]
                cell.alignment = right
            else:
                cell.alignment = left

        for c in gap_cells:
            ws.cell(r, c).fill = gap_fill

    # 模板在 Notes 列留了 [A]/[B] 两个占位（原本指向下面的说明框），
    # 底稿里这两个字没用，跟着数据行一起清掉
    _notes_c = col["amount"] + FX_SHIFT + 14
    for r in range(data_start, last_data + 1):
        # 有的模板页一格写 "[A]"、有的写 "[A][B]"（两页排版不一样），一律按子串清
        _v = str(ws.cell(r, _notes_c).value or "").strip()
        if "[A]" in _v or "[B]" in _v:
            ws.cell(r, _notes_c).value = None

    # OAR 口径列按内容展开：口径名长短差很多（「预提费用（含预提和冲销）」13 个字），
    # 固定列宽会把字压没。中文按 2 个字符宽算。
    _ow = 8
    for r in range(data_start, last_data + 1):
        t = str(bucket_by_row.get(r) or "")
        if not t:
            t = str(ws.cell(r, oar_col).value or "")
            if t.startswith("="):
                t = ""
        _ow = max(_ow, sum(2 if ord(ch) > 127 else 1 for ch in t))
    ws.column_dimensions[get_column_letter(oar_col)].width = min(52, max(14, _ow + 4))

    # 上期三列最后再钉一次宽度：插列时的列宽搬移会被后面的列宽设置覆盖掉
    for _c in (col["amount"] + FX_SHIFT + 6, col["amount"] + FX_SHIFT + 7, col["amount"] + FX_SHIFT + 8):
        ws.column_dimensions[get_column_letter(_c)].width = 18

    # 合计行重写（含表头行，模板原式是 SUM(E29:E33)）
    sum_cols = range(col["amount"], col["amount"] + FX_SHIFT + 9)
    for c in sum_cols:
        L = get_column_letter(c)
        ws.cell(new_sum, c).value = "=SUM(%s%d:%s%d)" % (L, hdr, L, new_sum - 1)
    # 变动额/变动率（合计行）同样遵守「上期空则留空」
    Lc = get_column_letter(col["amount"] + FX_SHIFT + 9)      # N 变动金额
    Mc = get_column_letter(col["amount"] + FX_SHIFT + 10)     # O 变动%
    Jc = get_column_letter(col["amount"] + FX_SHIFT + 5)      # J 本期审定
    Kc = get_column_letter(col["amount"] + FX_SHIFT + 8)      # M 上期审定
    ws.cell(new_sum, col["amount"] + FX_SHIFT + 9).value = '=IF(%s%d="","",%s%d-%s%d)' % (
        Kc, new_sum, Jc, new_sum, Kc, new_sum)
    ws.cell(new_sum, col["amount"] + FX_SHIFT + 10).value = '=IF(OR(%s%d="",%s%d=0),"",%s%d/%s%d)' % (
        Kc, new_sum, Kc, new_sum, Lc, new_sum, Kc, new_sum)

    # 合计行的 O 列：模板原式同样会在 L/M 为空时 #VALUE!
    o_sum = col["amount"] + FX_SHIFT + 12
    if o_sum <= ws.max_column:
        # 合计行的「是否需调查」同理：OR/AND 不短路，M 为空时那个 ABS("") 会算成 #VALUE!
        ws.cell(new_sum, o_sum).value = (
            '=IF(%s%d="",IF(ABS(%s%d)>=$C$22,"是（本期达阈值）","否"),'
            'IF(ABS(%s%d)>=$C$22,"是",IF(%s%d="","否",IF(ABS(%s%d)>=$C$23,"是","否"))))'
            % (Mc, new_sum, get_column_letter(col["amount"]), new_sum,
               Lc, new_sum, Mc, new_sum, Mc, new_sum))

    # A3 行（原 36 行）：审计确认数。**留空，由项目组人工填**（用户 2026-11-07 口径，
    # 与货币资金 BKD 一致）。
    # 以前这里默认链到合计行（G←未审、J←本期审定、M←上期审定），本意是让 Diff 归零好看；
    # 但那等于程序替审计师签了「这三个数我确认过了」—— Diff 恒为 0，check 就永远查不出问题。
    # Diff 公式保留：人工填完自动归零，没填时显示「合计 − 空」，就是一条待办。
    # 只设数字格式方便直接输入；模板里已有的值（人工填过再生成）不动。
    a3_row = new_sum + 2
    for _c in (col["amount"] + FX_SHIFT + 2, col["amount"] + FX_SHIFT + 5, col["amount"] + FX_SHIFT + 8):
        _tgt = ws.cell(a3_row, _c)
        _tgt.number_format = MONEY_FMT
        _tgt.alignment = right

    # Rx 标记行（原 35 行）与 Diff 行（原 37 行）随插行下移，**按新列序重写**。
    # 以前这里写死了 G/J/M 三个字母（模板插列前的「未审/本期审定/上期审定」），
    # 插了两列汇率以后 G 已经是「本年账面数（人民币）」，
    # check 就变成「账面合计 − 账面A3」，永远对不上也永远查不出问题。
    _rx_row = new_sum + 1
    _off_rx = (0, 1, 2, 5)          # 人民币 / 账表调整 / 未审 / 本期审定
    for _c in range(2, col["amount"] + FX_SHIFT + 16):
        _hit = ws.cell(_rx_row, _c)
        if str(_hit.value or "").strip() == "Rx":
            _hit.value = None
    for _off in _off_rx:
        ws.cell(_rx_row, col["amount"] + FX_SHIFT + _off).value = "Rx"

    diff_row = new_sum + 3
    for _off in (2, 5, 8):          # 未审 / 本期审定 / 上期审定
        _c = col["amount"] + FX_SHIFT + _off
        _L = get_column_letter(_c)
        ws.cell(diff_row, _c).value = "=%s%d-%s%d" % (_L, new_sum, _L, new_sum + 2)

    # ---- 波动说明区：按 OAR 口径汇总 + 超阈值逐条分析（会在下方插行）----
    var_info = {}
    try:
        _pv = ((prior_bkd or {}).get("VC" if "VC" in sheet_label else "VD") or {})
        var_info = _write_variance_block(ws, col, data_start, last_data,
                                         oar_col, bucket_by_row, groups=groups,
                                         threshold=threshold, a3_row=a3_row,
                                         prior_map=prior_map,
                                         prior_vals=_pv.get("values"),
                                         cm=cm) or {}
    except Exception as e:
        var_info = {"status": "error",
                    "reason": "%s: %s" % (type(e).__name__, str(e)[:140])}

    # ---- 数据区收尾：整块金额/比例列的格式与对齐统一 ----
    # 模板自带的格子是 `_(* ... \(...\)` 写法，程序写的是 `_ * ... -...`，
    # 中文 Excel 里显示一样，但文件里混着会让「单元格格式」看着不一致；
    # 合计行、A3/check 行是从模板继承的，也要一起规范。
    try:
        for _r in range(data_start, diff_row + 1):
            for _c in range(2, oar_col + 1):
                _cell = ws.cell(_r, _c)
                _k = _kind_of(col, _c)
                if _k in ("money", "pct", "rate"):
                    _cell.number_format = _KIND_STYLE[_k][0]
                    _cell.alignment = right
    except Exception:
        pass

    # ---- 调整汇总表放在最后建（上面的插行会把它推下去）----
    adj = _setup_adjust_block(ws)
    if adj:
        _ku = get_column_letter(col["amount"] + FX_SHIFT + 6)     # K 上期未审数
        for r in range(data_start, last_data + 1):
            # 本期调整数：取调整表的「本期调整数」列（F）
            ws.cell(r, col["amount"] + FX_SHIFT + 4).value = (
                '=SUMIFS($F$%d:$F$%d,$C$%d:$C$%d,$A%d,$H$%d:$H$%d,$D%d)'
                % (adj["data_start"], adj["data_end"],
                   adj["data_start"], adj["data_end"], r,
                   adj["data_start"], adj["data_end"], r))
            # 上期调整数：取调整表的「上期调整数」列（G）。
            # 上年没有这个科目（K 空）就留空，不写 0 —— 0 会被读成"上年调整为零"。
            ws.cell(r, col["amount"] + FX_SHIFT + 7).value = (
                '=IF($%s%d="","",SUMIFS($G$%d:$G$%d,$C$%d:$C$%d,$A%d,$H$%d:$H$%d,$D%d))'
                % (_ku, r, adj["data_start"], adj["data_end"],
                   adj["data_start"], adj["data_end"], r,
                   adj["data_start"], adj["data_end"], r))

    return {"sheet": sheet_label, "rows": need, "sum": _r2(sum(r["amount"] for r in rows)),
            "sum_row": new_sum, "data_start": data_start, "last_data": last_data,
            "bucketByRow": bucket_by_row, "variance": var_info,
            "adjustRows": ([adj["data_start"], adj["data_end"]] if adj else None)}


#: 费用 BKD 各列的数字格式（按「本期账面数」列的偏移量定，不写列字母）
def _kind_of(col, c):
    amt = (col.get("amount") or 5) + FX_SHIFT
    if c == col.get("amount"):
        return "money"
    if c == col.get("amount") + 1:          # 汇率：原样显示，不四舍五入
        return "rate"
    if c == amt:                            # 本年账面数（人民币）
        return "money"
    if c in (col.get("code"),):
        return "code"
    # +6/+7/+8 是上期未审/上期调整/上期审定，+9 是变动金额；+10 是变动率
    if c in (amt + 1, amt + 2, amt + 4, amt + 5, amt + 6, amt + 7, amt + 8, amt + 9):
        return "money"
    if c in (amt + 3, amt + 10):
        return "pct"
    return "text"


_KIND_STYLE = {
    "money": (MONEY_FMT, "right"),
    "rate": ("General", "right"),
    "pct": ("0.00%", "right"),
    "code": ("@", "center"),
    "text": ("General", "left"),
}


def _insert_rows_keep_layout(ws, at, n):
    """插入 n 行，并把下方的合并区 / 行高一起搬下去。

    openpyxl 的 insert_rows **不动合并单元格**，也不搬行高。模板里
    `B40:F42` 是「波动和特殊性质费用说明」的输入框，插入 171 行之后它原地不动，
    正好压在新数据区的第 40~42 行上 —— 那三行的科目编码/名称/金额全写不进去
    （MergedCell 只允许左上角一格有值），而且**不报错**。
    实测就是这么丢了 3 行 VC 数据。

    返回搬动过的合并区数量。
    """
    if n <= 0:
        return 0
    merges, heights = [], []
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row >= at:
            merges.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
            ws.unmerge_cells(str(mc))
    for r in list(ws.row_dimensions.keys()):
        if r >= at:
            h = ws.row_dimensions[r].height
            if h:
                heights.append((r, h))

    ws.insert_rows(at, n)

    for (r1, r2, c1, c2) in merges:
        ws.merge_cells(start_row=r1 + n, start_column=c1,
                       end_row=r2 + n, end_column=c2)
    for (r, h) in heights:
        try:
            ws.row_dimensions[r + n].height = h
        except Exception:
            pass
    return len(merges)


def _normalize_bkd(ws, col, data_start, last_data, sum_row):
    """把 BKD 数据区按「行 × 列」统一刷一遍，并补齐空白高亮。

    为什么不能只刷有值的格子：插进来的新行没有行高、科目编码列是 General
    （编码会变成科学计数法）、H/M 两个比例列没有百分号格式。同一张表里
    两种字号、两套格式，一眼就能看出哪几行是程序插的。

    **不涂底色**：空白提示交给模板自带的 containsBlanks 条件格式
    （填上值自动消失）。硬涂的黄色填了值也不消失，反而看着假。
    """
    from openpyxl.styles import Alignment, Font, Border, Side
    from openpyxl.utils import get_column_letter

    if not last_data or last_data < data_start:
        return 0
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h = ws.row_dimensions[data_start].height or 15
    c1, c2 = 2, max(col.get("amount", 5) + FX_SHIFT + 8, 15)

    for r in range(data_start, last_data + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            kind = _kind_of(col, c)
            fmt, align = _KIND_STYLE[kind]
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.number_format = fmt
            cell.alignment = Alignment(horizontal=align, vertical="center")
        ws.row_dimensions[r].height = h

    # 合计行也统一（模板给的是会计格式 _(* #,##0_)，和明细行的 #,##0.00 不是一套）
    if sum_row:
        # VD 模板的合计行只有 SUM 公式、没有「合计」两个字（VC 有）——
        # 不补的话这张表的合计行在屏幕上看不出是合计，打印出来更明显。
        lbl = ws.cell(sum_row, col.get("name") or 4)
        if not _T(lbl.value):
            lbl.value = "合计"
            lbl.alignment = Alignment(horizontal="left", vertical="center")
        for c in range(c1, c2 + 1):
            cell = ws.cell(sum_row, c)
            kind = _kind_of(col, c)
            fmt, align = _KIND_STYLE[kind]
            cell.font = Font(name="Arial", size=10, bold=True)
            cell.border = border
            cell.number_format = fmt
            cell.alignment = Alignment(horizontal=align, vertical="center")

    return _ensure_blank_cf(ws, data_start, last_data, c1, c2)


def _ensure_blank_cf(ws, r1, r2, c1, c2):
    """把模板自带的「空白高亮」铺满数据区。

    模板只在几处写了 CF（VC 没有数据区的，VD 只有 B30:E32 三行）。
    插行之后数据区扩到几十行，后大半截没有提示——而这两张表的
    「哪里要补」全靠这层浅黄。颜色仍然取模板原规则（FFFFFFCC），不自造。
    """
    import copy as _copy
    from openpyxl.formatting.formatting import ConditionalFormattingList
    from openpyxl.utils import get_column_letter

    src = None
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            if rule.dxf is not None and rule.dxf.fill is not None:
                src = rule
                break
        if src is not None:
            break
    if src is None:
        return 0

    keep = []
    for cf in ws.conditional_formatting:
        rng = str(cf.sqref)
        inside = False
        try:
            m = list(cf.cells.ranges)
            # 行列都要在范围内才算「被新规则覆盖」；
            # 只比行会把 P34 这类「同区但不同列」的规则一并删掉。
            inside = all(r1 <= x.min_row and x.max_row <= r2
                         and c1 <= x.min_col and x.max_col <= c2 for x in m)
        except Exception:
            inside = False
        if not inside:
            keep.append((rng, list(cf.rules)))

    ws.conditional_formatting = ConditionalFormattingList()
    for rng, rules in keep:
        for rule in rules:
            ws.conditional_formatting.add(rng, rule)

    rule = _copy.deepcopy(src)
    rule.formula = ["LEN(TRIM(%s%d))=0" % (get_column_letter(c1), r1)]
    ws.conditional_formatting.add(
        "%s%d:%s%d" % (get_column_letter(c1), r1, get_column_letter(c2), r2), rule)
    return 1


#: 费用 BKD 认定区：B 列标签关键字 → 断言键（顺序与模板 B8:B11 一致）
BKD_ASSERTIONS = (("完整", "completeness"), ("发生", "occurrence"),
                  ("计量", "measurement"), ("列报", "presentation"))


def _lead_te_ref(wb):
    """Lead 里 TE 的绝对引用（带表名），按标签现找，找不到退回 C4。

    模板公式写死 `'Uexp.00 Lead'!$C$4`；换版模板 TE 行一挪，阈值公式就乘错格子，
    算出来的数看着正常其实是错的。
    """
    from openpyxl.utils import get_column_letter
    lead = wb["Uexp.00 Lead"] if "Uexp.00 Lead" in wb.sheetnames else None
    if lead is not None:
        for r in range(1, min(lead.max_row or 0, 14) + 1):
            for c in range(1, 5):
                t = _T(lead.cell(r, c).value).replace(" ", "")
                if t and "可容忍误差" in t:
                    for cc in range(c + 1, c + 5):
                        if lead.cell(r, cc).value not in (None, ""):
                            return "'Uexp.00 Lead'!$%s$%d" % (get_column_letter(cc), r)
    return "'Uexp.00 Lead'!$C$4"


def _fill_bkd_cra(wb, pcts=None, grades=None):
    """把 CRA 等级与百分比写进 VC/VD 两张 BKD 的认定区。

    模板 D8 是 `=IF(C8="Minimal",'Uexp.00 Lead'!$C$4*50%,IF(C8="Low",...))`，
    百分比固化在公式里。这里按平台配置重新生成同一串 IF，只换百分比数字 ——
    改配置、重新生成，阈值跟着变，不需要任何中间表。E8 的 `=MIN($D$8:$D$11)` 不动。

    pcts 缺省取 baseinfo 里「销管费用」的默认值（50/25/15/10，与模板一致）。
    """
    try:
        from . import baseinfo as BI
    except ImportError:
        import baseinfo as BI
    from openpyxl.styles import Alignment
    pcts = dict(BI.default_pcts("销管费用")) if not pcts else dict(pcts)
    grades = grades or {}

    def _pct(k, dflt):
        v = pcts.get(k)
        return "%.6g%%" % (float(dflt if v is None else v) * 100)

    te_ref = _lead_te_ref(wb)
    done = {}
    for sheet in (SHEET_VC, SHEET_VD):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        rows = {}
        for r in range(6, min(ws.max_row or 0, 20) + 1):
            t = _T(ws.cell(r, 2).value).replace(" ", "")
            for kw, key in BKD_ASSERTIONS:
                if key not in rows and kw in t:
                    rows[key] = r
                    break
        if not rows:
            continue
        for key, r in rows.items():
            if grades.get(key):
                c = ws.cell(r, 3, grades[key])
                c.alignment = Alignment(horizontal="center", vertical="center")
            f = ('=IF($C{r}="","",IF($C{r}="N/A","-",'
                 'IF($C{r}="Minimal",{te}*{m},IF($C{r}="Low",{te}*{l},'
                 'IF($C{r}="Moderate",{te}*{mo},IF($C{r}="High",{te}*{h},""))))))').format(
                     r=r, te=te_ref,
                     m=_pct("Minimal", 0.50), l=_pct("Low", 0.25),
                     mo=_pct("Moderate", 0.15), h=_pct("High", 0.10))
            cell = ws.cell(r, 4, f)
            cell.alignment = Alignment(horizontal="right", vertical="center")
        done[sheet] = rows
    return {"rows": done, "teRef": te_ref, "pcts": pcts}


def _write_te(wb, te=None, sad=None, standard="", currency=""):
    """把基础信息写进 Lead：C4 TE / C5 SAD / C6 会计准则 / C7 记账本位币。

    不写 TE 会静默毁掉整个抽样判断：C22 = E8 = TE×50%，TE 为 0 时
    「进一步调查」的公式 `ABS(L)>=0` 恒真，176 行全部标「是」。
    抽出来是为了让测试能和主流程跑同一段代码，而不是各写一遍。

    C6/C7 模板是空的：项目组每次都手填，AI 也猜不出客户用哪套准则、什么本位币。
    平台让填就写进来；没填就保持空白（不编一个默认准则 —— 猜错比空着更糟）。
    """
    te_val, sad_val = te, sad
    try:
        lead_ws = wb["Uexp.00 Lead"]
        if not isinstance(te_val, (int, float)):
            v = lead_ws.cell(4, 3).value
            te_val = float(v) if isinstance(v, (int, float)) else None
        if not isinstance(te_val, (int, float)):
            te_val = 550000.0          # 客户指定值；用默认时在 info 里标注
        if not isinstance(sad_val, (int, float)):
            sad_val = round(te_val * 0.1, 2)
        for _r, _v in ((4, te_val), (5, sad_val)):
            c = lead_ws.cell(_r, 3)
            c.value = _v
            c.number_format = "#,##0.00"
        for _r, _v in ((6, standard), (7, currency)):
            _v = _T(_v)
            if _v:
                lead_ws.cell(_r, 3).value = _v
    except Exception:
        pass
    return te_val, sad_val


def _fix_lead_links(wb, sheets_info, customer="", balance_date=""):
    """重写 Uexp.00 Lead 的链接公式，指向 BKD 实际的合计行。

    模板里写的是 `'VC.00 销售费用BKD'!E34` —— E34 是「插行前」的合计行。
    一旦某个科目有 5 行明细，合计行挪到 35，Lead 就静默指向明细行了，
    而且不会报错（数还是个数，只是少算了最后一行）。

    顺带把客户名称写到 **C2**：模板 B2 是「客户名称」标签，值在 C2；
    原来代码写的是 D2，而 D2:K7 是说明文字的合并区，永远写不进去（静默无效）。
    """
    if "Uexp.00 Lead" not in wb.sheetnames:
        return {}
    lead = wb["Uexp.00 Lead"]
    by_sheet = {st["sheet"]: st for st in sheets_info}
    done = {}

    # 表1：B 列科目名 / C 列索引号里含工作表名
    hdr = None
    for r in range(1, min(lead.max_row, 60) + 1):
        if "科目名称" in _T(lead.cell(r, 2).value):
            hdr = r
            break
    if hdr is None:
        return {"error": "Lead 里找不到「科目名称」表头行，未改链接"}

    for r in range(hdr + 1, min(lead.max_row, hdr + 12) + 1):
        nm = _T(lead.cell(r, 2).value)
        idx = _T(lead.cell(r, 3).value)
        if not nm and not idx:
            continue
        if nm in ("合计", "总计", "小计"):
            break
        target = None
        for sheet, st in by_sheet.items():
            if sheet in idx or (nm and nm in sheet):
                target = (sheet, st)
                break
        if not target:
            continue
        sheet, st = target
        sum_row = st["sum_row"]
        # D 本期账面数 ← E 合计 / E 账表调整 ← F / G 审计调整 ← I / I 上期审定 ← K
        for lead_c, bkd_c in ((4, "E"), (5, "F"), (7, "I"), (9, "M")):
            cell = lead.cell(r, lead_c, "='%s'!%s%d" % (sheet, bkd_c, sum_row))
            cell.number_format = MONEY_FMT
        done[nm] = {"sheet": sheet, "sumRow": sum_row}

    if customer and not _T(lead.cell(2, 3).value):
        lead.cell(2, 3, customer)
    # 日期以**命令行/界面传进来的**为准：模板里那个 2026-12-31 是上一年的默认值，
    # 不是用户填的，照抄会让整份底稿的口径都错（OAR 的期间也取这里）。
    if balance_date:
        lead.cell(3, 3, balance_date)
        lead.cell(3, 3).number_format = "yyyy-mm-dd"
    return done


def build_expense_workbook(conn, tpl_dir, period_from="1", period_to="6",
                           balance_date=None, customer="", extra_map=None,
                           prior_map=None, pre_end=None, oar_evidence=None,
                           oar_conclusions=None, oar_chg_min=None,
                           oar_chg_pct=0.10, te=None, sad=None,
                           standard="", currency="", assurance=None,
                           prior_oar=None, prior_shorts=None, prior_note="",
                           cra=None, prior_tod=None, prior_tod_sheet=None):
    """生成费用底稿。返回 (bytes, info dict) 或 (None, {"error": ...})。"""
    from openpyxl import load_workbook

    tpl = _find_tpl(tpl_dir)
    if not tpl:
        return None, {"error": "找不到费用模板：%s" % EXP_TPL_NAME}

    rows, rmeta = read_expense(conn, period_from, period_to)
    if rmeta.get("error"):
        return None, rmeta
    if not rows:
        return None, {"error": "序时账里没有取到「销售费用 / 管理费用」的发生额。"
                               "请确认已合并序时账，且期间填得对（当前 %s~%s）。" % (period_from, period_to)}

    book_map = read_book_map(conn, extra=extra_map)

    vc, vd, other = split_vc_vd(rows)
    gaps, stats, mapping_used = [], [], []

    wb = load_workbook(tpl)

    # 客户 PBC「公司mapping」只读一次：简称、上期主体折算、汇率列都用它
    cm = load_company_map(conn)

    # ---- 口径 Mapping：必须先建表，BKD 的 XLOOKUP 才有地方指 ----
    map_info = {}
    cm_ref = None
    try:
        from . import oar_map as OM
        map_rows = []
        for r in rows:
            comp, _how = resolve_company(r["book"], book_map)
            map_rows.append({"company": comp, "name": r["name"], "amount": r["amount"]})
        ai_path = OM.ai_map_path()
        # 客户 PBC 里的「公司mapping」优先，上年底稿的简称只作兜底
        _shorts_pbc = cm.get("shorts") or {}
        # 汇率两列也来自同一张「公司mapping」：先把每个主体补齐（客户没填的，
        # 外币主体按 FX_RATE、人民币主体按 1），补完写回 cm["rates"]，
        # 保证「口径Mapping 公司段写出去的值」与「Python 侧汇总用的值」是同一份。
        _names = []
        for _r in map_rows:
            _c = _T(_r.get("company"))
            if _c and _c not in _names:
                _names.append(_c)
        for _c in (prior_shorts or {}):
            if _c and _c not in _names:
                _names.append(_c)
        _missing = []
        _rates = cm.setdefault("rates", {})
        for _c in _names:
            _pair = _rates.setdefault(_c, {})
            if _pair.get("pl") is None:
                _pair["pl"] = (FX_RATE if any(h.upper() in _c.upper() for h in FX_BOOK_HINTS)
                               else 1.0)
                _missing.append(_c)
            if _pair.get("bal") is None:
                _pair["bal"] = _pair["pl"]

        map_info = OM.build(wb, prior_oar or {}, prior_shorts or {}, map_rows,
                            OM.load_ai_map_all(ai_path), book_map=book_map,
                            company_shorts=_shorts_pbc, company_rates=_rates)
        map_info["companyShorts"] = len(_shorts_pbc)
        # 公司段（简称 + 余额/损益汇率）就在口径Mapping 里，不另建 sheet；
        # BKD 的折算汇率列、两个样本池的折算率列都 XLOOKUP 指向这一段。
        cm_ref = find_company_map_ref(wb)
        map_info["companyMapSheet"] = cm_ref.get("sheet") if cm_ref else ""
        map_info["rateMissing"] = sorted(_missing)
        map_info["rateCols"] = ({"bal": cm_ref.get("bal"), "pl": cm_ref.get("pl")}
                                if cm_ref else {})
        if prior_note:
            map_info["priorNote"] = prior_note
    except Exception as e:
        map_info = {"status": "error",
                    "reason": "%s: %s" % (type(e).__name__, str(e)[:160])}

    # ---- TE / SAD 先写进 Lead（必须早于 _fill_sheet）----
    # 模板里 C4/C5 是空的，不写的话 C22（波动幅度）= TE×50% 全是 0，
    # O 列公式 ABS(L)>=0 恒真，整张 BKD 的「进一步调查」全是"是"，等于没判；
    # 截止性门槛、BKD 表2 的「是否需调查」与 NB 编号都取 C22，同样依赖它。
    te_val, sad_val = _write_te(wb, te, sad, standard, currency)

    # ---- CRA 认定风险：百分比写进两张 BKD 的 D8:D11 ----
    # 模板原公式把 50/25/15/10 写死在 IF 里，改风险偏好只能逐格改公式；
    # 这里按平台配置重生成同一串 IF，改配置→重新生成即生效。
    cra_info = {}
    try:
        cra_info = _fill_bkd_cra(wb, (cra or {}).get("pcts"), (cra or {}).get("assertions"))
    except Exception as e:
        cra_info = {"error": "%s: %s" % (type(e).__name__, str(e)[:140])}

    # ---- 序时账证据分组：OAR 页和 BKD 表2 的变动分析共用这一份 ----
    # 提到 fill 之前，是因为表2 的「从序时账找原因」要在写 BKD 时就用。
    _groups = {}
    try:
        from . import expense_oar as _EOAR0
        _groups = _EOAR0.group_account(conn, period_from, period_to)
    except Exception:
        _groups = {}

    # ---- 上一期 BKD：整页搬入（只搬值、保留格式），本期「上期未审数」改为链接它 ----
    # 必须在写 BKD 明细之前搬：K 列要引用它的 sheet 名与数据区行号。
    prior_bkd = {}
    if prior_tod:
        try:
            from . import expense_tod as _ETOD0
            # 上期页的「账套」是账簿名，本年写公司全称；搬进来时折一次，
            # 之后 K 列的 SUMIFS 才能按「主体 + 科目名称」对上。
            _pbc = cm

            def _resolve_prior(b, _bm=book_map, _pb=_pbc.get("books") or {},
                               _pc=_pbc.get("compacts") or {}):
                """上期账簿名 → 公司全称：本年账簿映射 → 客户 mapping 的 2025/2026
                账簿名 → 客户简称 → 紧凑键（去地域/公司性质）。都没有就留空，不猜。"""
                c, _how = resolve_company(b, _bm)
                if c:
                    return c
                k = _T(b)
                for got in (_pb.get(k), _pb.get(norm_book_key(k)),
                            _pc.get(compact_name(k))):
                    if got:
                        return got
                return ""

            prior_bkd = _ETOD0.copy_prior_bkd_as_sheet(
                wb, prior_tod, resolve=_resolve_prior)
        except Exception as e:
            prior_bkd = {"error": "%s: %s" % (type(e).__name__, str(e)[:140])}

    # 本年 0 且上年 0 的科目先剔掉（表上留一行 0 只会让人怀疑漏数）
    zero_dropped = {"VC": 0, "VD": 0}
    if prior_bkd and not prior_bkd.get("error"):
        vc, zero_dropped["VC"] = _drop_zero_rows(prior_bkd, vc, "VC", book_map)
        vd, zero_dropped["VD"] = _drop_zero_rows(prior_bkd, vd, "VD", book_map)

    # 上年有、本年没数的科目补行：不补，上期合计就对不上上年底稿
    prior_added = {}
    if prior_bkd and not prior_bkd.get("error"):
        try:
            prior_added = _augment_prior_only(wb, prior_bkd, vc, vd, book_map)
        except Exception as e:
            prior_added = {"error": "%s: %s" % (type(e).__name__, str(e)[:120])}

    for sheet, data in ((SHEET_VC, vc), (SHEET_VD, vd)):
        if sheet not in wb.sheetnames:
            return None, {"error": "模板里找不到工作表「%s」。当前工作表：%s"
                                   % (sheet, "、".join(wb.sheetnames))}
        if not data:
            continue
        st = _fill_sheet(wb[sheet], data, book_map, gaps, sheet,
                         prior_map=prior_map, map_info=map_info, groups=_groups,
                         balance_date=balance_date or "",
                         threshold=(float(te_val) * 0.5) if te_val else 0.0,
                         assurance=assurance, prior_bkd=prior_bkd,
                         cm=cm, cm_ref=cm_ref)
        stats.append(st)

    # 账套映射留痕（决策 2：B 列填公司名，但账簿要可追溯）
    for r in rows:
        company, how = resolve_company(r["book"], book_map)
        mapping_used.append({"book": r["book"], "company": company, "how": how})

    # 日期（不覆盖用户已填的内容，空了才写）
    if balance_date:
        for sheet in (SHEET_VC, SHEET_VD):
            if sheet in wb.sheetnames:
                ws = wb[sheet]
                if not _T(ws.cell(28, 5).value):
                    ws.cell(28, 5).value = balance_date

    # ---- 行格式统一 + 空白高亮（数据区按「行 × 列」全量刷）----
    for st in stats:
        if st["sheet"] in wb.sheetnames:
            _normalize_bkd(wb[st["sheet"]], _locate(wb[st["sheet"]]),
                           st["data_start"], st.get("last_data"), st["sum_row"])

    # ---- Lead 链接改指真实合计行（模板写死 E34，插行后会静默少算一行）----
    lead_info = _fix_lead_links(wb, stats, customer=customer, balance_date=balance_date or "")

    # ---- OAR 独立生成（不复用上年模板）----
    oar_info = {}
    try:
        from . import expense_oar as EOAR
    except ImportError:
        EOAR = None
    if EOAR is not None:
        try:
            # 证据：序时账按「对方单位 / 业务类型」分组。工具只摆证据，不下结论。
            groups = _groups or EOAR.group_account(conn, period_from, period_to)
            oar_info = EOAR.build_oar(
                wb, stats, meta={"customer": customer, "balance_date": balance_date or "",
                                 "period": rmeta.get("period")},
                groups=groups, te=te_val, sad=sad_val,
                conclusions=oar_conclusions or None,
                chg_amount_min=oar_chg_min, chg_pct_min=oar_chg_pct,
                prior_bkd=prior_bkd, cm=cm)
            oar_info["evidenceAccounts"] = len(groups)
            # 证据导出：AI 读它写结论，再 oar-fill 回填。不导出就没法让 AI 参与。
            if oar_evidence:
                try:
                    EOAR.export_evidence(
                        oar_info,
                        {"customer": customer, "balance_date": balance_date or "",
                         "period": rmeta.get("period")},
                        oar_evidence)
                    oar_info["evidenceFile"] = oar_evidence
                except Exception as e:
                    oar_info["evidenceError"] = "%s: %s" % (type(e).__name__, str(e)[:120])
        except Exception as e:
            oar_info = {"status": "error", "reason": "%s: %s" % (type(e).__name__, str(e)[:160])}

    # ---- 截止性测试：样本池-cut off + VC&VD.01.4 ----
    # 门槛跟 BKD 的波动幅度（C22 = TE×50%）保持一致，改 TE 两处一起动。
    cutoff_info = {}
    try:
        from . import expense_cutoff as ECUT
        thr = (float(te_val) * 0.5) if te_val else 50000.0
        # 工资薪酬桶按 OAR 口径判定，与 BKD 的「OAR 口径」列同一批科目。
        # 键用 (主体, 科目名)；BKD 的主体是「账簿名→公司全称」，序时账写的是
        # company_name，两边不一定字面相同，所以再补一层 "同名科目口径唯一" 的
        # 兜底键 ("", 科目名) —— 同一科目在多个主体下口径打架时不兜底，宁可退回关键词。
        _oar_map = {}
        try:
            _res = (map_info or {}).get("resolved") or {}
            _by_name, _amb = {}, set()
            for (_c, _n), _b in _res.items():
                if _n in _by_name and _by_name[_n] != _b:
                    _amb.add(_n)
                _by_name[_n] = _b
            for _n in _amb:
                _by_name.pop(_n, None)
            _oar_map = {(c, n): b for (c, n), b in _res.items()}
            for _n, _b in _by_name.items():
                _oar_map[("", _n)] = _b
        except Exception:
            _oar_map = {}
        cutoff_info = ECUT.build(wb, conn, balance_date or "", threshold=thr,
                                 stats=stats, tpl_dir=tpl_dir,
                                 pre_end=pre_end, period_from=period_from,
                                 prior_tod_bytes=prior_tod,
                                 prior_tod_sheet=prior_tod_sheet,
                                 oar_map=_oar_map)
    except Exception as e:
        cutoff_info = {"status": "error",
                       "reason": "%s: %s" % (type(e).__name__, str(e)[:160])}

    # 补外围蓝框：右侧整列 + 底部整行，让底稿被包住（模板只铺到固定行，明细一多就断）
    for _sn in (SHEET_VC, SHEET_VD):
        if _sn in wb.sheetnames:
            try:
                _paint_frame(wb[_sn])
            except Exception:
                pass

    # 表2 的待分析清单（两张 BKD 合并）：金额/比例是事实，结论等 AI 写。
    # 走 `ptw.py exp --ai-out <json>` 导出，`ptw.py ai-fill --file <json>` 回填。
    t2_items = []
    for _st in stats:
        for _it in ((_st.get("variance") or {}).get("analysisItems") or []):
            _it["sheet"] = _st.get("sheet") or _it.get("sheet")
            t2_items.append(_it)

    # ---- AI 待办：三份清单落固定路径 + tasks[]（V1.2.85）----
    # 平台不调模型，这三件事天生要外部 AI 做；不把"读哪、写哪、怎么算完成"钉成
    # 自包含的任务卡，换一台电脑的 AI 就完全接不上。
    ai_tasks, ai_warns = [], []
    try:
        ai_tasks, ai_warns = _export_ai_tasks(
            oar_info, map_info, t2_items,
            customer=customer, balance_date=balance_date or "",
            period=rmeta.get("period"))
    except Exception as e:
        ai_warns = ["AI 待办清单导出失败：%s: %s" % (type(e).__name__, str(e)[:120])]

    _append_check_sheet(wb, stats, gaps, mapping_used, rmeta, balance_date, other)

    # ---- 呈现顺序：统一在落盘前排一次（OAR 最前、口径Mapping 紧跟、池子跟各自页）----
    sheet_order = _reorder_sheets(wb)
    # ---- 隐藏纯说明页（详细测试 TOD 只有标题，不填数）----
    hidden_sheets = _hide_aux_sheets(wb)

    bio = io.BytesIO()
    wb.save(bio)
    seg = read_expense_by_segment(conn, period_from, period_to,
                                  pre_end or period_to)
    info = {
        "tpl": os.path.basename(tpl), "vc": len(vc), "vd": len(vd), "other": len(other),
        # 分段金额：预审段 / 年审段。BKD 不按段拆，这里只供 TOD 池与报告引用
        "bySegment": ({k: seg.get(k) for k in ("preEnd", "pre", "final", "total")}
                      if not seg.get("error") else {"error": seg.get("error")}),
        # Lead 链接实际指向的合计行（换模板/插行后第一眼看这里）
        "leadLinks": lead_info,
        # OAR 页生成情况
        "oar": oar_info,
        # 截止性测试：样本池口径统计 + 表3 落点
        "cutoff": cutoff_info,
        # 口径 Mapping：几条、多少待判断（明细走 oarmap-draft 导出）
        # resolved / pending 是内部用的（key 是「主体,科目」元组，JSON 序列化会直接
        # TypeError: keys must be str），别往外传 —— 留着会让整个接口 400。
        "oarMap": {k: v for k, v in (map_info or {}).items() if k != "resolved"},
        # 客户公司mapping 里没填汇率的主体（公式会回退 1，外币主体会算错）
        "rateMissing": (map_info or {}).get("rateMissing") or [],
        "amount": rmeta["sum"], "period": rmeta["period"],
        # 表2 待 AI 分析的口径清单（金额/构成/对手方），供导出给平台 AI 写结论
        "table2": t2_items,
        # AI 待办（口径判定 / 表2 结论 / OAR 分析）：结构化，调用方可机器核对结清没有
        "tasks": ai_tasks,
        "aiTasks": [t.get("id") for t in ai_tasks],
        "sheets": stats, "gaps": len(gaps), "gapList": gaps[:300],
        # 最终 sheet 顺序（排完序的真实结果，排查「表跑哪去了」第一眼看这里）
        "sheetOrder": sheet_order,
        # 被隐藏的说明页（不是丢了，是 sheet_state=hidden）
        "hiddenSheets": hidden_sheets,
        # CRA：写进 BKD 认定区的百分比与落点行
        "cra": cra_info,
        "needsManual": _field_count(gaps),
        "vc_sum": _r2(sum(r["amount"] for r in vc)),
        "vd_sum": _r2(sum(r["amount"] for r in vd)),
        "closeOut": rmeta.get("close_out", 0),
        "priorLoaded": bool(prior_map),
        # 上年有、本年没数而补进来的行数（VC / VD）
        "priorOnlyAdded": prior_added,
        # 本年 0 且上年 0、按口径不列示而剔掉的行数
        "zeroRowsDropped": zero_dropped,
        "priorMatched": (sum(1 for r in rows
                             if prior_map and (
                                 prior_map.get((norm_book_key(r["book"]), r["name"])) is not None
                                 or prior_map.get(("", r["name"])) is not None))
                         if prior_map else 0),
        # 待办说明（给人看）：每条都指向清单文件与下一步，别让人猜
        "warnings": ai_warns,
    }
    return bio.getvalue(), info


def _field_count(gaps):
    out = {}
    for g in gaps:
        k = g.get("field", "")
        out[k] = out.get(k, 0) + 1
    return out


def _paint_frame(ws, col=None):
    """把最外侧的蓝色装饰补齐：右侧整列 + 底部**一行**，让底稿被蓝框包住。

    模板在固定行铺了一行深蓝横带、右侧一列深蓝；明细一多、表3 一插行就断，
    而且会多出一条 —— 所以先把横贯的蓝带全部清掉，再只在内容末行下面补一行。
    蓝条列不写死：插了两列之后它在 OAR 口径右边，按表头找。
    """
    from openpyxl.styles import PatternFill

    band = PatternFill("solid", fgColor="333399")

    def _is_band(c):
        try:
            return str(c.fill.fgColor.rgb or "").endswith("333399")
        except Exception:
            return False

    if not col:
        col = 19
        for r in range(1, min(ws.max_row or 0, 60) + 1):
            hit = None
            for c in range(1, (ws.max_column or 20) + 1):
                if str(ws.cell(r, c).value or "").strip() == "OAR 口径":
                    hit = c + 1
                    break
            if hit:
                col = hit
                break

    last = ws.max_row or 1

    # 内容末行：底边要贴着它，不能贴着 max_row —— 模板留了一堆空白格式行，
    # 按 max_row 画会离内容几十行远，看着像悬在空中
    content_last = 1
    for r in range(last, 0, -1):
        hit = False
        for c in range(1, col):
            cell = ws.cell(r, c)
            if cell.value not in (None, ""):
                hit = True
                break
            # 空表也要算「有内容」：调整汇总表整片数据区是空的，
            # 只认 value 会把底边画在它第一行上，直接横穿表格
            b = cell.border
            if any(getattr(x, "style", None) for x in (b.left, b.right, b.top, b.bottom)):
                hit = True
                break
        if hit:
            content_last = r
            break
    bottom = content_last + 1

    # 1) 抹掉所有「横贯」的蓝带（≥6 格同色），只留右侧竖条
    for r in range(1, last + 1):
        n = sum(1 for c in range(1, col) if _is_band(ws.cell(r, c)))
        if n >= 6:
            for c in range(1, col + 1):
                ws.cell(r, c).fill = PatternFill(fill_type=None)

    # 2) 右侧竖条：只画到内容底边，再往下的空白不补（否则尾巴拖几十行）
    for r in range(1, bottom + 1):
        cell = ws.cell(r, col)
        if not cell.fill.patternType:
            cell.fill = band
    for r in range(bottom + 1, last + 1):
        cell = ws.cell(r, col)
        if _is_band(cell):
            cell.fill = PatternFill(fill_type=None)

    # 3) 底部只补一行，紧贴最后一个有内容的行
    for c in range(1, col + 1):
        ws.cell(bottom, c).fill = band


#: 生成后的 sheet 呈现顺序（用户定的，不要按代码创建顺序走）。
#: OAR 最前 → 口径Mapping 紧跟 → 三张底稿 → TOD 两页 → TOD 池 → 法律费用 →
#: 01.4 → 截止性池。名单外的（Skywind / DS_INTERNAL_* / 费用检查）保持原相对顺序排最后。
SHEET_ORDER = [
    # OAR 一页拆两页（V1.2.79）：销售费用、管理费用各一张，紧跟在前面成对出现
    "VC&VD OAR 销售费用",
    "VC&VD OAR 管理费用",
    "口径Mapping",

    "汇总",
    "Uexp.00 Lead",
    "VC.00 销售费用BKD",
    "VD.00 管理费用BKD",
    "VC&VD.01.2 详细测试 TOD",
    "VC&VD.01.2 TOD 年审",
    "VC&VD.01.2 TOD 预审+剩余期间",
    "样本池-TOD",
    "VD.01.3 复核法律费用",
    "VC&VD.01.4 截止性测试",
    "样本池-cut off",
]


def _reorder_sheets(wb):
    """按 SHEET_ORDER 重排工作表，返回最终顺序。

    各模块是「谁先建谁在前」：OAR / 口径Mapping 最后建就沉到末尾，两张样本池
    建在 01.4 之前。呈现顺序不能靠各模块抢位置，统一在落盘前排一次。
    用 move_sheet 逐个挪到目标下标，它是公开 API，比直接改 _sheets 稳。
    """
    names = list(wb.sheetnames)
    want = [n for n in SHEET_ORDER if n in names]
    rest = [n for n in names if n not in want]
    # 「上期TOD_xxx」跟着 TOD 预审页走，「上期BKD_xxx」跟着本期 VD 走（都是运行时名字）
    prior_tods = [n for n in rest if n.startswith("上期TOD_")]
    prior_bkds = [n for n in rest if n.startswith("上期BKD_")]
    if prior_tods or prior_bkds:
        rest = [n for n in rest if n not in prior_tods and n not in prior_bkds]
    order = []
    for n in want:
        order.append(n)
        if n == "VD.00 管理费用BKD":
            order.extend(prior_bkds)
        if n == "VC&VD.01.2 TOD 预审+剩余期间":
            order.extend(prior_tods)
    order.extend(rest)
    if order == names:
        return order
    for i, n in enumerate(order):
        try:
            cur = wb.sheetnames.index(n)
            if cur != i:
                wb.move_sheet(n, offset=i - cur)
        except Exception:
            pass
    return list(wb.sheetnames)


def _hide_aux_sheets(wb):
    """隐藏模板自带的纯说明页，返回被隐藏的表名。

    `VC&VD.01.2 详细测试 TOD` 只有标题和几行说明文字（真正的程序在
    年审 / 预审+剩余期间两页里做），留在前面占一张标签，项目组每次都要划过去。
    按「详细测试 + TOD」模糊命中，不写死全名 —— 模板改叫法也能跟着走。
    隐藏只改 sheet_state，不影响读值、不影响跨表公式。
    """
    hidden = []
    for name in list(wb.sheetnames):
        # 「详细测试 TOD」是模板自带的说明页；
        # 「上期BKD_/上期TOD_」是搬进来做链接依据的证据页，项目组平时不用翻，藏起来
        if (("详细测试" in name and "TOD" in name)
                or name.startswith("上期BKD_") or name.startswith("上期TOD_")):
            try:
                wb[name].sheet_state = "hidden"
                hidden.append(name)
            except Exception:
                pass
    return hidden


def _append_check_sheet(wb, stats, gaps, mapping_used, rmeta, balance_date, other):
    """追加「检查」表：概览 + 账套映射 + 待确认清单。"""
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    if "费用检查" in wb.sheetnames:
        del wb["费用检查"]
    ws = wb.create_sheet("费用检查")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    t15 = Font(name="宋体", size=15, bold=True)
    t9 = Font(name="宋体", size=9, color="666666")
    tb = Font(name="宋体", size=10, bold=True)
    t10 = Font(name="宋体", size=10)
    sec = Font(name="宋体", size=10, bold=True, color="1F3864")
    secf = PatternFill("solid", fgColor="F2F2F2")
    headf = PatternFill("solid", fgColor="D9E2F3")
    gapf = PatternFill("solid", fgColor=GAP_COLOR)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    center = Alignment(horizontal="center", vertical="center")

    for i, w in enumerate([34, 30, 26, 22, 18, 18, 46, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 1
    c = ws.cell(r, 1, "费用底稿（VC/VD BKD）生成检查表")
    c.font = t15
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    r += 1
    c = ws.cell(r, 1, "资产负债表日：%s　　期间：%s　　科目：销售费用/管理费用"
                % (balance_date or "（未填）", rmeta.get("period", "")))
    c.font = t9
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    r += 2

    def sect(txt):
        nonlocal r
        c = ws.cell(r, 1, txt)
        c.font = sec
        c.fill = secf
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        r += 1

    def head(cols):
        nonlocal r
        for i, h in enumerate(cols, start=1):
            c = ws.cell(r, i, h)
            c.font = tb
            c.border = border
            c.fill = headf
            c.alignment = center
        r += 1

    sect("一、概览")
    head(["项目", "行数", "本期账面数合计", "说明", "", "", "", ""])
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=8)
    for st in stats:
        c = ws.cell(r, 1, st["sheet"]); c.font = t10; c.border = border
        c = ws.cell(r, 2, st["rows"]); c.font = t10; c.border = border; c.alignment = right
        c = ws.cell(r, 3, st["sum"]); c.font = t10; c.border = border; c.alignment = right
        c.number_format = MONEY_FMT
        c = ws.cell(r, 4, "来源：序时账按「账簿 × 科目编码」汇总（借 − 贷）")
        c.font = t10; c.border = border; c.alignment = left
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
        r += 1
    if other:
        c = ws.cell(r, 1, "未能归类的费用行")
        c.font = t10; c.border = border
        c = ws.cell(r, 2, len(other)); c.font = t10; c.border = border
        c = ws.cell(r, 4, "一级科目既不含「销售费用」也不含「管理费用」，未写入底稿，请人工确认归属")
        c.font = t10; c.border = border; c.alignment = left
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
        r += 1

    r += 1
    sect("二、账套映射（B 列填公司名，原始账簿在此可追溯）")
    head(["账簿名称", "映射后公司名称", "命中方式", "", "", "", "", ""])
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=8)
    seen = set()
    for m in mapping_used:
        key = (m["book"], m["company"], m["how"])
        if key in seen:
            continue
        seen.add(key)
        c = ws.cell(r, 1, m["book"]); c.font = t10; c.border = border; c.alignment = left
        c = ws.cell(r, 2, m["company"] or "（未命中）"); c.font = t10; c.border = border; c.alignment = left
        c = ws.cell(r, 3, m["how"]); c.font = t10; c.border = border; c.alignment = center
        if not m["company"]:
            c.fill = gapf
            ws.cell(r, 2).fill = gapf
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        r += 1

    r += 1
    sect("三、待确认清单（底稿中标黄的格子，共 %d 处）" % len(gaps))
    if not gaps:
        c = ws.cell(r, 1, "无。")
        c.font = t10
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        return
    by_reason = {}
    for g in gaps:
        by_reason[g["reason"]] = by_reason.get(g["reason"], 0) + 1
    for reason, n in sorted(by_reason.items(), key=lambda x: -x[1]):
        c = ws.cell(r, 1, "· %s —— %d 处" % (reason, n))
        c.font = t10
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        r += 1
    r += 1
    head(["工作表", "单元格", "字段", "公司名称", "科目名称", "账簿", "原因", ""])
    ws.merge_cells(start_row=r - 1, start_column=7, end_row=r - 1, end_column=8)
    for g in gaps:
        ref = "%s%d" % (get_column_letter(g["col"]), g["row"])
        for i, v in enumerate([g["sheet"], ref, g["field"], g["company"], g["subject"],
                               g["acct"], g["reason"]], start=1):
            c = ws.cell(r, i, v)
            c.font = t10
            c.border = border
            c.alignment = left
            if i == 2:
                c.fill = gapf
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
        r += 1
