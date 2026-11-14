# -*- coding: utf-8 -*-
"""
货币资金 BKD 生成器 · 后端服务
================================
- 直接读取PTW平台 data/ledger.duckdb 的 tb 表，自动取「货币资金」科目余额（零导入）
- 接收「上年 BKD 底稿」上传，按「公司名称 + 银行账号」匹配
- 生成 BKD 底稿 + 检查表（openpyxl）

独立文件夹运行，不改动PTW平台任何代码。
"""
import os
import sys
import json
import re
import io
import time
import socket
import urllib.parse
import mimetypes
import threading
import webbrowser
from datetime import date as _date, datetime as _datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

APP_DIR = os.path.dirname(os.path.abspath(__file__))
# PTW平台「程序」目录 = 本文件上两级（modules/bkd -> modules -> 程序）
KZ_DIR = os.path.abspath(os.path.join(APP_DIR, "..", ".."))
# 数据目录统一在项目根目录 data/（本文件上三级 程序/modules/bkd -> 程序 -> 根）
_ROOT = os.path.abspath(os.path.join(KZ_DIR, ".."))
KZ_DATA = os.environ.get("KZ_DATA_DIR") or os.path.join(_ROOT, "data")
KZ_DB = os.path.join(KZ_DATA, "ledger.duckdb")

# BKD 的模板 / 输出 / 临时目录统一归到数据目录下（与 duckdb、规则、用户 excel 同处 data/）
WORK = os.path.join(KZ_DATA, "bkd_work")
OUT = os.path.join(KZ_DATA, "bkd_out")
TPL_DIR = os.path.join(KZ_DATA, "bkd_tpl")        # 内置/可替换模板目录
TPL_DEFAULT = os.path.join(TPL_DIR, "C SWP 货币资金 模板 V7.xlsx")
TPL_VERSION = "V7"
for d in (WORK, OUT, TPL_DIR):
    os.makedirs(d, exist_ok=True)
# 兼容：若历史模板仍在模块目录下，则迁移/回退到数据目录
_legacy_tpl = os.path.join(APP_DIR, "_tpl", "C SWP 货币资金 模板 V7.xlsx")
if os.path.exists(_legacy_tpl) and not os.path.exists(TPL_DEFAULT):
    os.makedirs(TPL_DIR, exist_ok=True)
    import shutil
    shutil.copy2(_legacy_tpl, TPL_DEFAULT)

try:
    import duckdb
except ImportError:
    # 源码运行时借 PTW 自带 runtime 的 site-packages；打包后依赖已内置，不走这里
    _rt = os.path.join(KZ_DIR, "runtime", "python", "Lib", "site-packages")
    if os.path.isdir(_rt):
        sys.path.insert(0, _rt)
    import duckdb

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
    from openpyxl.comments import Comment
except ImportError:
    sys.path.insert(0, os.path.join(KZ_DIR, "runtime", "python", "Lib", "site-packages"))
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
    from openpyxl.comments import Comment

import copy as _copy
from copy import copy as _copy_style

# 老式 .xls 读取（与PTW平台导入共用 core/xlsread.py）
try:
    from core import xlsread as _xlsread
except Exception:                                    # 独立运行时兜底
    try:
        sys.path.insert(0, KZ_DIR)
        from core import xlsread as _xlsread
    except Exception:
        _xlsread = None


def is_xls_bytes(data):
    return bool(_xlsread) and _xlsread.is_xls(data)


def load_wb_any(data, **kw):
    """统一入口：bytes → workbook，支持 .xlsx/.xlsm 与老式 .xls。"""
    if is_xls_bytes(data):
        bio = io.BytesIO()
        _xlsread.xls_to_xlsx(data, bio)
        bio.seek(0)
        return load_workbook(bio, **kw)
    return load_workbook(io.BytesIO(data), **kw)


# ============================================================ 文本/数值工具
def to_half(s):
    s = "" if s is None else str(s)
    out = []
    for ch in s:
        c = ord(ch)
        if 0xFF01 <= c <= 0xFF5E:
            out.append(chr(c - 0xFEE0))
        else:
            out.append(ch)
    return "".join(out).replace("\u3000", " ").replace("\u00a0", " ")


def T(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if isinstance(v, (int,)):
        return str(v)
    return str(v).replace("\u3000", " ").replace("\u00a0", " ").strip()


def _no_err(v):
    """把 Excel 错误值（#N/A / #REF! / #VALUE! …）当成空。

    上年底稿的某些列（尤其「性质和用途」）常留着 #N/A，那是上一期没填完的痕迹，
    不是内容。原样搬进「上年 C.00 BKD」再被本年 K 列引用，底稿里就整片 #REF/#N/A，
    看起来像程序算错。
    """
    t = T(v)
    if t.startswith("#") and t.upper() in (
            "#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!",
            "#GETTING_DATA"):
        return ""
    return t


def _is_number_text(v):
    """判断一个值是不是"纯数字/金额"（性质列出现这种值 = 取错列的强信号）。

    性质与用途应是「经营户 / 保证金户 / 第三方支付」这类文字；取到 1、7.1、
    1,234,567.89 这类值，几乎只可能是上年底稿列位错位（比如把折算率或金额列
    当成性质列）。不阻断生成，但必须留痕，不能静默写进底稿。
    """
    t = T(v).replace(",", "").replace("，", "").replace("%", "").strip()
    if not t:
        return False
    try:
        float(t)
        return True
    except (TypeError, ValueError):
        return False


def K(s):
    return re.sub(r"[\s\-—_/\\.,()（）\[\]{}:：;；'\"“”‘’、。·]", "", to_half(T(s))).upper()


def CK(s):
    t = K(T(s))
    t = re.sub(r"[（(][^）)]*[）)]", "", t)
    t = re.sub(r"(股份)?有限(责任)?公司$", "", t)
    t = re.sub(r"集团$", "", t)
    return t


def acct_key(s):
    return re.sub(r"[^0-9A-Z]", "", K(s))


def N(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) if v == v else 0.0  # NaN -> 0
    s = to_half(str(v)).replace(",", "").replace(" ", "").replace("¥", "").replace("￥", "").replace("$", "").replace("元", "")
    if not s:
        return 0.0
    neg = False
    if re.match(r"^\(.+\)$", s):
        neg = True
        s = s[1:-1]
    if re.match(r"^[负亏]", s):
        neg = True
        s = s[1:]
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m:
        return 0.0
    n = float(m.group(0))
    return -abs(n) if neg else n


def r2(v):
    return round(float(v or 0), 2)


def pick_acct(s):
    t = to_half(T(s)).replace(" ", "")
    if not t:
        return ""
    ms = re.findall(r"[0-9][0-9\-]{4,}", t)
    if not ms:
        # 短号：明确写了「账号/帐号/账户/卡号」的 2~4 位数字（脱敏数据、测试数据常见）。
        # 真实账号都是长串，走上面那个分支；这里只兜住显式标注过的短号。
        m2 = re.search(r"(?:账号|帐号|账户|卡号|账号1)[:：]?\s*([0-9][0-9\-]{1,})", t)
        if m2:
            return m2.group(1).strip("-")
        return ""
    best = max(ms, key=len)
    return best.strip("-")


#: 一眼能看出是「公司名」的词根。辅助核算项里装的常常是主体全称或往来单位，
#: 不是开户行——靠它挡掉，不能靠"等于本行主体名"这种精确判断：
#: 同一张表里 aux1 可能是分公司、关联方、或者大小写不同的同一个主体。
COMPANY_WORDS = re.compile(
    r"有限公司|股份公司|集团|分公司|子公司|控股|"
    r"Ltd|Limited|Inc\b|LLC|GmbH|B\.?V\.?|N\.?V\.?|Pte|Corp|Corporation|"
    r"Holdings|Technology|Technologies|Electronics|"
    r"科技|智能|贸易|软件|实业|物流|供应链",
    re.I)


def looks_like_company(t):
    """这个值像不像"单位名称"（而不是开户行）。"""
    return bool(COMPANY_WORDS.search(T(t)))


def derive_bank(name):
    n = T(name).replace(" ", "")
    # 一级科目本身不是开户行，别把「库存现金」当成银行名填进去
    if n in ("库存现金", "现金", "银行存款", "其他货币资金"):
        return ""
    parts = [p.strip() for p in re.split(r"[-－—–_/|]+", T(name))]
    parts = [p for p in parts
             if p and not re.fullmatch(r"\d+", p) and p != "银行存款"
             and not re.fullmatch(r"[借贷]", p)
             # 「其他货币资金-支付宝」拆出来的「其他货币资金」还是科目名，不是机构
             and not re.fullmatch(r"(其他)?货币资金|银行存款|库存现金", p)
             and not looks_like_company(p)]
    return parts[0] if parts else ""


def norm_cur(v):
    """币种归一。要吃掉「港币(HKD)」「美元(USD)」这种带代码后缀的写法（ERP 导出常见）。"""
    t0 = T(v)
    t = t0.upper()
    if not t:
        return ""
    if t in ("￥", "¥") or t.startswith(("人民币", "RMB", "CNY")):
        return "人民币"
    if t.startswith(("美元", "USD")):
        return "美元"
    if t.startswith(("港币", "港元", "HKD")):
        return "港币"
    if t.startswith(("欧元", "EUR")):
        return "欧元"
    if t.startswith(("日元", "JPY")):
        return "日元"
    if t.startswith(("新加坡元", "SGD")):
        return "新加坡元"
    if t.startswith(("英镑", "GBP")):
        return "英镑"
    return t0


#: 币种 → ISO 4217 三字母代码。
#: BKD 的「币种」列和 C.02a「外币测算」的「原币币种」必须用**同一套写法**，
#: 否则 C.02a 按币种 SUMIF BKD 时一个字都对不上：原币金额、账面本位币、差异
#: 三列全变 0，看着像"本期没有汇兑差异"，其实是根本没取到数。
CUR_ALIAS = {
    "人民币": "CNY", "RMB": "CNY", "CNY": "CNY", "本币": "CNY", "本位币": "CNY",
    "美元": "USD", "美金": "USD", "USD": "USD",
    "港币": "HKD", "港元": "HKD", "香港元": "HKD", "HKD": "HKD",
    "欧元": "EUR", "EUR": "EUR",
    "日元": "JPY", "日圆": "JPY", "日本日圆": "JPY", "JPY": "JPY",
    "英镑": "GBP", "GBP": "GBP",
    "新加坡元": "SGD", "新加坡币": "SGD", "SGD": "SGD",
    "新台币": "TWD", "TWD": "TWD",
    "澳币": "AUD", "澳大利亚元": "AUD", "澳元": "AUD", "AUD": "AUD",
    "墨西哥币": "MXN", "墨西哥比索": "MXN", "MXN": "MXN",
    "巴西雷亚尔": "BRL", "BRL": "BRL",
    "挪威克朗": "NOK", "NOK": "NOK",
    "韩元": "KRW", "KRW": "KRW",
    "泰铢": "THB", "THB": "THB",
    "加元": "CAD", "加拿大元": "CAD", "CAD": "CAD",
    "瑞士法郎": "CHF", "CHF": "CHF",
    "印度卢比": "INR", "INR": "INR",
}


def cur_code(v):
    """币种 → ISO 代码（USD / CNY / HKD…）。认不出的原样返回，不猜。

    账套里同一个币种有四五种写法（美元 / 美金 / USD / 美元(USD)），
    不归一就会出现"同一张表里两种写法"，SUMIF 和 VLOOKUP 全部失效。
    """
    t = T(v).strip()
    if not t:
        return ""
    up = t.upper()
    if up in CUR_ALIAS:
        return CUR_ALIAS[up]
    # 「美元(USD)」「港币（HKD）」这类带后缀的
    m = re.search(r"[（(]\s*([A-Za-z]{2,4})\s*[)）]", t)
    if m:
        code = CUR_ALIAS.get(m.group(1).upper())
        if code:
            return code
    base = re.sub(r"[（(].*?[)）]", "", t).strip().upper()
    return CUR_ALIAS.get(base, t)


def norm_confirm(v):
    t = T(v)
    if not t:
        return ""
    if re.match(r"^(是|Y|YES|√|✓|已函证|是（.*）)$", t, re.I):
        return "是"
    if re.match(r"^(否|N|NO|×|✗|未函证)$", t, re.I):
        return "否"
    return t


# ============================================================ 列识别规则
TB_SPEC = [
    ("company", "公司名称", 3, re.compile(r"公司名称|单位名称|主体名称|被审计单位|^公司$|^单位$|客户名称|企业名称")),
    ("acctNo", "银行账号", 3, re.compile(r"账号|帐号|账户号|卡号|银行账号")),
    ("acctCode", "科目编码", 2, re.compile(r"科目(编码|代码|编号)|^编码$|^代码$|^科目号$|科目号")),
    ("acctName", "科目名称", 3, re.compile(r"科目名称|全称|科\s*目$|科目")),
    ("bankName", "银行/机构名称", 1, re.compile(r"开户行|银行名称|存款机构|金融机构|开户银行")),
    ("currency", "币种", 1, re.compile(r"币种|货币|原币种")),
    ("endOrig", "期末余额（原币）", 2, re.compile(r"期末.*原币|原币.*期末|原币余额|外币余额|原币金额|期末外币")),
    ("endLocal", "期末余额（本币）", 2, re.compile(r"期末.*(本币|本位币|折合)|(本币|本位币|折合).*(余额|金额)")),
    ("debitEnd", "期末余额-借方", 2, re.compile(r"(期末|年末).*借")),
    ("creditEnd", "期末余额-贷方", 2, re.compile(r"(期末|年末).*贷")),
    ("endBal", "期末余额", 2, re.compile(r"期末(余额|数|金额|结存)|年末余额|^余额$|余额")),
    ("dir", "余额方向", 1, re.compile(r"方向")),
]

BKD_SPEC = [
    ("company", "公司名称", 3, re.compile(r"公司名称|被审计单位|单位名称|企业名称|^公司$|^单位$")),
    ("acctNo", "账号", 3, re.compile(r"账号|帐号|账户号|卡号"), True),
    ("subject", "科目名称", 2, re.compile(r"科目名称|^科目$|科目")),
    ("jNote", "J列：判断理由/替代程序", 1, re.compile(r"替代性程序|记录我们判断|替代程序|不重大")),
    ("bankName", "银行/存款机构名称", 2, re.compile(r"银行(?!函证)(名称|机构|\s*\/)|^银行|存款机构|开户行|金融机构|机构名称")),
    ("currency", "币种", 1, re.compile(r"币种|货币")),
    ("origEnd", "期末金额（原币）", 2, re.compile(r"原币")),
    ("localEnd", "期末金额（本币）", 2, re.compile(r"本币|本位币|折合")),
    ("endAny", "期末金额", 2, re.compile(r"期末|余额|金额")),
    ("nature", "性质和用途", 1, re.compile(r"性质|用途")),
    ("confirm", "是否函证", 1, re.compile(r"函证")),
    ("notes", "Notes/备注", 1, re.compile(r"^notes?$|备注", re.I)),
]


def detect_header(rows, spec, limit=50):
    best = {"i": 0, "score": 0}
    for i in range(min(len(rows), limit)):
        r = rows[i] or []
        sc, hit = 0, set()
        for cell in r:
            t = T(cell)
            if not t or len(t) > 60:
                continue
            for k, label, w, re_, *_ in spec:
                if k in hit:
                    continue
                if re_.search(t):
                    hit.add(k)
                    sc += w
        if sc > best["score"]:
            best = {"i": i, "score": sc}
    return best


def map_columns(headers, spec):
    mp, taken = {}, set()
    cells = [(i, T(h)) for i, h in enumerate(headers) if T(h) and len(T(h)) <= 80]
    for k, label, w, re_, *rest in spec:
        multi = bool(rest and rest[0])
        for i, t in cells:
            if not multi and i in taken:
                continue
            if not re_.search(t):
                continue
            if multi:
                mp.setdefault(k, [])
                if i not in mp[k]:
                    mp[k].append(i)
                taken.add(i)
            else:
                # 单值字段：匹配到第一列即占用并停止，避免同一正则误占后续列
                taken.add(i)
                mp[k] = i
                break
    return mp


# ============================================================ 从 DuckDB tb 表取数
def _open_db_readonly(db_path):
    """以只读方式打开PTW平台的 DuckDB（BKD 只读、绝不改账）。

    若PTW平台正在运行并独占锁住数据库，抛出带清晰提示的异常。
    """
    try:
        con = duckdb.connect(db_path, read_only=True)
        con.execute("SHOW TABLES").fetchall()  # 触发真实 IO，暴露锁冲突
        return con
    except Exception as e:
        msg = str(e)
        if ("另一个程序正在使用" in msg or "already open" in msg
                or "IO Error" in msg or "Permission" in msg or "Errno 13" in msg):
            raise RuntimeError(
                "PTW平台正在运行，数据库已被占用。请先关闭PTW平台，再使用 BKD 生成器。"
            ) from e
        raise


def read_tb_from_duckdb(keyword="货币资金", conn=None):
    """从PTW平台 DuckDB 的 tb 表读货币资金账户（库存现金/银行存款/其他货币资金），返回 (records, meta)。

    conn：若传入已有的共享 duckdb 连接（整合进PTW平台同一进程时），直接复用，
    不再单独 read_only 打开（避免同一文件不同配置的连接冲突）；否则回退到只读打开。
    """
    close_after = False
    if conn is None:
        if not os.path.exists(KZ_DB):
            return [], {"error": "找不到PTW平台数据库：%s" % KZ_DB, "db": KZ_DB}
        try:
            con = _open_db_readonly(KZ_DB)
            close_after = True
        except RuntimeError as e:
            return [], {"error": str(e), "db": KZ_DB}
    else:
        con = conn
    try:
        tabs = [r[0] for r in con.execute("SHOW TABLES").fetchall()]
        if "tb" not in tabs:
            return [], {"error": "PTW平台里还没有「科目余额表」（tb 表），请先在PTW平台里清洗并保存科目余额表。", "db": KZ_DB}
        n = con.execute("SELECT count(*) FROM tb").fetchone()[0]
        if n == 0:
            return [], {"error": "科目余额表（tb 表）是空的，请先在PTW平台里清洗科目余额表。", "db": KZ_DB, "rows": 0}
        # 列存在性先验：清洗后没映射「一级科目」时 tb 里就没有 acct_l1 列，
        # 下面的 SQL 会直接 BinderException（500）——先给可执行提示，别让人去猜
        _cols = [d[0] for d in con.execute("SELECT * FROM tb LIMIT 0").description]
        if "acct_l1" not in _cols:
            return [], {"error": "科目余额表没有「一级科目」列。请到「清洗与映射」启用"
                                  "「生成一级科目 / 提取一级科目」动作，或映射一级科目列，"
                                  "然后重新合并。", "db": KZ_DB}
        kw = T(keyword).replace(" ", "")
        # 严格只取「一级科目」为 银行存款 / 库存现金 / 其他货币资金 的行，其余一律不要。
        # 判断优先级：acct_l1（PTW平台清洗出的「一级科目」列）→ 兜底 acct_code / acct_name。
        # 注意是「精确匹配一级科目」，不用 LIKE 前缀，避免把 100201 等二级/三级明细误抓进来。
        sql = """
        SELECT company_name, acct_code, acct_name, acct_l1, currency,
               end_bal, end_debit, end_credit, end_bal_fx, end_debit_fx, end_credit_fx,
               bank_account, aux1, aux2, aux3, aux_text
        FROM tb
        WHERE (
            NULLIF(trim(acct_l1), '') IN ('1001','1002','1012','库存现金','银行存款','其他货币资金')
            OR (
                NULLIF(trim(acct_l1), '') IS NULL
                AND (
                    trim(acct_code) IN ('1001','1002','1012')
                    OR trim(acct_name) IN ('库存现金','银行存款','其他货币资金')
                )
            )
        )
        -- 排除「核算维度 = 合计/小计」的汇总行：多账簿导出里每个主体每个科目都有一条，
        -- 与维度明细并存时会让 BKD 金额翻倍。这一条是防御性的，用户没清洗干净也不会算错。
        AND NOT (
              trim(COALESCE(aux1, '')) IN ('合计','小计','总计','合计数')
           OR trim(COALESCE(aux2, '')) IN ('合计','小计','总计','合计数')
           OR trim(COALESCE(aux3, '')) IN ('合计','小计','总计','合计数')
           OR trim(COALESCE(aux_text, '')) IN ('合计','小计','总计','合计数')
        )
        -- 排除「综合本位币」行：它是该账号所有币别的本币合计，与各币别明细行重复，
        -- 全取会让金额正好翻倍（实测 14.2 亿 → 28.4 亿）。BKD 一个币种一行，要的是明细行。
        AND trim(COALESCE(currency, '')) NOT LIKE '%综合本位币%'
        """
        rows = con.execute(sql).fetchall()
    except Exception:
        # 异常路径先关连接再抛出；正常路径把关闭推到诊断之后（诊断还要查 tb）
        if close_after:
            con.close()
        raise

    SKIP = re.compile(r"^(合计|小计|总计|累计|合计数|本期合计|本年累计|其中)")
    records, skipped, skipped_company, last_co = [], 0, 0, ""
    for r in rows:
        # 16 列：company, code, name, acct_l1, currency, end_bal, end_debit, end_credit,
        #        end_bal_fx, end_debit_fx, end_credit_fx, bank_account, aux1, aux2, aux3, aux_text
        vals = list(r) + [None] * (16 - len(r)) if len(r) < 16 else list(r)
        (company, code, name, _l1, currency,
         end_bal, end_debit, end_credit,
         end_bal_fx, end_debit_fx, end_credit_fx,
         bank_account, aux1, aux2, aux3, aux_text) = vals
        nm = T(name)
        if not nm and not code:
            skipped += 1
            continue
        if SKIP.match(nm.replace(" ", "")) or SKIP.match(T(code)):
            skipped += 1
            continue
        co = T(company)
        # 合并方案层（多账簿/合并/汇总）不是真实主体，留着会让金额按主体数翻倍
        if re.search(r"多账簿|合并方案|汇总层|合并层", co) or co in ("合计", "汇总", "总计"):
            skipped_company += 1
            continue
        # 主体名归一：去掉 ERP 带出来的 (100) 这类编码后缀
        mco = re.match(r"^(.*?)\s*\(\s*\d{2,4}\s*\)\s*$", co)
        if mco and mco.group(1).strip():
            co = mco.group(1).strip()
        if co:
            last_co = co
        else:
            co = last_co
        # 银行账号：优先 bank_account，其次 aux 里的账号
        # 银行/存款机构名称 = 辅助核算项1 完整原值，一个字不拆、不改。
        # 客户把「开户行&账号」写在同一格，那就整串原样搬进底稿：
        # 复核时要按客户辅助核算逐条回查，拆过或改过名就对不上。
        bank = T(aux1) or T(aux2) or T(aux3) or T(aux_text)
        acct_no = pick_acct(bank_account) or pick_acct(aux_text) or pick_acct(nm)
        # 期末余额：优先用余额净额列；为空/0 时用 期末借方 - 期末贷方 兜底
        local = N(end_bal)
        if local == 0 and (end_debit is not None or end_credit is not None):
            local = N(end_debit) - N(end_credit)
        orig = N(end_bal_fx)
        if orig == 0 and (end_debit_fx is not None or end_credit_fx is not None):
            orig = N(end_debit_fx) - N(end_credit_fx)
        if orig == 0:
            orig = local
        # subject 一律用**一级科目**：账上科目名常带银行名（如
        # 「银行存款_招商银行-广州骏景广场支行（755976778910001）」），
        # 原样拿去做 bkd_key 会把同一科目拆成十几个，Lead 的 SUMIF 也会漏。
        subj_l1 = _norm_subject_l1(_l1) or _norm_subject_l1(nm) or (nm or "银行存款")
        records.append({
            "company": co, "subject": subj_l1, "bank": bank, "acctNo": acct_no,
            "currency": cur_code(currency) or "CNY",
            "orig": r2(orig), "local": r2(local),
            "acct_code": code,
        })

    # 合并同一公司+同一账号
    g, merged = {}, []
    for rec in records:
        # 键必须带科目和币种：只按「公司+账号」会把银行存款与其他货币资金并成一行、
        # 把多币种账号并成一行（币种取第一条，原币还会跨币种相加）
        key = bkd_key(rec["company"], rec["subject"], rec["acctNo"],
                      rec["bank"], rec["currency"])
        if key in g:
            e = g[key]
            e["orig"] = r2(e["orig"] + rec["orig"])
            e["local"] = r2(e["local"] + rec["local"])
            if not e["acctNo"] and rec["acctNo"]:
                e["acctNo"] = rec["acctNo"]
            if e["bank"] and rec["bank"] and rec["bank"] != e["bank"]:
                e["bank"] += "、" + rec["bank"]
            elif not e["bank"]:
                e["bank"] = rec["bank"]
        else:
            g[key] = dict(rec)
            merged.append(g[key])
    meta = {"rows": n, "banks": len(merged), "db": KZ_DB,
            "skipped": skipped, "skippedCompany": skipped_company}
    if not merged:
        meta.update(_diagnose_empty_tb(con, n))
    if close_after:
        con.close()
    return merged, meta


def _diagnose_empty_tb(con, n_rows):
    """取不到货币资金账户时**分因**给可执行指引（只诊断，不自动兜底）。

    为什么需要它：验收反馈「tb 里明明有银行存款 52 行却报没有账户」，实际是那 52 行
    全是「综合本位币」层，被上面 WHERE 的重复层排除规则整批滤掉了；而报错只说
    「确认一级科目已识别」，把人引到了错误方向，排查两轮都没找到真因。

    这里把 4 种真实原因分开说，每种都给出下一步动作。不改变取数语义——
    「按哪一层上数」是清洗阶段要显式配置的口径选择，不能藏在取数代码里替人决定。
    """
    try:
        cols = [d[0] for d in con.execute("SELECT * FROM tb LIMIT 0").description]
        one = lambda sql: con.execute(sql).fetchone()[0]
        money_where = ("(trim(coalesce(acct_l1,'')) IN ('1001','1002','1012','库存现金','银行存款','其他货币资金')"
                       " OR trim(coalesce(acct_code,'')) IN ('1001','1002','1012')"
                       " OR trim(coalesce(acct_name,'')) IN ('库存现金','银行存款','其他货币资金'))")
        total_money = one("SELECT count(*) FROM tb WHERE " + money_where)
        # ① 一级科目没派生：列不存在或整列为空
        if "acct_l1" not in cols:
            return {"error": "科目余额表没有「一级科目」列。请到「清洗与映射」启用"
                             "「生成一级科目」动作（或映射一级科目列）后重新合并。"}
        n_l1 = one("SELECT count(*) FROM tb WHERE nullif(trim(acct_l1),'') IS NOT NULL")
        if n_l1 == 0:
            return {"error": "「一级科目」整列为空（%s 行都没值）。请在清洗里启用"
                             "「生成一级科目 / 提取一级科目」动作后重新合并。" % format(n_rows, ",")}
        # ② 确实没有货币资金科目
        if total_money == 0:
            return {}          # 交给调用方用原有文案
        # ③ 全是「综合本位币」层 → 被重复层排除规则整批滤掉
        n_base = one("SELECT count(*) FROM tb WHERE %s AND coalesce(currency,'') LIKE '%%综合本位币%%'"
                     % money_where)
        n_detail = total_money - n_base
        if n_detail == 0:
            return {"error": "货币资金行（%d 行）全部是「综合本位币」汇总层，平台默认排除它"
                             "（与分币别明细重复，同时上数会翻倍），因此取不到账户。\n"
                             "请到「清洗与映射」二选一：\n"
                             "  ① 启用「删除综合本位币行」——保留分币别明细（底稿有原币与原币核对）；\n"
                             "  ② 启用「只保留综合本位币行」——按汇总层上数（底稿无原币明细列）。\n"
                             "两者都不配时，数据检查会把「综合本位币与明细并存」判为不达标并阻止合并。"
                             % total_money}
        # ④ 全是「核算维度=合计」的汇总行
        aux_hit = ("(trim(coalesce(aux1,'')) IN ('合计','小计','总计','合计数')"
                   " OR trim(coalesce(aux2,'')) IN ('合计','小计','总计','合计数')"
                   " OR trim(coalesce(aux3,'')) IN ('合计','小计','总计','合计数')"
                   " OR trim(coalesce(aux_text,'')) IN ('合计','小计','总计','合计数'))")
        n_aux = one("SELECT count(*) FROM tb WHERE %s AND %s" % (money_where, aux_hit))
        if n_aux == total_money:
            return {"error": "货币资金行（%d 行）全部是「核算维度=合计/小计」的汇总行，"
                             "没有账户级明细。请在清洗时保留辅助核算明细行，"
                             "或重新导出到末级/账户级。" % total_money}
        # ⑤ 兜底：把实际取值列出来，让用户自己看出差异（带空格、全角、写法不同）
        try:
            samples = [r[0] for r in con.execute(
                "SELECT DISTINCT trim(acct_l1) FROM tb WHERE %s LIMIT 8" % money_where).fetchall()]
        except Exception:
            samples = []
        return {"error": "货币资金行存在（%d 行）但都不满足精确匹配。一级科目实际取值：%s。"
                         "平台按「库存现金/银行存款/其他货币资金」或编码 1001/1002/1012 精确匹配，"
                         "不做模糊匹配——请检查是否有空格、全角、后缀（如「银行存款-工行」）。"
                         % (total_money, "、".join(str(s) for s in samples) or "-")}
    except Exception:
        return {}              # 诊断失败不影响主流程，仍走原有文案


# ============================================================ Cutoff 截止性测试（C.03）
CUTOFF_SHEET = "C.03 Cutoff"
CUTOFF_BEFORE_START = 20        # 期末前银行间转账：数据区首行（模板 20-24 共 5 行）
CUTOFF_AFTER_START = 29         # 期末后银行间转账：数据区首行（模板 29-33 共 5 行）
CUTOFF_PRESET_ROWS = 5          # 模板每个数据块预置的行数（超出时自动插入行）
CUTOFF_CHECK_COL = 16           # P 列 Check（=M+I）

_MONEY_CODES = ("1001", "1002", "1012")
_MONEY_NAMES = ("库存现金", "银行存款", "其他货币资金")
#: 「这一格该有内容、但业务上确实没有」的占位符。
#: 库存现金没有开户行和账号、第三方支付没有银行账号 —— 都不是缺资料，
#: 既不能留空（会被当成没填）也不能填 0（会被当成金额），用 / 明确表示"不适用"。
NO_INFO = "/"


def _is_money_val(v):
    """单个值是否指向货币资金科目。

    序时账里一级科目常常留空、明细科目是 100201 这种下级码、名称是「银行存款-工行」，
    所以这里不能用全等，要按「编码前 4 位 / 名称包含关键字」判断。
    """
    t = T(v).replace(" ", "").replace("　", "")
    if not t:
        return False
    if t in _MONEY_CODES or t in _MONEY_NAMES:
        return True
    for kw in _MONEY_NAMES:
        if kw in t:
            return True
    digits = re.sub(r"[^0-9]", "", t)
    return len(digits) >= 4 and digits[:4] in _MONEY_CODES


def is_money_account(code, name, l1):
    """是否属于货币资金三个一级科目（库存现金 / 银行存款 / 其他货币资金）。

    三个值任一命中即算：一级科目（若已映射）、科目名称、科目编码。
    """
    return _is_money_val(l1) or _is_money_val(name) or _is_money_val(code)


def _as_date(v):
    """把 DuckDB 的 DATE / 字符串统一成 datetime.date，失败返回 None。"""
    if v is None:
        return None
    if isinstance(v, _datetime):
        return v.date()
    if isinstance(v, _date):
        return v
    s = T(v)[:10]
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y.%m.%d"):
        try:
            return _datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _dstr(v):
    return v.strftime("%Y-%m-%d") if isinstance(v, _date) else ""


def read_journal_rows(conn=None, date_from=None, date_to=None):
    """读序时账（journal 表）指定日期窗口内的货币资金科目行，返回 (rows, meta)。

    只保留属于库存现金/银行存款/其他货币资金的凭证行（截止性测试只看银行间转账）。
    """
    close_after = False
    if conn is None:
        if not os.path.exists(KZ_DB):
            return [], {"error": "找不到PTW平台数据库：%s" % KZ_DB}
        try:
            con = _open_db_readonly(KZ_DB)
            close_after = True
        except RuntimeError as e:
            return [], {"error": str(e)}
    else:
        con = conn
    try:
        tabs = [r[0] for r in con.execute("SHOW TABLES").fetchall()]
        if "journal" not in tabs:
            return [], {"error": "PTW平台里还没有「序时账」（journal 表），请先导入并合并序时账。"}
        total = con.execute("SELECT count(*) FROM journal").fetchone()[0]
        if total == 0:
            return [], {"error": "序时账（journal 表）是空的，请先在PTW平台里合并序时账。", "rows": 0}
        cols = [r[0] for r in con.execute("DESCRIBE journal").fetchall()]

        def _sel(n):
            return n if n in cols else ("NULL AS " + n)

        ORDER = ("vdate", "vno", "summary", "acct_code", "acct_name", "acct_l1",
                 "debit", "credit", "currency", "company_name", "book_name",
                 "bank_account", "aux1", "aux2", "aux3", "aux_text",
                 "amount", "direction")
        base = "SELECT " + ", ".join(_sel(c) for c in ORDER) + " FROM journal"
        conds, params = [], []
        if date_from is not None:
            # 用 TRY_CAST：日期列若混有非标准格式/空值，不会整条 SQL 报错中断
            conds.append("TRY_CAST(vdate AS DATE) >= TRY_CAST(? AS DATE)")
            params.append(str(date_from))
        if date_to is not None:
            conds.append("TRY_CAST(vdate AS DATE) <= TRY_CAST(? AS DATE)")
            params.append(str(date_to))
        sql = base
        if conds:
            sql += " WHERE " + " AND ".join(conds)
        sql += " ORDER BY vdate, vno" + (", src_row" if "src_row" in cols else "")
        raw = con.execute(sql, params).fetchall()

        # 兜底：期间内用 SQL 一行都没取到时，说明日期列格式不被识别（如 20251231、空值），
        # 改为全表取回后按 Python 解析日期过滤，避免「明明有数据却抓不到」。
        if conds and not raw and total:
            raw_all = con.execute(base + (" ORDER BY src_row" if "src_row" in cols else "")).fetchall()
            _lo = _as_date(date_from) if date_from is not None else None
            _hi = _as_date(date_to) if date_to is not None else None
            picked = []
            for _r in raw_all:
                _d = _as_date(_r[0])
                if _d is None:
                    continue
                if _lo is not None and _d < _lo:
                    continue
                if _hi is not None and _d > _hi:
                    continue
                picked.append(_r)
            raw = picked
    finally:
        if close_after:
            con.close()

    out = []
    for r in raw:
        vals = list(r) + [None] * (18 - len(r))
        (vdate, vno, summary, code, name, l1, debit, credit, currency,
         company, book, bank_account, aux1, aux2, aux3, aux_text,
         amount, direction) = vals[:18]
        if not is_money_account(code, name, l1):
            continue
        d, c = r2(N(debit)), r2(N(credit))
        # 兼容「金额(单列) + 借贷方向」的导出格式
        if abs(d) < 0.005 and abs(c) < 0.005:
            amt = abs(r2(N(amount)))
            if amt >= 0.005:
                dirs = T(direction)
                if "贷" in dirs or dirs.lower().startswith(("c", "cr")):
                    c = amt
                elif "借" in dirs or dirs.lower().startswith(("d", "dr")):
                    d = amt
        if abs(d) < 0.005 and abs(c) < 0.005:
            continue
        # 与 BKD 同一口径：辅助核算项1 完整原值，不拆
        bank = T(aux1) or T(aux2) or T(aux3) or T(aux_text)
        acct_no = pick_acct(bank_account) or pick_acct(aux_text) or pick_acct(name)
        out.append({
            "vdate": _as_date(vdate), "vno": T(vno), "summary": T(summary),
            "code": T(code), "name": T(name), "l1": T(l1),
            "debit": d, "credit": c,
            "currency": cur_code(currency) or "CNY",
            "company": T(company), "book": T(book),
            # bank 空就空着，不再拿科目名（「银行存款」）冒充机构名
            "acctNo": acct_no, "bank": bank,
        })
    return out, {"rows": total, "moneyRows": len(out)}


def _row_amount(row):
    """取该行的发生额（借方为正，贷方取正数），返回 (方向, 金额)。方向：recv=收，pay=付。"""
    if abs(row["debit"]) >= abs(row["credit"]) and abs(row["debit"]) >= 0.005:
        return "recv", abs(row["debit"])
    if abs(row["credit"]) >= 0.005:
        return "pay", abs(row["credit"])
    return "", 0.0


def find_transfers(rows, end_date, before_days=1, after_days=1, threshold=100000.0):
    """【已停用 · 2026-11-06】序时账配对找银行间转账。

    C.03 的样本改为取自银行流水（见 routes._build_cutoff 的说明），这里不再被调用。
    保留代码只为将来若要做"序时账口径"的辅助分析；**不要**再把它接回 C.03——
    序时账看不到收付款双方，凑出来的配对是推测，不是截止性测试的证据。
    """
    """从序时账货币资金行里识别「银行间转账」，配对成截止性测试样本。

    配对规则（按优先级）：
      1. 同一凭证号 + 同一主体 + 金额相等的一借一贷（同一天完成的转账）
      2. 跨期配对：付款日与收款日分别在资产负债表日两侧、金额相等、主体相同
         （这是截止性测试真正关注的：一笔钱期末转出、期后才到账）
    每笔配对生成一行样本：付款方填 I 列负数、收款方填 M 列正数，P 列 Check 应归零。
    """
    end = _as_date(end_date)
    if end is None:
        raise ValueError("截止期间（资产负债表日）格式不对，应为 YYYY-MM-DD")
    before_days = max(0, int(before_days or 0))
    after_days = max(0, int(after_days or 0))
    threshold = float(threshold or 0)

    pays, recvs, single = [], [], []
    for r in rows:
        kind, amt = _row_amount(r)
        if not kind or amt < 0.005:
            continue
        r = dict(r, _kind=kind, _amt=amt)
        (pays if kind == "pay" else recvs).append(r)

    used_recv, used_pay = set(), set()
    pairs = []

    def _key(r):
        # 凭证号按月重复（记-1 每月都有），所以必须带日期，避免跨月错配
        return (CK(r["company"]) + "|" + K(r["vno"]) + "|"
                + (r["vdate"].isoformat() if r["vdate"] else ""))

    # ---- 第一轮：同凭证号 + 同主体 + 同一天 + 等额 一借一贷
    idx = {}
    for j, rv in enumerate(recvs):
        idx.setdefault(_key(rv), []).append(j)
    for i, py in enumerate(pays):
        for j in idx.get(_key(py), []):
            if j in used_recv:
                continue
            rv = recvs[j]
            if abs(rv["_amt"] - py["_amt"]) <= 0.01:
                used_pay.add(i)
                used_recv.add(j)
                # 同凭证号但日期分处截止日两侧，同样算跨期
                py_d, rv_d = py["vdate"], rv["vdate"]
                is_cross = bool(py_d and rv_d and ((py_d <= end < rv_d) or (rv_d <= end < py_d)))
                pairs.append((py, rv, is_cross))
                break

    # ---- 第二轮：跨期配对（剩余单边）——付款与收款分处截止日两侧、金额相等
    left_pay = [i for i in range(len(pays)) if i not in used_pay]
    left_recv = [j for j in range(len(recvs)) if j not in used_recv]
    for i in left_pay:
        py = pays[i]
        py_d = py["vdate"]
        if py_d is None:
            continue
        for j in left_recv:
            if j in used_recv:
                continue
            rv = recvs[j]
            rv_d = rv["vdate"]
            if rv_d is None:
                continue
            if abs(rv["_amt"] - py["_amt"]) > 0.01:
                continue
            if CK(rv["company"]) != CK(py["company"]):
                continue
            # 必须一前一后跨过资产负债表日
            if not ((py_d <= end < rv_d) or (rv_d <= end < py_d)):
                continue
            used_pay.add(i)
            used_recv.add(j)
            pairs.append((py, rv, True))
            break

    # ---- 未配对的单边行（只转出没转入 / 只转入没转出）留档提示
    for i, py in enumerate(pays):
        if i not in used_pay:
            single.append(dict(py, _side="付款方未配对"))
    for j, rv in enumerate(recvs):
        if j not in used_recv:
            single.append(dict(rv, _side="收款方未配对"))

    # ---- 生成样本
    def _sample(py, rv, cross, no):
        amt = py["_amt"] if py["_amt"] else rv["_amt"]
        notes_ = []
        if cross:
            notes_.append("跨期：付款与到账分处资产负债表日两侧")
        if not py["acctNo"]:
            notes_.append("付款账号未识别")
        if not rv["acctNo"]:
            notes_.append("对方账号未识别")
        return {
            "no": no,
            "company": py["company"] or rv["company"],
            "payBank": py["bank"], "payAcct": py["acctNo"],
            "payDate": _dstr(py["vdate"]), "vno": py["vno"],
            "currency": cur_code(py["currency"] or rv["currency"]) or "CNY",
            "payAmt": r2(-amt), "recvAmt": r2(amt),
            "recvBank": rv["bank"], "recvAcct": rv["acctNo"],
            "recvDate": _dstr(rv["vdate"]),
            "cross": "是" if cross else "否",
            "note": "；".join(notes_),
            "summary": py["summary"] or rv["summary"],
            "amt": r2(amt),
        }

    samples = []
    for k, (py, rv, cross) in enumerate(pairs):
        samples.append(_sample(py, rv, cross, 0))

    # ---- 过滤阈值 + 归入期末前/期末后（按付款日判断）
    before, after = [], []
    for s in samples:
        if abs(s["amt"]) < threshold:
            continue
        d = _as_date(s["payDate"])
        if d is not None and d > end:
            after.append(s)
        else:
            before.append(s)
    before.sort(key=lambda x: (_as_date(x["payDate"]) or _date(1900, 1, 1), -abs(x["amt"])))
    after.sort(key=lambda x: (_as_date(x["payDate"]) or _date(1900, 1, 1), -abs(x["amt"])))
    for i, s in enumerate(before):
        s["no"] = "B%02d" % (i + 1)
    for i, s in enumerate(after):
        s["no"] = "A%02d" % (i + 1)

    win_start = end - timedelta(days=before_days)
    win_end = end + timedelta(days=after_days)
    stat = {
        "endDate": _dstr(end),
        "windowStart": _dstr(win_start), "windowEnd": _dstr(win_end),
        "beforeDays": before_days, "afterDays": after_days,
        "threshold": r2(threshold),
        "journalRows": len(rows),
        "transferPairs": len(samples),
        "beforeCount": len(before), "afterCount": len(after),
        "crossCount": sum(1 for s in before + after if s["cross"] == "是"),
        "belowThreshold": len(samples) - len(before) - len(after),
        "singleCount": len(single),
        # 现在会把全部样本写进 C.03（超出模板 5 行时自动插行），shown* 即实际写入数
        "shownBefore": len(before),
        "shownAfter": len(after),
    }
    return {"before": before, "after": after, "single": single, "stat": stat}


def probe_journal(conn=None):
    """序时账自检：把 journal 表的结构/日期范围/科目分布摊出来，用于排查"抓不到样本"。"""
    close_after = False
    if conn is None:
        if not os.path.exists(KZ_DB):
            return {"error": "找不到PTW平台数据库：%s" % KZ_DB}
        try:
            con = _open_db_readonly(KZ_DB)
            close_after = True
        except RuntimeError as e:
            return {"error": str(e)}
    else:
        con = conn
    try:
        tabs = [r[0] for r in con.execute("SHOW TABLES").fetchall()]
        if "journal" not in tabs:
            return {"error": "PTW平台里还没有「序时账」（journal 表）", "tables": tabs}
        total = con.execute("SELECT count(*) FROM journal").fetchone()[0]
        cols = [r[0] for r in con.execute("DESCRIBE journal").fetchall()]
        out = {"total": total, "columns": cols}
        if total == 0:
            out["error"] = "序时账是空的：请先到「导入与清洗 → 合并成表」把序时账合并进来"
            return out

        def _one(sql):
            try:
                return con.execute(sql).fetchone()
            except Exception:
                return None

        r = _one("SELECT min(TRY_CAST(vdate AS DATE)), max(TRY_CAST(vdate AS DATE)) FROM journal") \
            if "vdate" in cols else None
        out["dateMin"] = _dstr(r[0]) if r and r[0] else ""
        out["dateMax"] = _dstr(r[1]) if r and r[1] else ""
        if "vdate" in cols:
            bad = _one("SELECT count(*) FROM journal WHERE TRY_CAST(vdate AS DATE) IS NULL")
            out["badDate"] = int(bad[0]) if bad else None

        # 科目分布：按编码+名称聚合，并标注是否被识别为货币资金
        if "acct_code" in cols or "acct_name" in cols:
            code = "acct_code" if "acct_code" in cols else "NULL"
            name = "acct_name" if "acct_name" in cols else "NULL"
            l1 = "acct_l1" if "acct_l1" in cols else "NULL"
            rows = con.execute(
                "SELECT CAST(%s AS VARCHAR), CAST(%s AS VARCHAR), CAST(%s AS VARCHAR), count(*) "
                "FROM journal GROUP BY 1,2,3 ORDER BY 4 DESC LIMIT 40" % (code, name, l1)).fetchall()
            top, money_cnt, money_rows = [], 0, 0
            for c_, n_, l_, cnt in rows:
                ok = is_money_account(c_, n_, l_)
                if ok:
                    money_cnt += 1
                    money_rows += int(cnt)
                top.append({"code": T(c_), "name": T(n_), "l1": T(l_), "count": int(cnt), "money": ok})
            out["topAccounts"] = top
            out["moneyAccounts"] = money_cnt
            out["moneyRows"] = money_rows
        # 借贷列是否存在有效金额
        for k in ("debit", "credit", "amount", "direction"):
            out["has_" + k] = k in cols
        if "debit" in cols and "credit" in cols:
            r = _one("SELECT count(*) FROM journal WHERE COALESCE(debit,0)=0 AND COALESCE(credit,0)=0")
            out["zeroAmountRows"] = int(r[0]) if r else None
        return out
    finally:
        if close_after:
            con.close()


# ============================================================ 上年 BKD 底稿解析
def parse_prior_xlsx(data, sheets_info):
    """解析上年底稿（openpyxl），返回记录。支持 .xlsx/.xlsm/.xls。"""
    wb = load_wb_any(data, data_only=True)
    SKIP = re.compile(r"^(合计|小计|总计|累计|合计数|本期累计|本年累计|其中)")
    out, info = [], []
    for sh in sheets_info:
        name = sh["sheet"]
        header_row = sh.get("header_row", 1)
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        # 读成二维数组
        grid = [[c.value for c in row] for row in ws.iter_rows()]
        hd = header_row - 1
        headers = [T(c) for c in (grid[hd] if hd < len(grid) else [])]
        mp = sh.get("map") or map_columns(headers, BKD_SPEC)
        if not (mp.get("acctNo") or mp.get("bankName") or mp.get("subject")):
            info.append({"name": name, "count": 0, "skipped": True})
            continue
        n, last_co = 0, ""
        for i in range(hd + 1, len(grid)):
            r = grid[i] or []
            if not any(T(c) for c in r):
                continue
            co = T(r[mp["company"]]) if mp.get("company") is not None else ""
            if co:
                last_co = co
            else:
                co = last_co or name
            if re.match(r"^(公司名称|单位名称|被审计单位|合计|小计|总计)$", co.replace(" ", "")):
                continue
            bank = T(r[mp["bankName"]]) if mp.get("bankName") is not None else ""
            subj = T(r[mp["subject"]]) if mp.get("subject") is not None else ""
            if SKIP.match(subj.replace(" ", "")) and not bank:
                continue
            accts = []
            if mp.get("acctNo"):
                idxs = mp["acctNo"] if isinstance(mp["acctNo"], list) else [mp["acctNo"]]
                for ci in idxs:
                    v = pick_acct(r[ci]) or T(r[ci])
                    if v and len(acct_key(v)) >= 3 and v not in accts:
                        accts.append(v)
            if not accts:
                v2 = pick_acct(bank)
                if v2:
                    accts.append(v2)
            if not accts and not bank and not subj:
                continue
            rec = {
                "company": co, "bank": bank, "subject": subj,
                "currency": cur_code(r[mp["currency"]]) if mp.get("currency") is not None else "",
                "orig": N(r[mp["origEnd"]]) if mp.get("origEnd") is not None else None,
                "local": N(r[mp["localEnd"]]) if mp.get("localEnd") is not None else None,
                "endAny": N(r[mp["endAny"]]) if mp.get("endAny") is not None else None,
                "nature": _no_err(r[mp["nature"]]) if mp.get("nature") is not None else "",
                "confirm": norm_confirm(r[mp["confirm"]]) if mp.get("confirm") is not None else "",
                "jNote": T(r[mp["jNote"]]) if mp.get("jNote") is not None else "",
                "notes": T(r[mp["notes"]]) if mp.get("notes") is not None else "",
            }
            if rec["local"] is None:
                rec["local"] = rec["endAny"] if rec["endAny"] is not None else (rec["orig"] if rec["orig"] is not None else 0)
            if rec["orig"] is None:
                rec["orig"] = rec["local"]
            rec["local"] = r2(rec["local"])
            rec["orig"] = r2(rec["orig"])
            if accts:
                for a in accts:
                    out.append(dict(rec, acctNo=a))
                    n += 1
            else:
                out.append(dict(rec, acctNo=""))
                n += 1
        info.append({"name": name, "count": n, "skipped": False})
    return out, info


# ============================================================ 匹配
def bank_name_only(v):
    """取辅助核算里的「银行名称」部分——去掉 & 后面的账号。

    只用于**匹配**，不用于写入：本年 BKD 的 D 列仍然是完整的辅助核算原值，
    但两年比对时，「汇丰银行阿姆斯特丹分行&旧号」和「汇丰银行阿姆斯特丹分行&新号」
    是同一家银行的同一个户，必须能对上，否则会误判成销户+新开户两件事。
    """
    t = T(v)
    for sep in ("&", "＆"):
        if sep in t:
            return t.split(sep, 1)[0].strip()
    return t


PRIOR_SHEET = "上年 C.00 BKD"
#: 上年 sheet 里「账号」所在列（1-based）。公式引用要用它，写死字母迟早错列。
PRIOR_ACCT_COL = 4

#: 上年 sheet 的列（与本年 BKD 同构，便于并排看）
PRIOR_COLS = (
    ("company", "公司名称", 26),
    ("subject", "科目名称", 14),
    ("bank", "银行/存款机构名称", 40),
    ("acctNo", "账号1", 24),
    ("currency", "币种", 8),
    ("nature", "货币资金账户的性质和用途", 26),
    ("orig", "期末账面金额（原币）", 18),
    ("local", "期末账面金额（本币）", 18),
    ("rmb", "期末账面金额（人民币）", 18),
    ("notes", "Notes", 30),
)
#: 「性质和用途」在上年 sheet 的列号：本年 K 列要写公式引这一列
PRIOR_NATURE_COL = [k for k, _h, _w in PRIOR_COLS].index("nature") + 1


def embed_prior_sheet(wb, prior_rows, index_after="C.00 BKD"):
    """把上年底稿的明细写成一个独立 sheet（**纯数值，不留公式**）。

    返回 {匹配键: sheet 行号}，供本年账号列写公式引用。
    匹配键先用「银行名+币种+账号」，再用「银行名+币种」顺序兜底：
    同一家银行可能开多个户，只按名字配会把几个户全指向同一个号。
    """
    if not prior_rows:
        return {}, {}
    if PRIOR_SHEET in wb.sheetnames:
        del wb[PRIOR_SHEET]
    pos = (wb.sheetnames.index(index_after) + 1) if index_after in wb.sheetnames else len(wb.sheetnames)
    ws = wb.create_sheet(PRIOR_SHEET, pos)

    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.cell(1, 1, "上年度货币资金明细（来源：上年底稿，纯数值）").font = Font(name="Arial", size=11, bold=True)
    for j, (_k, head, w) in enumerate(PRIOR_COLS, start=1):
        c = ws.cell(2, j, head)
        c.font = Font(name="Arial", size=10, bold=True)
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = PatternFill("solid", fgColor="DCE6F1")
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[2].height = 32
    ws.freeze_panes = "A3"

    exact, bybank, byacct = {}, {}, {}
    r = 3
    for p in prior_rows:
        bank = T(p.get("bank"))
        acct = T(p.get("acctNo"))
        cur = cur_code(p.get("currency")) or "CNY"
        vals = {
            "company": T(p.get("company")), "subject": _norm_subject_l1(p.get("subject")) or T(p.get("subject")),
            "bank": bank, "acctNo": acct, "currency": cur,
            "nature": _no_err(p.get("nature")),
            "orig": r2(p.get("orig")), "local": r2(p.get("local")),
            "rmb": r2(p.get("local")), "notes": _no_err(p.get("notes")),
        }
        for j, (k, _h, _w) in enumerate(PRIOR_COLS, start=1):
            v = vals.get(k)
            cell = ws.cell(r, j, v)
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            if k in ("orig", "local", "rmb"):
                cell.number_format = ACCT_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif k == "acctNo":
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        if acct:
            exact.setdefault((K(bank_name_only(bank)), K(cur), acct_key(acct)), r)
        # 按（账号, 币种）索引：本年账号列直接引这里的行。
        # 同一账号多币种（如 798271635 的 GBP / EUR 两行）必须各指各的行，
        # 只按账号索引会让两行抢同一个格，另一个币种的账号就取不到。
        _ak = acct_key(acct) or _acct_of_prikey(bank)
        if _ak:
            byacct.setdefault((_ak, K(cur)), r)
            byacct.setdefault((_ak, None), r)      # 币种对不上时的兜底
        # 记 (行号, 账号)：同名多户时要按账号相似度挑，光有行号挑不准
        bybank.setdefault((K(bank_name_only(bank)), K(cur)), []).append((r, acct))
        r += 1

    # 合计行也写成数值（不留 SUM 公式）。
    # 原币列不做合计：一列里混着 USD/EUR/JPY，相加没有意义，写个数字反而误导。
    ws.cell(r, 1, "合计").font = Font(name="Arial", size=10, bold=True)
    for j, (k, _h, _w) in enumerate(PRIOR_COLS, start=1):
        if k in ("local", "rmb"):
            cell = ws.cell(r, j, r2(sum((p.get("local") or 0) for p in prior_rows)))
            cell.number_format = ACCT_FMT
            cell.font = Font(name="Arial", size=10, bold=True)
        elif k == "orig":
            cell = ws.cell(r, j, "—")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="Arial", size=10, bold=True)
    return exact, bybank, byacct


# ============================================================ 匹配
def bkd_prikey(v):
    """两年比对的主键 = 「银行/存款机构名称」整串（辅助核算项1，通常「开户行&账号」）。

    为什么用它而不是账号列：TB 的「银行账户」列经常是空的、或只给一串数字，
    而上年 BKD 与本年 BKD 都完整保留了客户的辅助核算原值（一个字没拆）。
    两年能直接对上的只有这一串，账号反而是它的一个片段。
    """
    t = T(v).replace(" ", "").replace("\u3000", "").replace("＆", "&").upper()
    if t in ("", "/", "-", "—", "N/A", "NA", "无"):
        return ""
    return t


def _acct_of_prikey(v):
    """从主键里取账号片段（「开户行&账号」取 & 后面那段），取不到返回空。"""
    t = T(v)
    if not t:
        return ""
    for sep in ("&", "＆"):
        if sep in t:
            return acct_key(t.rsplit(sep, 1)[-1])
    return acct_key(pick_acct(t) or "")


def bkd_key(company, subject, acct_no, bank, currency):
    """BKD 一行唯一键：主体 + 一级科目 + 账号(或银行) + 币种。

    缺一项就会出真实差错（飞书 1,744.34 万并进银行存款、61.20 万币种标错）：
      - 少了「科目」：同一银行账号下的银行存款和其他货币资金被并成一行；
      - 少了「币种」：同一账号的美元行和新加坡币行并成一行，币种取第一条记录，
        原币还会被跨币种相加。
    账号为空时才退化成「银行名 / 科目」，此时科目已在键里，不会串科目。
    """
    l1 = _norm_subject_l1(subject) or K(subject)
    ident = acct_key(acct_no) or ("B" + K(bank) if T(bank) else "S" + K(subject))
    cur = K(cur_code(currency) or "CNY")
    return "|".join([CK(company), l1, ident, cur])


def merge(tb, prior, rate_threshold=0.5, amt_threshold=100000):
    pmap, pri, dup = {}, [], 0
    for p in prior:
        k = bkd_key(p["company"], p["subject"], p.get("acctNo", ""),
                    p.get("bank", ""), p.get("currency"))
        if k in pmap:
            dup += 1
            continue
        pmap[k] = p
        pri.append(p)

    # 库存现金不参与两年比对：它没有银行、没有账号，主键是「/」，
    # 硬按「公司+科目」配出来的"已匹配"没有任何信息量（用户 2026-11-07 指出）。
    # 两年池子都剔掉它，否则上年的现金行会掉进「本期账面无此账号」白列一行。
    def _is_cash(x):
        return _norm_subject_l1((x or {}).get("subject")) == "库存现金"

    pri_all = list(pri)
    pri = [p for p in pri_all if not _is_cash(p)]

    idx_key, idx_acct, idx_subj = {}, {}, {}
    for i, p in enumerate(pri):
        p["_i"] = i
        ck = CK(p["company"])
        pk = bkd_prikey(p.get("bank"))
        ak = _acct_of_prikey(p.get("bank")) or acct_key(p.get("acctNo", ""))
        cu = K(cur_code(p.get("currency")) or "CNY")
        if pk:
            idx_key.setdefault(ck + "|" + pk, []).append(p)
        if ak:
            idx_acct.setdefault(ck + "|" + ak, []).append(p)
        if p["subject"]:
            # 主键为空（库存现金这种没有银行没有账号的）才用「公司+科目+币种」兜底
            idx_subj.setdefault(ck + "|S" + K(p["subject"]) + "|" + cu, []).append(p)
    used = set()

    def take(arr, cu=None):
        """取一条还没用过的上年记录。

        同币种优先：同一个主键可能对应多币种行（一个账户收多种币），
        只按顺序取会把美元行配到欧元行上，原币金额跨币种相加就是错的。
        """
        cand = [x for x in (arr or []) if x["_i"] not in used]
        if cu:
            for x in cand:
                if K(cur_code(x.get("currency")) or "CNY") == cu:
                    return x
        return cand[0] if cand else None

    rows, news, nonexact, matched = [], [], [], 0
    for t in tb:
        ck = CK(t["company"])
        pk = bkd_prikey(t.get("bank"))
        cu = K(cur_code(t.get("currency")) or "CNY")
        # 库存现金直接跳过：不进比对、不计新增、识别结果列留空
        cash_row = _is_cash(t)
        p, how = None, ("不适用" if cash_row else "")
        # ① 唯一主键（辅助核算项1）完全相同：这是主判据
        if not cash_row and pk and (ck + "|" + pk) in idx_key:
            p = take(idx_key[ck + "|" + pk], cu)
            if p:
                how = "唯一主键"
        # ② 主键写法变了（如「汇丰银行」改成「汇丰银行（中国）」），但账号没变：
        #    仍算同一户。判成「销户 + 新开户」会凭空造出一个审计关注点。
        if not p and not cash_row:
            ak_t = _acct_of_prikey(t.get("bank"))
            if ak_t and (ck + "|" + ak_t) in idx_acct:
                p = take(idx_acct[ck + "|" + ak_t], cu)
                if p:
                    how = "账号"
        if p:
            used.add(p["_i"])
            matched += 1
        r = {
            "company": t["company"] or (p and p["company"]) or "",
            # 科目和银行必须用**本年的**：账上这一期挂在哪个科目、哪个支行，是本期的
            # 事实；上年的科目只是去年的分类。用上年覆盖会让本年银行存款被标成
            # 其他货币资金，Lead 的 SUMIF 把金额归到错的科目上——而合计仍然是对的，
            # 所以从合计根本看不出来。
            "subject": t["subject"] or (p and p["subject"]) or "",
            "bankName": t["bank"] or (p and p["bank"]) or "",
            "acctNo": (p.get("acctNo", "") if p else ""),
            "currency": cur_code(t.get("currency") or (p and p.get("currency"))) or "CNY",
            "orig": t["orig"], "local": t["local"],
            "nature": p["nature"] if p else "",
            "confirm": p["confirm"] if p else "",
            "notes": p["notes"] if p else "",
            "priorLocal": p["local"] if p else None,
            "priorOrig": p["orig"] if p else None,
            "priorBank": p["bank"] if p else "",
            "priorAcct": p.get("acctNo", "") if p else "",
            "priorCompany": p["company"] if p else "",
            "_how": how or "新增",
            "_matched": bool(p),
        }
        rows.append(r)
        if not p:
            if not cash_row:          # 库存现金不算「本年新增」
                news.append(r)
                nonexact.append(r)
        elif how == "账号":
            nonexact.append(r)        # 账号一致但银行名写法变了，值得单独列示

    miss = [p for p in pri if p["_i"] not in used]

    # 上期有、本期没有的账户也要进 BKD（金额 0），否则「本期账面无此账号」这个
    # 审计关注点会从表上消失——最终底稿的 126 行里就有 38 个这类/零余额账户。
    for p in miss:
        rows.append({
            "company": p["company"], "subject": p["subject"], "bankName": p["bank"],
            "acctNo": p.get("acctNo", ""),
            "currency": cur_code(p.get("currency")) or "CNY",
            "orig": 0.0, "local": 0.0,
            "nature": p["nature"] if p else "",
            # 本期账面无余额：函证与否留给项目组判断，程序不预填（见 _fill_bkd 同处说明）
            "confirm": "",
            "notes": "本期账面无此账号，如已销户请补充销户证明",
            "priorLocal": p["local"], "priorOrig": p["orig"],
            "priorBank": p["bank"], "priorAcct": p.get("acctNo", ""),
            "priorCompany": p["company"],
            "_how": "本期无", "_matched": True, "_priorOnly": True,
        })

    changes = []
    for r in rows:
        if r.get("_priorOnly"):
            continue            # 本期无此账号单独在检查表列示，不混进大额变动
        if not r["_matched"] or r["priorLocal"] is None:
            continue
        d = r2(r["local"] - r["priorLocal"])
        base = abs(r["priorLocal"])
        rate = d / base if base >= 0.005 else (1 if abs(d) > 0.005 else 0)
        if abs(d) >= amt_threshold and abs(rate) >= rate_threshold:
            changes.append({"r": r, "delta": d, "rate": rate})
    changes.sort(key=lambda x: -abs(x["delta"]))

    sum_tb = r2(sum(r["local"] for r in rows))
    # 统计口径用 pri_all：库存现金虽然不参与比对，但"上年底稿有多少户、合计多少"要照实算
    sum_pri = r2(sum(p["local"] for p in pri_all))

    stats = {
        "priorCount": len(pri_all), "priorRawCount": len(prior), "dupPrior": dup,
        "tbCount": len(rows), "matched": matched,
        "newCount": len(news), "missCount": len(miss),
        "sumTb": sum_tb, "sumPrior": sum_pri, "diff": r2(sum_tb - sum_pri),
        # 两年比对口径（2026-11-07 起）：主键 = 辅助核算项1 整串
        "byKey": sum(1 for r in rows if r["_how"] == "唯一主键"),
        "byAcct": sum(1 for r in rows if r["_how"] == "账号"),
        "bySubject": sum(1 for r in rows if r["_how"] == "科目"),
        "priorOnlyCount": sum(1 for r in rows if r.get("_priorOnly")),
    }
    # 兼容旧前端字段名：exact/onlyAcct/byBank 分别对应主键/账号/（已废弃的按银行名）
    stats["exact"] = stats["byKey"]
    stats["onlyAcct"] = stats["byAcct"]
    stats["byBank"] = 0
    return {"rows": rows, "news": news, "miss": miss, "nonexact": nonexact,
            "changes": changes, "stats": stats, "prior": pri}


# ============================================================ 生成 Excel
DEFAULT_TPL_COLS = [
    ("公司名称", 26), ("科目名称", 20), ("银行/存款机构名称", 24), ("账号1", 24), ("币种", 8),
    ("期末账面金额\n（原币）", 16), ("期末账面金额\n（本币）", 16), ("货币资金账户的性质和用途", 22),
    ("是否对账户执行银行函证程序？", 12), ("记录我们判断…替代性程序。", 34), ("Notes", 26),
]

#: 金额统一格式：Excel「会计专用 · 无货币符号」——两位小数、千分位、
#: 零显示为「-」、负数用括号（Excel 内置「会计专用」的规范写法）、数字右留一位对齐。
#: 为什么不用 "#,##0.00"：零会显示成 0.00，底稿里"零余额"与"没填"看起来一样；
#: 为什么不用 "[Red]-"：负数变红，与模板自带的会计格式并存时一列两色。
ACCT_FMT = '_ * #,##0.00_ ;_ * (#,##0.00)_ ;_ * "-"??_ ;_ @_ '
MONEY_FMT = ACCT_FMT


def _resolve_tpl_cols(tpl_headers, tpl_widths):
    """根据模板表头决定列语义，返回 [(字段, 表头, 宽), ...]。"""
    if not tpl_headers:
        cols = [(None, h, w) for h, w in DEFAULT_TPL_COLS]
    else:
        cols = []
        for i, h in enumerate(tpl_headers):
            w = tpl_widths[i] if i < len(tpl_widths) else None
            cols.append((None, h, w))
    # 映射字段
    mp = map_columns([h for _, h, _ in cols], BKD_SPEC)
    rev = {}
    for k, v in mp.items():
        for i in (v if isinstance(v, list) else [v]):
            rev.setdefault(i, k)
    result = []
    for i, (_, h, w) in enumerate(cols):
        result.append((rev.get(i), h, w))
    return result


def _find_tpl():
    """返回当前使用的模板路径。优先用 _tpl 下用户最新替换的模板，否则用内置 V7。"""
    if os.path.exists(TPL_DEFAULT):
        return TPL_DEFAULT
    # 兜底：找 _tpl 下任意 xlsx
    for f in sorted(os.listdir(TPL_DIR)):
        if f.lower().endswith((".xlsx", ".xlsm")):
            return os.path.join(TPL_DIR, f)
    return None


# Lead 表三个一级科目名称 → Lead 数据行号（第 36 行表头，37/38/39 数据）
#: @deprecated 已由 locate_lead_rows() 运行时定位取代。仅保留给旧脚本 import，
#: 不要在新代码里用——写死行号会在客户模板多插一行时静默错位。
LEAD_ROW_BY_SUBJECT = {"库存现金": 37, "银行存款": 38, "其他货币资金": 39}

# BKD 明细表表头行 / 数据起始行 / 模板预置数据行数
BKD_HEADER_ROW = 9
BKD_DATA_START = 10
BKD_PRESET_ROWS = 6          # 模板预置 6 行数据位（第 10-15 行）

# ---- BKD 数据区的统一外观 ----
# 同一张表里字体、边框、对齐、数字格式必须逐格一致。以前只有"这一格有值"的
# 单元格才刷样式，空列（记录判断/Notes）就留着模板默认的等线 11 + 无边框，
# 一行里出现两种字号、两种边框，看着像两张表拼起来的。
_BKD_FONT = Font(name="Arial", size=10)
_BKD_FONT_B = Font(name="Arial", size=10, bold=True)
_BKD_ALIGN_L = Alignment(horizontal="left", vertical="center")
_BKD_ALIGN_R = Alignment(horizontal="right", vertical="center")
_BKD_ALIGN_C = Alignment(horizontal="center", vertical="center")
_BKD_BORDER = Border(left=Side(style="thin", color="808080"),
                     right=Side(style="thin", color="808080"),
                     top=Side(style="thin", color="808080"),
                     bottom=Side(style="thin", color="808080"))
_BKD_NOFILL = PatternFill(fill_type=None)


def _clear_input_fill(cell):
    """清掉模板预置的输入底色（浅黄硬底纹）。

    模板把「要人填/要人确认」的格铺了浅黄，这是**静态填充**，不是条件格式。
    程序一旦往格里写了公式或数，这格就不该再提示"待填"——否则填了数还黄，
    看着像漏填，与事实相反。只清程序确实写了值的格；仍留空、仍靠人工判断的格
    保持模板原样（清掉等于替审计师表示"这格我确认过了"）。
    """
    cell.fill = _BKD_NOFILL
    return cell
#: BKD 表格最外侧那圈深蓝（模板与费用底稿同色）
BKD_FRAME_BLUE = "FF333399"

#: 列类型 → (数字格式, 对齐)。金额列一律 MONEY_FMT（= ACCT_FMT，会计专用无货币符号），
#: 不能再出现「有的 '#,##0.00'、有的 '[Red]-#,##0.00'」——同表两种「零」的写法，
#: 复核时 0 与「没填」分不清（2026-11-12 统一）。
_BKD_KIND_STYLE = {
    "money": (MONEY_FMT, _BKD_ALIGN_R),
    "rate": ("0.0000", _BKD_ALIGN_R),
    "acct": ("@", _BKD_ALIGN_L),
    "text": ("General", _BKD_ALIGN_L),
}

#: 「两年账户识别结果」列：表头文字与逐户结论（按 merge 的 `_how` 翻译成人话）。
#: 结论只写"这两年怎么对上的/有没有"，不写金额，金额在左右两列。
IDRES_HEAD = "两年账户识别结果"
IDRES_TEXT = {
    "唯一主键": "上年已匹配（唯一主键一致）",
    "账号": "上年已匹配（账号一致）",
    "科目": "上年已匹配（按公司+科目）",
    "新增": "本年新增（上年无此账户）",
    "本期无": "本期账面无此账号（上年有）",
    "不适用": "",            # 库存现金：没有银行/账号，不参与两年比对
}
#: 旧版塞在 Notes 里的识别标记。上年底稿里也带着这些字样，回捞时会一起抄过来，
#: 写 Notes 前必须剥掉——识别结果已经单独成列，留在 Notes 里就是同一句话说两遍。
_ID_TAG_RE = re.compile(
    r"【(?:新增[·•][^】]*|账号取自上年[^】]*)】")


def _norm_subject_l1(subject):
    """把科目名称归类到「库存现金/银行存款/其他货币资金」三个一级科目。"""
    s = T(subject).replace(" ", "")
    if not s:
        return ""
    # 先精确
    if s in ("库存现金", "现金"):
        return "库存现金"
    if s in ("银行存款", "银行"):
        return "银行存款"
    if s in ("其他货币资金",):
        return "其他货币资金"
    # 再前缀/包含
    if "现金" in s:
        return "库存现金"
    if "银行" in s:
        return "银行存款"
    if ("其他货币" in s or "货币资金" in s or "存出投资款" in s or "外埠" in s
            or "信用卡" in s or "保证金" in s or "微信" in s or "支付宝" in s
            or "银行本票" in s or "银行汇票" in s or "信用证" in s or "存出" in s):
        return "其他货币资金"
    return ""


def _fill_derived(wb, res, meta=None):
    """派生表入口。延迟 import，避免 derived ↔ core 的循环引用。

    双路 import：生产环境走包（`from . import derived`），但测试脚本会把
    `程序/modules/bkd` 直接加到 sys.path 后 `import core`，那时没有父包，
    相对 import 会抛 "attempted relative import with no known parent package"。
    """
    try:
        from . import derived as DRV
    except ImportError:                       # pragma: no cover - 脚本直跑的情况
        import derived as DRV
    return DRV.fill(wb, res, meta)


def build_workbook(res, tpl_headers=None, tpl_widths=None, meta=None, highlight_new=True, mark_new=True,
                   cutoff=None):
    """基于 EY 模板生成：保留模板全部 sheet，往 C.00 BKD 填本年明细、往 C.00 Lead 打链接公式、
    往 C.03 Cutoff 填截止性测试样本，追加检查表。

    cutoff：find_transfers() 的返回值；为 None 时 C.03 保持模板原样。
    返回 (bytes, tpl_path 或 None)。
    """
    meta = meta or {}
    tpl_path = _find_tpl()
    if not tpl_path:
        return None, None
    wb = load_workbook(tpl_path)

    # ---- 共用基础信息先落封面 ----
    # 顺序很关键：C4/C5 写进去之后，下面 read_te_sad 才读得到真 TE；
    # 反过来的话 TE 永远走兜底 100000，认定阈值、截止阈值、波动判断集体失准。
    try:
        try:
            from . import baseinfo as BI          # 作为包导入
        except ImportError:
            import baseinfo as BI                 # 脚本模式：sys.path 里直接有 modules/bkd
        given = {k: meta.get(k) for k in ("customer", "balance_date", "standard",
                                          "currency", "te", "sad")
                 if meta.get(k) not in (None, "")}
        _info = BI.merged(KZ_DATA, **given)
        # 关键信息与 CRA 都只在平台侧（project_info.json）维护，**不往底稿里插配置表**：
        # 底稿是交付物，多两张内部配置表，复核人还要问这是干什么的。
        # CRA 百分比在生成时直接写进公式（见 _fill_lead_cra），底稿里看不到中间表。
        res["craCfg"] = BI.cra_of(_info)
        BI.fill_lead_header(wb, _info, overwrite=False)      # 先填空格
        if given:
            meta["baseInfo"] = BI.fill_lead_header(
                wb, {k: _info[k] for k in given}, overwrite=True)
        else:
            meta["baseInfo"] = BI.fill_lead_header(wb, _info)
    except Exception as e:
        meta["baseInfo"] = {"error": "%s: %s" % (type(e).__name__, str(e)[:140])}

    # 阈值从工作簿读（Skill §3：不写死）。模板没填时 TE_DEFAULT 与旧硬编码一致，
    # 所以“没填 TE 的项目”输出与改前完全相同；填了才按真值算。
    te, sad, te_info = read_te_sad(wb)
    meta["te"], meta["sad"] = te, sad
    meta["teSource"] = te_info["source"]
    # 同时写进 res：调用方持有 res，meta 常是临时字面量，出了函数就拿不到了
    res["te"], res["sad"] = te, sad
    res["teSource"] = te_info["source"]
    res["teWarnings"] = te_info["warnings"]

    # 上年底稿：作为一个独立 sheet 嵌进本年底稿（纯数值，无公式），供两年对比。
    # 必须在 _fill_bkd 之前建好——本年账号列要按银行名指向这个 sheet 的具体行。
    if res.get("prior"):
        try:
            _exact, _bybank, _byacct = embed_prior_sheet(wb, res["prior"])
            res["priorSheet"] = {"sheet": PRIOR_SHEET, "rows": len(res["prior"]),
                                 "exact": _exact, "bybank": _bybank, "byacct": _byacct}
        except Exception as e:
            res["priorSheet"] = {"error": "%s: %s" % (type(e).__name__, str(e)[:140])}

    # 上年 Lead：整表复制成本年文件里的一张 sheet（保留格式、只留数值），
    # 再由 _fill_lead 按科目名链接 K 列与表2 上年数。
    if res.get("priorLead") and not (res["priorLead"] or {}).get("error"):
        try:
            got = copy_prior_lead_sheet(wb, res["priorLead"].get("data"))
            if got.get("error"):
                res["priorLead"]["error"] = got["error"]
            else:
                res["priorLead"].update(got)
                for w in got.get("warnings") or []:
                    res.setdefault("warnings", []).append(w)
        except Exception as e:
            res["priorLead"]["error"] = "%s: %s" % (type(e).__name__, str(e)[:140])

    _fill_bkd(wb, res, mark_new=mark_new, highlight_new=highlight_new)
    _fill_lead(wb, res)
    _fill_cutoff(wb, cutoff)
    # 派生表（C.01b / C.00b / C.02 / C.02a）：模板里有就填，没有就记理由跳过
    res.setdefault("derived", {})
    res["derived"] = _fill_derived(wb, res, meta)

    # 追加检查表（保留原有检查表逻辑；若已存在同名则替换）
    if "检查" in wb.sheetnames:
        del wb["检查"]
    _build_check_sheet(wb, res, meta, rate_threshold=0.5, amt_threshold=te, cutoff=cutoff)
    if res.get("gaps"):
        _add_gap_section(wb["检查"], res)

    # 上年两张对照表默认隐藏（2026-11-08 用户定）：它们是追溯用的底稿，
    # 不是给人日常翻的表；右键「取消隐藏」仍可查看，跨表公式照常计算。
    _hidden = []
    for _n in (PRIOR_SHEET, PRIOR_LEAD_SHEET):
        if _n in wb.sheetnames:
            wb[_n].sheet_state = "hidden"
            _hidden.append(_n)
    res["hiddenSheets"] = _hidden

    bio = io.BytesIO()
    wb.save(bio)
    data = bio.getvalue()

    # 汇率公告截图（C.02a）：内置运行时没有 Pillow，openpyxl 的 add_image 用不了，
    # 所以保存完再手工往 zip 里补图片部件。插失败就用原文件，绝不让底稿打不开。
    _sp = res.get("c02aShot")
    if not _sp:
        res["c02aShotMissing"] = True
    if _sp and os.path.isfile(_sp):
        try:
            from core import xlsx_image as XIMG
            with open(_sp, "rb") as fh:
                _png = fh.read()
            # 锚在 Notes 合并框的下一行：框本身有高度、会随币种行数上下移，
            # 写死行号会压在框上把说明文字盖住。
            _ar = int(res.get("c02aNotesEndRow") or 0) + 1 or 22
            _new = XIMG.add_png(data, "C.02a 外币测算", _png, anchor_row=_ar, anchor_col=2)
            if _new:
                data = _new
                res["c02aShotEmbedded"] = True
            else:
                # 插不进去要留痕：以前这里静默失败，底稿里没图也没人知道为什么
                res["c02aShotError"] = "add_png 返回空（图 %d 字节）" % len(_png)
        except Exception as _e:
            res["c02aShotError"] = "%s: %s" % (type(_e).__name__, str(_e)[:120])
    return data, tpl_path


def _locate_bkd_cols(ws):
    """按表头文字定位 C.00 BKD 的列，返回 {字段: 列号(1-based), "_header_row": 行号}。

    为什么要动态定位：本地 V7 模板是 B~L 十一列，瑞云实际底稿是 A~M 十三列再加班 P 列受限资金，
    两者列字母完全不同。写死列号在换一份底稿时就会整表错位，而且错得很安静。

    关键字段（公司名称 / 科目名称 / 期末本币）缺失时返回 None，让调用方报错，不降级硬写。
    """
    scan = min(ws.max_row, 40)
    rows = [[ws.cell(r, c).value for c in range(1, min(ws.max_column, 30) + 1)]
            for r in range(1, scan + 1)]
    best = detect_header(rows, BKD_SPEC, limit=scan)
    if not best["score"]:
        return None
    i = best["i"]
    heads = rows[i]
    mp = map_columns(heads, BKD_SPEC) or {}
    col = {}
    for k, v in mp.items():
        idx = v[0] if isinstance(v, list) else v
        if idx is not None:
            col[k] = idx + 1
    col["_header_row"] = i + 1
    col["_score"] = best["score"]
    # 三个字段缺一不可：公司名称、科目名称、期末本币
    if not col.get("company") or not col.get("subject") or not col.get("localEnd"):
        return None
    return col


def locate_cols(ws, hdr_row, spec, limit=40):
    """按表头文字定位列，返回 ({字段: 列号}, [没找到的字段])。

    spec 每项 ``(字段, 关键字元组[, 第几次出现])``。

    为什么要有「第几次出现」：C.03 的样本区左右对称——左边付款方、右边收款方，
    两边的表头文字**一字不差**（都是「银行名称」「银行账号」「交易金额」）。
    只按文字匹配会把收款方写进付款方列，而合计照样配平（I+M=0 仍成立），
    错误看不出来。必须靠出现次序区分。

    匹配规则：
      · 表头先去掉空格与换行（模板表头普遍带 \\n 折行）；
      · 英文字母转小写（Check / CHECK 都认）；
      · 同一列不会被两个字段占用（先到先得，按 spec 顺序）。

    找不到的字段不报错，由调用方决定哪些是必需的——必需项缺失时必须中止写入，
    否则就是「静默错位」：数字写进隔壁列，程序还返回成功。
    """
    heads = {}
    for c in range(1, min(ws.max_column, limit) + 1):
        t = T(ws.cell(hdr_row, c).value)
        t = t.replace(" ", "").replace("\n", "").replace("\u3000", "")
        if t:
            heads[c] = t.lower()
    out, used = {}, set()
    for item in spec:
        key, kws = item[0], item[1]
        nth = item[2] if len(item) > 2 else 1
        # 先把所有命中列按列序排好，再取第 nth 个。
        # 不能边遍历边数「第几次」——上一轮已经占用掉的列会被跳过，
        # 序号重新从 1 开始，第二个「银行名称」永远找不到。
        cand = [c for c in sorted(heads)
                if any(k.lower() in heads[c] for k in kws)]
        if len(cand) >= nth:
            c = cand[nth - 1]
            if c not in used:
                out[key] = c
                used.add(c)
    miss = [i[0] for i in spec if i[0] not in out]
    return out, miss


#: C.01b 特殊账户程序的列（按表头文字，不按列号）
C01B_COLS = (
    ("company", ("公司名称",), 1),
    ("subject", ("科目名称",), 1),
    ("bank", ("存款机构", "银行"), 1),
    ("acct", ("账号",), 1),
    ("currency", ("币种",), 1),
    ("orig", ("账面金额（原币）", "账面金额(原币)"), 1),
    ("local", ("账面金额（本币）", "账面金额(本币)"), 1),
    ("statement", ("对账单", "资产证明"), 1),
    ("diff", ("差异",), 1),
    ("nature", ("账户类型",), 1),
    ("restricted", ("实际受限金额", "受限金额"), 1),
    ("reason", ("记录受限原因",), 1),
)
#: 缺了就不许写：公司 / 科目 / 原币 / 本币 / 差异
C01B_REQUIRED = ("company", "subject", "orig", "local")

#: C.02a 外币测算的列
C02A_COLS = (
    ("currency", ("原币币种",), 1),
    ("orig", ("原币金额",), 1),
    ("rate", ("折算汇率", "汇率"), 1),
    ("conv", ("折合本位币",), 1),
    ("book", ("账面本位币", "账面人民币"), 1),
    ("diff", ("差异",), 1),
)
C02A_REQUIRED = ("currency", "orig", "rate", "conv", "book", "diff")

#: C.03 样本区的列。左右两半表头同字，靠第 1/第 2 次出现区分付款方 / 收款方。
C03_COLS = (
    ("no", ("样本编号",), 1),
    ("company", ("公司名称",), 1),
    ("payBank", ("银行名称",), 1),
    ("payAcct", ("银行账号",), 1),
    ("payTime", ("收款/付款时间", "付款时间"), 1),
    ("vno", ("交易编号",), 1),
    ("currency", ("币种",), 1),
    ("payAmt", ("交易金额",), 1),
    ("recvBank", ("银行名称",), 2),
    ("recvAcct", ("银行账号",), 2),
    ("recvDate", ("交易日期",), 1),
    ("recvAmt", ("交易金额",), 2),
    ("cross", ("跨期",), 1),
    ("note", ("备注",), 1),
    ("check", ("check",), 1),
)
C03_REQUIRED = ("company", "payBank", "payAcct", "payAmt",
                "recvBank", "recvAcct", "recvAmt", "check")


FX_RATE_HEAD = ("折算率", "折算汇率")
FX_RMB_KEY = "期末账面金额（人民币）"
#: 认得的「折合人民币」表头写法（旧文件/客户模板可能用别的名）
FX_RMB_HEADS = ("期末账面金额（人民币）", "期末账面金额(人民币)", "折算人民币", "折合人民币")
RESTRICTED_HEAD = ("受限资金", "受限金额", "本期末实际受限金额")

#: 读不到 TE 时的兜底。100,000 是旧版硬编码值，保持它能让“模板没填 TE”的项目
#: 输出与以前完全一致；填了 TE 的模板才会用真值。
TE_DEFAULT = 100000.0
#: SAD = TE × 该比例（EY GAM 的常规口径，模板 D8 也是这么算的）
SAD_RATIO = 0.1


def read_te_sad(wb):
    """读 TE（可容忍误差）与 SAD（名义金额），返回 (te, sad, info)。

    读的顺序（Skill §3 阈值来源：阈值不写死，从工作簿取）：
      1. PMTE 表的 B5/B6 —— 项目主数据表，TE 的唯一权威来源
      2. C.00 Lead!C4/C5 —— 模板自带的阈值输入位（D14 等公式引用的就是 C4）
      3. 都没有 → TE_DEFAULT + `TE_NOT_READ` 警告，不静默用错值

    为什么必须读而不能写死：TE 一变，认定层阈值、截止性测试阈值、变动调查阈值
    全都跟着变。硬编码会让整份底稿的“是否进一步调查”判断集体失准，而且看不出来。
    """
    info = {"source": "", "warnings": []}
    te = sad = None
    # 1) PMTE
    for name in wb.sheetnames:
        if "PMTE" in name.upper().replace(" ", ""):
            ws = wb[name]
            try:
                v = ws.cell(5, 2).value          # B5
                s = ws.cell(6, 2).value          # B6
                if isinstance(v, (int, float)) and v > 0:
                    te = float(v)
                    sad = float(s) if isinstance(s, (int, float)) and s > 0 else None
                    info["source"] = "PMTE!B5"
                    break
            except Exception:
                pass
    # 2) Lead 的阈值输入位
    if te is None and "C.00 Lead" in wb.sheetnames:
        lead = wb["C.00 Lead"]
        v = lead.cell(4, 3).value                # C4 = TE
        s = lead.cell(5, 3).value                # C5 = SAD
        if isinstance(v, (int, float)) and v > 0:
            te = float(v)
            if isinstance(s, (int, float)) and s > 0:
                sad = float(s)
            info["source"] = "C.00 Lead!C4"
    # 3) 兜底
    if te is None:
        te = TE_DEFAULT
        info["source"] = "default"
        info["warnings"].append(
            "TE_NOT_READ：模板里没有 PMTE 表，C.00 Lead!C4 也是空的，"
            "暂按默认 TE=%s 计算阈值。请在模板里填可容忍误差后重跑。" % format(TE_DEFAULT, ",.2f"))
    if sad is None:
        sad = round(te * SAD_RATIO, 2)
        info["warnings"].append("SAD 未填，按 TE×%d%% 推算 = %s" % (SAD_RATIO * 100, format(sad, ",.2f")))
    info["te"], info["sad"] = r2(te), r2(sad)
    return r2(te), r2(sad), info


#: 第三方支付/收款平台。这类账户没有银行账号（账号是平台商户号/绑定邮箱），
#: 底稿的「账号」栏应当写「—」而不是留空——留空会被当成"缺资料"，让客户去补一个不存在的东西。
PLATFORM_KEYS = ("支付宝", "微信", "财付通", "银联", "PAYPAL", "PAYPAL",
                 "STRIPE", "空中云汇", "连连", "PINGPONG", "万里汇", "第三方",
                 "快钱", "汇付", "易宝", "平台", "店铺", "PAYONEER", "WISE")


def is_platform_acct(bank, name=""):
    """这个账户是不是第三方支付/收款平台（而不是银行账户）。"""
    txt = (T(bank) + " " + T(name)).upper()
    return any(k.upper() in txt for k in PLATFORM_KEYS)


def _locate_restricted_col(ws, hdr_row):
    """定位（或分配）BKD 的「受限资金」列，返回列号。

    金额写在 BKD（每笔一次），Lead 再从「其他货币资金」重分类出来列示——
    两处不能重复计数。列接在模板已用列最右侧、折算列之前，
    与 `_locate_fx_cols` 的分配顺序配合（先受限、后折算），列位置才稳定。
    """
    top = min(ws.max_column, 200)
    for c in range(1, top + 1):
        h = T(ws.cell(hdr_row, c).value).replace(" ", "").replace("\n", "")
        if any(k in h for k in RESTRICTED_HEAD) and "受限资产" not in h:
            return c
    last_used = 1
    for r in (max(1, hdr_row - 1), hdr_row):
        for c in range(1, top + 1):
            if T(ws.cell(r, c).value):
                last_used = max(last_used, c)
    return last_used + 1


def _locate_fx_cols(ws, hdr_row):
    """查已有的「折算率」「期末账面金额（人民币）」两列，返回 (列号, 列号)。

    只查不建。需要"没有就建"时用 `_ensure_fx_cols`。
    表头匹配用「人民币」而不是旧的「折算人民币」——列名改成"期末账面金额（人民币）"后，
    只认"折算"会漏掉新表头，于是把它当成"没有折算列"，退回去用本币列，静默少一次折算。
    """
    top = min(ws.max_column, 200)
    c_rate = c_rmb = None
    for c in range(1, top + 1):
        h = T(ws.cell(hdr_row, c).value).replace(" ", "").replace("\n", "")
        if not h:
            continue
        if h in FX_RATE_HEAD or any(k in h for k in FX_RATE_HEAD):
            c_rate = c
        elif any(k in h for k in FX_RMB_HEADS) or ("人民币" in h and "金额" in h):
            c_rmb = c
    return c_rate, c_rmb


def _col_index(letter):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


def _clone_style(ws, row, src_col, dst_col):
    """把同一行 src_col 的样式复制到 dst_col（只复制格式，不复制值）。"""
    if not row or not src_col or not dst_col or src_col == dst_col:
        return
    try:
        ws.cell(row, dst_col)._style = _copy_style(ws.cell(row, src_col)._style)
    except Exception:
        pass


def _style_header_like(ws, row, col, text, width=None, horizontal="center"):
    """写一个"新增列"的表头格：样式整格照抄邻居，只改文字。

    为什么必须统一走这里：新增列（折算率 / 期末账面金额（人民币）/ 受限资金）的表头
    原先各写各的，`cell.alignment = _BKD_ALIGN_C` 这类写法会**整体替换**对齐对象，
    把邻居带过来的 wrap_text 一并丢掉——表头行的字体、底色、边框都对，唯独这几格不换行，
    肉眼很难判断是代码问题，还以为是模板本来就这样。

    做法：font / fill / border / alignment / number_format 全部沿邻居，
    最后强制 wrap_text=True（表头一律允许换行，这是行高 43.9 的前提）。
    """
    cell = ws.cell(row, col, text)
    src = None
    for d in (-1, 1, -2, 2):
        c = col + d
        if c < 1 or c > (ws.max_column or 1):
            continue
        if T(ws.cell(row, c).value):
            src = c
            break
    if src:
        s = ws.cell(row, src)
        cell.font = _copy_style(s.font)
        cell.fill = _copy_style(s.fill)
        cell.border = _copy_style(s.border)
        cell.number_format = s.number_format
        al = _copy_style(s.alignment)
    else:
        al = Alignment(horizontal="center", vertical="center")
    al.horizontal = horizontal or al.horizontal
    al.vertical = al.vertical or "center"
    al.wrap_text = True
    cell.alignment = al
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return cell


def _insert_cols_keep_layout(ws, idx, n):
    """在 idx 前插 n 列，并把列宽 / 合并区 / 样式一起搬过去。

    为什么不能只调 openpyxl 的 insert_cols：它只搬单元格的值和样式，
    **不搬列宽、不动合并区、不管公式里的列引用**。插两列之后会出现三种
    "整张表被拆开"的现象：

      1. 列宽整体错位 —— 42.8 宽的「记录判断」列宽留在「性质」列上，
         真正的记录判断只剩 2.6、Notes 只剩 1.6，文字全挤成一条竖线；
      2. 标题横幅（第 1 行深色底）在插入位置断成两截，中间一道白缝；
      3. 新列是裸单元格，没有边框没有底色，读起来像两张表拼在一起。

    这三件必须一起修，少修一件都还是难看。
    """
    try:
        from openpyxl.cell.cell import MergedCell
    except Exception:                                     # pragma: no cover
        MergedCell = ()

    # 1) 备份 idx 及右侧的列宽
    widths = []
    for key, dim in list(ws.column_dimensions.items()):
        try:
            ci = _col_index(key)
        except Exception:
            continue
        if ci >= idx and dim.width:
            widths.append((ci, dim.width, dim.hidden, dim.bestFit))

    # 2) 备份跨过插入点 / 位于其右侧的合并区（openpyxl 不会动它们）
    merges = []
    for mc in list(ws.merged_cells.ranges):
        if mc.max_col >= idx:
            merges.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
            ws.unmerge_cells(str(mc))

    # 3) 样式源 = 左侧相邻列（横幅底色、表头带、数据行边框都靠它传过去）
    src = idx - 1
    snap = {}
    for r in range(1, ws.max_row + 1):
        try:
            snap[r] = _copy_style(ws.cell(r, src)._style)
        except Exception:
            pass

    ws.insert_cols(idx, n)

    # 4) 列宽：右侧整体右移 n，新列沿用左邻列宽。
    #    注意 openpyxl 的 ColumnDimension.width 不接受 None（会抛 TypeError），
    #    清空必须用 del —— 否则旧宽度留在原地，插完列会出现"两列都 42.8 宽"。
    for (ci, _w, _h, _b) in widths:
        key = get_column_letter(ci)
        if key in ws.column_dimensions:
            try:
                del ws.column_dimensions[key]
            except Exception:
                pass
    for (ci, w, hidden, best) in widths:
        d = ws.column_dimensions[get_column_letter(ci + n)]
        d.width = w
        if hidden:
            d.hidden = True
        if best:
            d.bestFit = True
    src_w = next((w for (ci, w, _h, _b) in widths if ci == src), None)
    for k in range(n):
        d = ws.column_dimensions[get_column_letter(idx + k)]
        if src_w:
            d.width = src_w

    # 5) 合并区：跨过插入点的扩宽，右侧的右移
    for (min_r, max_r, min_c, max_c) in merges:
        if min_c < idx <= max_c:
            max_c += n
        else:
            min_c += n
            max_c += n
        ws.merge_cells(start_row=min_r, start_column=min_c,
                       end_row=max_r, end_column=max_c)

    # 6) 样式复制（值清空，表头文字由调用方写）
    for r, st in snap.items():
        if r > ws.max_row:
            continue
        for k in range(n):
            cell = ws.cell(r, idx + k)
            if MergedCell and isinstance(cell, MergedCell):
                continue
            cell._style = _copy_style(st)
            cell.value = None


def _ensure_fx_cols(ws, hdr_row, money_col):
    """确保「折算率」「期末账面金额（人民币）」紧跟在本币列右侧。返回 (率, 人民币, 是否插了列)。

    为什么必须紧跟本币列而不是放最右：BKD 的读数顺序是「原币 → 本币 → 折算率 → 人民币」，
    折算是本币的下一跳。放到最右边、隔着性质/函证/Notes/受限，看一行数要横向来回跳，
    而且很容易把折算列当成某个说明列漏掉。

    **插列会让原来 I 列之后的列全部右移，调用方必须在插入后重新定位列号。**
    """
    c_rate, c_rmb = _locate_fx_cols(ws, hdr_row)
    if c_rate and c_rmb and c_rate > money_col:
        return c_rate, c_rmb, False
    # 没有（或位置不对）→ 在本币列后面插两列。模板每次重新载入，所以这里一定是"没有"。
    _insert_cols_keep_layout(ws, money_col + 1, 2)
    return money_col + 1, money_col + 2, True


def _rate_of(row, default=1.0):
    """取该行的折算率；没给或给了非数字都按 1 处理（本位币即人民币）。"""
    v = row.get("rate")
    try:
        f = float(v)
    except (TypeError, ValueError):
        return default
    return f if f > 0 else default


# ============================================================ 公司mapping（两列汇率）
#: 客户给的「公司维度的两列汇率」表：公司简称 + 余额汇率 + 损益汇率。
#: 口径：**余额类科目**（货币资金/借款/应收应付）用余额汇率，**损益类**（费用/收入/成本）用损益汇率。
#: 这张表由客户原样提供，平台只做两件事：把它原封不动放进工作簿、用 XLOOKUP 取数。
MAP_SHEET_NAME = "公司mapping"
_MAP_FILE_HINTS = ("mapping", "对照", "汇率")
#: 表头认法：余额汇率 / 损益汇率各有几种常见写法，按关键词认，不按列号
_MAP_BAL_WORDS = ("余额汇率", "期末汇率", "资产负债表日汇率", "时点汇率")
_MAP_PL_WORDS = ("损益汇率", "平均汇率", "期间汇率", "利润表汇率")
#: 损益类科目的名称特征（货币资金底稿里用不到，留给同结构的其它底稿）
_PL_SUBJECT_WORDS = ("费用", "收入", "成本", "税金", "支出", "收益")


def _map_sheet_of(wb):
    """工作簿里已有的 mapping 页（客户自己贴好的那种），没有返回 None。"""
    for n in wb.sheetnames:
        low = str(n).lower()
        if "mapping" in low or "对照" in low or "映射" in low:
            return wb[n]
    return None


def _map_header(ws):
    """在 mapping 页里认表头：简称 / 全称 / 余额汇率 / 损益汇率。

    表头行不写死：客户常在第 1~5 行放标题与说明，真正的列名在中间。
    """
    lim = min(ws.max_row or 0, 30)
    for r in range(1, lim + 1):
        texts = [T(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 30) + 1)]
        if not any("汇率" in t for t in texts):
            continue
        col = {}
        for c, t in enumerate(texts, start=1):
            tt = t.replace(" ", "")
            if not tt:
                continue
            if "简称" in tt and "short" not in col:
                col["short"] = c
            elif ("全称" in tt or "公司名称" in tt or "主体名称" in tt) and "full" not in col:
                col["full"] = c
            elif "汇率" in tt and any(w in tt for w in _MAP_BAL_WORDS) and "bal" not in col:
                col["bal"] = c
            elif "汇率" in tt and any(w in tt for w in _MAP_PL_WORDS) and "pl" not in col:
                col["pl"] = c
        if col.get("bal") or col.get("pl"):
            col["hdr"] = r
            col["first"] = r + 1
            col["last"] = max(ws.max_row or 0, r + 1)
            return col
    return None


def _find_map_file(data_dir):
    """数据目录里的 mapping 文件（文件名带 mapping/对照/汇率）。

    先看根目录，再看 uploads/ —— 平台上传的原始文件落在 uploads/，
    客户那张「公司mapping」通常就混在那里。
    """
    for base in (data_dir, os.path.join(data_dir or "", "uploads")):
        try:
            names = sorted(os.listdir(base))
        except OSError:
            continue
        for n in names:
            low = n.lower()
            if not low.endswith((".xlsx", ".xlsm")) or n.startswith("~$"):
                continue
            if any(h in low for h in _MAP_FILE_HINTS) or any(h in n for h in _MAP_FILE_HINTS):
                return os.path.join(base, n)
    return ""


def ensure_company_mapping(wb, data_dir=None):
    """确保工作簿里有「公司mapping」页，返回列信息；拿不到就 None（折算率退回默认 1）。

    来源优先级：
      1) 工作簿/模板里已有 mapping 页 —— 原样用，一格不改；
      2) 数据目录里有 mapping/对照/汇率 的 xlsx —— **原封不动**复制成一个新页（值 + 数字格式）；
      3) 都没有 —— 返回 None，调用方记 warning，不猜汇率。
    """
    ws = _map_sheet_of(wb)
    if ws is not None:
        info = _map_header(ws)
        if info:
            info.update({"sheet": ws.title, "source": "已有页", "rows": ws.max_row})
            return info
    path = _find_map_file(data_dir or KZ_DATA)
    if not path:
        return None
    try:
        src = load_workbook(path, data_only=True)
    except Exception:
        return None
    sws = src[src.sheetnames[0]]
    dst = wb.create_sheet(MAP_SHEET_NAME)
    for r in range(1, (sws.max_row or 0) + 1):
        for c in range(1, (sws.max_column or 0) + 1):
            sc = sws.cell(r, c)
            dc = dst.cell(r, c, sc.value)
            if sc.has_style:
                dc.number_format = sc.number_format
    for k, dim in list(sws.column_dimensions.items())[:30]:
        if dim.width:
            dst.column_dimensions[k].width = dim.width
    info = _map_header(dst) or {}
    info.update({"sheet": dst.title, "source": os.path.basename(path),
                 "rows": sws.max_row or 0})
    return info or None


def _rate_kind_of(subject):
    """这个科目该用哪一列汇率：余额类 → bal，损益类 → pl。"""
    s = T(subject)
    if _norm_subject_l1(s) in ("库存现金", "银行存款", "其他货币资金"):
        return "bal"
    return "pl" if any(w in s for w in _PL_SUBJECT_WORDS) else "bal"


def _find_sum_row(ws, data_start, money_col):
    """定位合计行：在 money_col 上找第一个 =SUM(...) 公式行。"""
    for r in range(data_start, min(ws.max_row, data_start + 500) + 1):
        v = ws.cell(r, money_col).value
        if isinstance(v, str) and v.strip().upper().startswith("=SUM("):
            return r
    return None


def _paint_outer_frame(ws, max_scan=400):
    """把模板自带的「右竖边 + 底横边」深蓝装饰带（FF333399）补齐。

    适用 BKD 与 Lead：两张表都是这个画法。会断的原因有两类——
      · BKD：明细区有一条"一律白底"的清理循环，把底横边连同右竖边一起刷白；
      · Lead：插行（定期存款 / 大额存单 / 受限资金等）时新行不继承邻行样式，
        右竖边在插入的那几行上出现缺口，看上去就是"蓝边被截断"。
    位置不写死：底边认「第 1 列被涂成深蓝的那一行」（插行后自动下移），
    右竖边认「第 1 行最右边那格蓝色」。模板本身没有蓝框就什么都不画（不无中生有）。
    """
    blue = PatternFill("solid", fgColor=BKD_FRAME_BLUE)

    def _is_blue(cell):
        f = cell.fill
        return bool(f and f.fill_type
                    and str(getattr(f.fgColor, "rgb", None) or "").upper().endswith("333399"))

    bottom = None
    for r in range(1, min(ws.max_row, max_scan) + 1):
        if _is_blue(ws.cell(r, 1)):
            bottom = r
            break
    right = None
    for c in range(min(ws.max_column, 60), 0, -1):
        if _is_blue(ws.cell(1, c)):
            right = c
            break
    if bottom is None and right is None:
        return None
    bottom = bottom or min(ws.max_row, max_scan)
    right = right or min(ws.max_column, 60)

    for r in range(1, bottom + 1):
        if not _is_blue(ws.cell(r, right)):
            ws.cell(r, right).fill = blue
    for c in range(1, right + 1):
        if not _is_blue(ws.cell(bottom, c)):
            ws.cell(bottom, c).fill = blue
    return right, bottom


#: 旧名保留（BKD 侧调用点较多），行为完全一致
_paint_bkd_frame = _paint_outer_frame


def _fill_bkd(wb, res, mark_new=True, highlight_new=True):
    """往 C.00 BKD 明细表填本年数据。"""
    if "C.00 BKD" not in wb.sheetnames:
        return
    ws = wb["C.00 BKD"]
    rows = res["rows"]

    # ---- 动态定位表头与合计行（不写死行号列号）----
    col = _locate_bkd_cols(ws)
    if not col:
        raise RuntimeError(
            "C.00 BKD 表头识别失败：找不到「公司名称 / 科目名称 / 期末账面金额（本币）」这三列。"
            "当前模板可能不是 C SWP 货币资金底稿，请先在页面上「上传替换模板」再生成。")
    # 折算两列紧跟在本币列右侧（原币 → 本币 → 折算率 → 人民币）。
    # 插列会把原 I 列之后全部右移，所以必须重定位——否则性质/函证/Notes 会写到错的列上。
    _c_rate0, _c_rmb0, _inserted = _ensure_fx_cols(ws, col["_header_row"], col["localEnd"])
    if _inserted:
        col = _locate_bkd_cols(ws) or col
        col["_fx_rate"], col["_fx_rmb"] = _c_rate0, _c_rmb0
    data_start = col["_header_row"] + 1
    sum_row = _find_sum_row(ws, data_start, col["localEnd"])

    # 模板预置数据位：合计行与起始行之间，最后一行若是长文本说明行则不算数据位
    preset, insert_at = BKD_PRESET_ROWS, data_start + BKD_PRESET_ROWS
    if sum_row:
        preset = max(1, sum_row - data_start)
        insert_at = sum_row
        head_txt = T(ws.cell(sum_row - 1, col["company"]).value)
        if len(head_txt) > 12:          # 合计行上面那行是说明文字，不是数据行
            preset -= 1
            insert_at = sum_row - 1

    # 需要的数据行数 vs 模板预置行数
    need = len(rows)
    if need > preset:
        n_insert = need - preset
        # openpyxl 的 insert_rows 不会移动合并单元格，需手动下移第 16 行及以下的合并单元格，
        # 否则 Notes 填写框（B22:F24 等）会错位覆盖数据。
        shifted_merges = []
        for mc in list(ws.merged_cells.ranges):
            if mc.min_row >= insert_at:
                shifted_merges.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
                ws.unmerge_cells(str(mc))
        ws.insert_rows(insert_at, n_insert)
        for (min_r, max_r, min_c, max_c) in shifted_merges:
            ws.merge_cells(start_row=min_r + n_insert, start_column=min_c,
                           end_row=max_r + n_insert, end_column=max_c)

    # 表头样式取自模板第 10 行（数据首行），逐列复制
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    # 明细区不再用底色区分「新增」和「缺资料」（用户要求整张表白底，见函数末尾的清理循环）。
    # 待补信息照旧逐格记进 gaps，落在「检查」表。
    gaps = res.setdefault("gaps", [])

    # ---- 受限资金列（人工填写）----
    # 先分配受限列、再分配折算两列，列顺序才是「…模板已用列 | 受限资金 | 折算率 | 折算人民币」。
    # 反过来写的话，折算列会先占住最右位，受限列又挤到折算列右边，看着像折算的一部分。
    c_res = _locate_restricted_col(ws, col["_header_row"])
    hdr_row = col["_header_row"]
    # 表头样式沿用左边的数据列，别自带一套浅橙——那一格会像贴上去的补丁，
    # 深色表头带在那里断一截。"这一列是系统加的"写在批注和检查表里，不靠颜色。
    res_head = _style_header_like(ws, hdr_row, c_res, "受限资金", 14)
    # 表头以上的横幅（第 1~8 行的深色条 / 索引标记带）也要补上这一格，
    # 否则整张表右上角缺一块，看着像纸被裁了一角。
    for r in range(1, hdr_row):
        _clone_style(ws, r, c_res - 1, c_res)
    res_head.comment = Comment(
        "人工填写：该账户期末受限资金金额（一笔只写一次）。\n"
        "C.00 Lead 披露「所有权或使用权受到限制的货币资金」会自动汇总本列，\n"
        "不要在 Lead 再手工加一遍，否则两处重复计数。", "PTW")

    # ---- 两年账户识别结果（紧跟受限资金右侧，程序自动填）----
    # 以前这些结论塞在 Notes 里（【新增·上年底稿无此账户】【账号取自上年 BKD…】），
    # 和项目组自己写的备注混在一列，既不好筛也不好数。单独一列，一格一个结论。
    c_idres = c_res + 1
    _insert_cols_keep_layout(ws, c_idres, 1)
    _style_header_like(ws, hdr_row, c_idres, IDRES_HEAD, 24)
    # 表头以上的横幅同样补一格，否则右上角缺一块
    for r in range(1, hdr_row):
        _clone_style(ws, r, c_res, c_idres)
    ws.cell(hdr_row, c_idres).comment = Comment(
        "程序按上年底稿逐户比对后的结论：\n"
        "  · 上年已匹配：账号+公司一致 / 按账号 / 按公司+银行 / 按公司+科目\n"
        "  · 本年新增（上年底稿无此账户）\n"
        "  · 本期账面无余额（上年有）——需核实是否销户\n"
        "未上传上年底稿时本列留空（没有可比的上年数据，不臆断新增）。", "PTW")

    # ---- 折算两列（紧跟本币列，位置由 _ensure_fx_cols 保证）----
    c_rate = col.get("_fx_rate") or (col["localEnd"] + 1)
    c_rmb = col.get("_fx_rmb") or (col["localEnd"] + 2)
    # 表头文字与左边两列同构：第一行「期末账面金额」、第二行括号币种，靠 wrap 折行
    # （行高 43.9 本来就是给两行准备的；写成一行会被列宽 18 断在「人」字上）
    for c, txt, w in ((c_rate, "折算率", 9), (c_rmb, "期末账面金额\n（人民币）", 18)):
        # 样式整格照抄左邻（含自动换行），只换文字
        _style_header_like(ws, hdr_row, c, txt, w)
    # 折算率是人工确认项：加批注，提醒本位币不是人民币时必须改
    ws.cell(hdr_row, c_rate).comment = Comment(
        "折算率来源，优先级从高到低：\n"
        "  1) 工作簿里的「公司mapping」页 —— 客户按公司维度给的两列汇率，\n"
        "     平台按「简称」XLOOKUP 到本列（余额类科目取余额汇率，损益类取损益汇率）；\n"
        "  2) 命令行 --rate 主体=汇率；\n"
        "  3) 都没有时默认 1（本位币即人民币）。\n\n"
        "注意：汇率是按**主体**定的，同一主体所有账户共用一个率。\n"
        "境内人民币主体保持 1；境外子公司（香港/新加坡）必须用当期期末汇率，\n"
        "否则「期末账面金额（人民币）」等于外币金额，C.02a 与报表披露都会差出上千万。",
        "PTW")
    Lm = get_column_letter(col["localEnd"])       # 本币列
    Lr = get_column_letter(c_rate)                # 折算率列
    Lrm = get_column_letter(c_rmb)                # 期末账面金额（人民币）列

    # ---- 公司mapping：客户给的「简称 + 余额汇率 / 损益汇率」 ----
    # 有这张表：A 列（第一列）按公司全称 XLOOKUP 出简称，折算率列再按简称 XLOOKUP 取汇率；
    # 没有：折算率沿用 _rate_of（默认 1），不猜汇率。
    # 第一列只写值，不动它的字体/数字格式（用户 2026-11-12 口径）。
    _map = ensure_company_mapping(wb, KZ_DATA)
    _map_msg = ""
    if _map:
        _ms = "'%s'" % _map["sheet"]
        _short_c = _map.get("short")
        _full_c = _map.get("full") or _short_c
        _rs, _re = _map["first"], max(_map["last"], _map["first"])
        _map_msg = ("公司mapping：已用 %s 的 %d~%d 行（来源：%s）"
                    % (_map["sheet"], _rs, _re, _map.get("source") or ""))
        if _short_c:
            # 第一列（A）表头写「简称」——只写文字，不改格式
            _a = ws.cell(hdr_row, 1)
            if not T(_a.value):
                _a.value = "简称"
            if (ws.column_dimensions[get_column_letter(1)].width or 0) < 8:
                ws.column_dimensions[get_column_letter(1)].width = 13
        # routes 把 res["warnings"] 透出成 buildWarnings，写这里才看得见
        res.setdefault("warnings", []).append(_map_msg)

    def _short_formula(r):
        """A 列：按公司全称取简称；取不到留空（不写"未匹配"占位，避免被当成结论）。"""
        return '=IFERROR(XLOOKUP($B%d,%s!$%s$%d:$%s$%d,%s!$%s$%d:$%s$%d),"")' % (
            r, _ms, get_column_letter(_full_c), _rs, get_column_letter(_full_c), _re,
            _ms, get_column_letter(_short_c), _rs, get_column_letter(_short_c), _re)

    def _rate_formula(r, kind):
        """折算率列：按简称取余额/损益汇率；取不到回 1（本位币即人民币）。"""
        rc = _map.get(kind) or _map.get("bal")
        sc = _short_c or _full_c
        return '=IFERROR(XLOOKUP($A%d,%s!$%s$%d:$%s$%d,%s!$%s$%d:$%s$%d),1)' % (
            r, _ms, get_column_letter(sc), _rs, get_column_letter(sc), _re,
            _ms, get_column_letter(rc), _rs, get_column_letter(rc), _re)

    # ---- 数据区的列类型与跨度（每一行的每一列都按这张表刷一遍）----
    kinds = {}
    for fld in ("origEnd", "localEnd"):
        if col.get(fld):
            kinds[col[fld]] = "money"
    kinds[c_rate] = "rate"
    kinds[c_rmb] = "money"
    kinds[c_res] = "money"
    kinds[c_idres] = "text"
    if col.get("acctNo"):
        kinds[col["acctNo"]] = "acct"
    _used = [c for k, c in col.items() if not k.startswith("_") and isinstance(c, int)]
    first_col = min(_used + [c_rate, c_rmb, c_res, c_idres])
    last_col = max(_used + [c_rate, c_rmb, c_res, c_idres])
    # 行高统一：插进来的行没有高度，和模板行高低不一，看着像分段
    row_h = ws.row_dimensions[data_start].height or 14.7

    # 上年 sheet 里已被占用的行：同一家银行开多个户时，避免两行指向同一个上年账号
    _used_prior_rows = set()
    # 有没有可比的上年账户清单：没有就整列留空，不能把"未比对"写成"本年新增"
    _has_prior = bool((res.get("stats") or {}).get("priorCount"))
    for idx, row in enumerate(rows):
        r = data_start + idx
        l1 = _norm_subject_l1(row["subject"])
        # 科目名称：优先填一级科目（BKD 表按一级科目归集，供 Lead 的 SUMIF 汇总）
        subject_disp = l1 or row["subject"] or ""
        gap_cols = set()
        vals = {}
        for fld, v in (
            ("company", row.get("company")),
            ("subject", subject_disp),
            ("bankName", row.get("bankName") or row.get("bank")),
            ("acctNo", row.get("acctNo")),
            ("currency", row.get("currency")),
            ("origEnd", row.get("orig")),
            ("localEnd", row.get("local")),
        ):
            c = col.get(fld) or (col.get("endAny") if fld in ("origEnd", "localEnd") else None)
            if not c:
                continue
            # 同列已被占用（模板只有一列「期末金额」时）就不要再覆盖
            if c in vals:
                continue
            vals[c] = v
        # 判断列 + 关键标识列：有值就填，没值留空。绝不写默认值（C 级人工判断）。
        is_cash = subject_disp == "库存现金"
        is_platform = (_norm_subject_l1(subject_disp) == "其他货币资金"
                       and is_platform_acct(row.get("bankName") or row.get("bank"),
                                            subject_disp))
        for fld, label, reason in (
            ("nature", "性质/用途", "无数据源：需人工判断（本次无上年 BKD 可回捞）"),
            ("confirm", "是否函证", "无数据源：需人工判断（本次无上年 BKD 可回捞）"),
            ("bankName", "银行/存款机构", "源表未提供银行名称（核算维度为空）"),
            ("acctNo", "账号", "源表未提供账号（核算维度为空）"),
        ):
            c = col.get(fld)
            if not c:
                continue
            if T(vals.get(c)):
                continue                       # 上面已经填到值了
            # 上年 BKD 回捞来的值（性质 / 是否函证 / Notes）：merge 阶段已经按
            # 「账号+公司」匹配好，这里必须真的写进去，否则等于白回捞。
            v0 = T(row.get(fld))
            if v0:
                # 性质/函证回捞到数字：几乎只能是上年底稿本身列位错位（把金额、折算率
                # 填进了性质列）。照写会把错值一路链进本年，所以必须留痕提醒复核。
                if fld in ("nature", "confirm") and _is_number_text(v0):
                    _mark = "_numFieldWarned_" + fld
                    if not res.get(_mark):
                        res[_mark] = True
                        res.setdefault("warnings", []).append(
                            "上年底稿「%s」列取到的是数字（%s），疑似该底稿列位错位"
                            "（比如把金额/折算率填进了这一列），本表按原值回捞——"
                            "请核对上年底稿该列并改正后重新生成" % (label, v0))
                vals[c] = v0
                continue
            # 「本期账面无此账号」的行不预填「否」：账上根本没有这个户，
            # 要不要函证、要不要替代程序，得人判断，程序不替它下结论。
            if fld == "confirm" and row.get("_priorOnly"):
                vals[c] = ""
                continue
            # 银行存款一律函证：这是准则要求，没有判断空间，不给 AI 也不给人留空。
            # 其他货币资金要看是不是银行账户（保证金户、第三方支付不函证），
            # 那个判断交给 AI 起草 + 人工确认，所以这里仍然留空。
            if fld == "confirm" and _norm_subject_l1(subject_disp) == "银行存款":
                vals[c] = "是"
                continue
            # 库存现金：银行/账号本来就没有；「性质与用途」「是否银行函证」也不适用——
            # 现金没法向第三方函证，硬留空会被当成"还没填"。四列统一写 /，不进待补清单。
            if is_cash:
                vals[c] = NO_INFO
                continue
            # 第三方支付/收款平台本来就没有银行账号（账号是平台商户号或绑定邮箱），
            # 标黄会让客户去补一个根本不存在的东西，而且会把真正缺账号的银行户淹掉。
            if fld == "acctNo" and is_platform:
                vals[c] = NO_INFO
                continue
            vals[c] = None
            gap_cols.add(c)
            gaps.append({
                "sheet": "C.00 BKD", "row": r, "col": c, "field": label,
                "company": T(row.get("company")), "subject": subject_disp,
                "acct": T(row.get("acctNo")),
                # 三种情况要分开说，否则用户会以为是"没上年数据"，
                # 实际上是"上年有这户但这格空着"或者"本年新增账户"
                "reason": ("上年 BKD 该账户这一项为空，需人工判断"
                           if row.get("_matched")
                           else ("本年新增账户（上年无对应记录），需人工判断"
                                 if (res.get("stats") or {}).get("priorCount")
                                 else reason)),
            })
        # Notes：只写项目组自己要看的内容（上年回捞的备注 + 待补提示）。
        # 两年比对的结论（【新增…】【账号取自上年…】）改放右侧独立列，
        # 混在 Notes 里既不好筛，也会把真正的业务备注淹掉。
        c_notes = col.get("notes")
        note = T(row.get("notes"))
        # 上年底稿里带着旧版的【新增…】【账号取自上年…】，回捞过来必须剥掉：
        # 这些结论现在单独一列，留在 Notes 里等于同一句话说两遍。
        if note and "【" in note:
            note = re.sub(r"\s{2,}", " ", _ID_TAG_RE.sub("", note)).strip()
        # 银行名称空、但账号有值 → TB 的「银行账户」列里只有账号，没有开户行。
        # 在 Notes 里点明要补什么，比让复核人对着一列空白猜要省事。
        _cb = col.get("bankName")
        if _cb and not T(vals.get(_cb)) and T(row.get("acctNo")) and not is_cash:
            note = (note + " " if note else "") + "【待补开户行：TB 只给了账号】"
        if c_notes:
            vals[c_notes] = note
        # 两年账户识别结果：没上传上年底稿就整列留空——没有可比数据时判"新增"是编的。
        vals[c_idres] = IDRES_TEXT.get(row.get("_how") or "", "") if _has_prior else ""

        for c, v in vals.items():
            ws.cell(r, c, v)

        # ---- 账号列 / 性质列：都从上年 BKD 取 ----
        # 账号（2026-11-07）：本年账号由「唯一主键（辅助核算项1）」匹配上来——
        # 匹配到就写公式引上年 sheet 的账号格（两年并排能追溯），匹配不到（本年新增）
        # 就留空让人补。主键那一列本来就带着完整账号，留空不会丢信息。
        # 性质与用途：同样改成链接上年（用户 2026-11-07 要求），不再复制静态值。
        _ps = res.get("priorSheet") or {}
        _c_acct = col.get("acctNo")
        _c_nat = col.get("nature")
        # 「本期无」的行本身就是上年记录，再指向上年 sheet 就成了自引用（循环引用），
        # Excel 会算成 0 并报错。这类行直接保留上年常量。
        if (_c_acct and _ps.get("sheet") and not _ps.get("error")
                and not is_cash and not row.get("_priorOnly") and row.get("_matched")):
            _pa = acct_key(row.get("priorAcct"))
            _cu2 = K(cur_code(row.get("currency")) or "CNY")
            _prow = ((_ps.get("byacct") or {}).get((_pa, _cu2))
                     or (_ps.get("byacct") or {}).get((_pa, None))) if _pa else None
            if _prow and _prow not in _used_prior_rows:
                _used_prior_rows.add(_prow)
                _cell = ws.cell(r, _c_acct)
                _cell.value = "='%s'!%s%d" % (PRIOR_SHEET,
                                              get_column_letter(PRIOR_ACCT_COL), _prow)
                _cell.number_format = "@"
                _cell.alignment = Alignment(horizontal="left", vertical="center")
                if _c_nat:
                    _nc = ws.cell(r, _c_nat)
                    _nc.value = "='%s'!%s%d" % (
                        PRIOR_SHEET, get_column_letter(PRIOR_NATURE_COL), _prow)
                    _nc.number_format = "General"
                    _nc.alignment = Alignment(horizontal="left", vertical="center",
                                              wrap_text=True)

        # 受限资金：人工列。空 = 不受限，不填任何默认值
        rv = row.get("restricted")
        if isinstance(rv, (int, float)) and rv:
            ws.cell(r, c_res, r2(rv))
            res["restrictedTotal"] = r2((res.get("restrictedTotal") or 0) + rv)

        # 折算两列：率来自「公司mapping」（按科目取余额/损益汇率），没有 mapping 时沿用
        # 人工/参数给的 _rate_of（默认 1）。折算人民币用公式连起来，改了率结果立刻跟着变。
        if _map:
            if _map.get("short"):
                ws.cell(r, 1).value = _short_formula(r)   # A 列简称，只写值不改格式
            ws.cell(r, c_rate, _rate_formula(r, _rate_kind_of(row.get("subject"))))
        else:
            ws.cell(r, c_rate, _rate_of(row))
        ws.cell(r, c_rmb,
                "=IF(%s%d=\"\",\"\",ROUND(%s%d*%s%d,2))" % (Lm, r, Lm, r, Lr, r))

        # 「记录判断理由」列：模板只在预置 4 行带 =IF(J="是","N/a","")，
        # 插列之后公式里的列引用又不会跟着右移（J 已经变成人民币列），
        # 于是这句判断去读一个金额。按行重写，引用指向真正的「是否函证」列。
        c_jn, c_cf = col.get("jNote"), col.get("confirm")
        if c_jn and is_cash:
            # 库存现金不执行银行函证，这一列对它是「不适用」：
            # 写常量 / 而不是留公式——公式在 L 不等于"是"时算出空字符串，
            # 看上去就跟"还没填"一样。
            ws.cell(r, c_jn).value = NO_INFO
        elif c_jn and c_cf:
            cur = ws.cell(r, c_jn).value
            if not T(cur) or (isinstance(cur, str) and cur.startswith("=") and "N/a" in cur):
                ws.cell(r, c_jn).value = '=IF(%s%d="是","N/a","")' % (
                    get_column_letter(c_cf), r)

        # 整行统一刷一遍：字体/边框/对齐/数字格式按列定，底色一律白。
        # 以前只刷"有值"的格子，空列（记录判断 / Notes / 受限资金）留着模板默认的
        # 等线 11 + 无边框，一行里两种字号两套边框 —— 就是"格式没刷过"的样子。
        for c in range(first_col, last_col + 1):
            cell = ws.cell(r, c)
            fmt, align = _BKD_KIND_STYLE[kinds.get(c, "text")]
            cell.font = _BKD_FONT
            cell.border = _BKD_BORDER
            cell.number_format = fmt
            cell.alignment = align
            cell.fill = _BKD_NOFILL
        ws.row_dimensions[r].height = row_h

    # D 列「银行/存款机构名称」装的是完整辅助核算原值（「开户行&账号」），
    # 模板原宽 16 放不下，右边 E 列又有值时会被截断——数据是全的，但人看不见。
    # 按实际最长值调宽（中文按 2 个字符宽、ASCII 按 1），上限 46，不动别的列。
    _cb = col.get("bankName")
    if _cb:
        _w = 0
        for _r in range(data_start, data_start + need):
            _s = T(ws.cell(_r, _cb).value)
            _w = max(_w, sum(2 if ord(ch) > 127 else 1 for ch in _s))
        if _w > 16:
            ws.column_dimensions[get_column_letter(_cb)].width = min(_w + 2, 46)

    # 「两年账户识别结果」也是长中文（"本年新增（上年无此账户）；账号取自上年"），
    # 固定 24 宽会被截断成 "本年新增（上年无此…"，结论看不见就等于没有。按内容调宽。
    _w = 0
    for _r in range(data_start, data_start + need):
        _s = T(ws.cell(_r, c_idres).value)
        _w = max(_w, sum(2 if ord(ch) > 127 else 1 for ch in _s))
    ws.column_dimensions[get_column_letter(c_idres)].width = min(max(_w + 2, 24), 46)

    # 修正合计公式的求和区间（插行后行号已下移）
    last_data_row = data_start + need - 1
    money_col = col["localEnd"]
    L = get_column_letter(money_col)
    sum_row = _find_sum_row(ws, data_start, money_col)
    if sum_row:
        ws.cell(sum_row, money_col).value = "=SUM(%s%d:%s%d)" % (L, data_start, L, last_data_row)
        # 折算人民币列也要求合计（Lead 取的是这一列）
        sum_rmb = ws.cell(sum_row, c_rmb, "=SUM(%s%d:%s%d)" % (Lrm, data_start, Lrm, last_data_row))
        # 原币列也要求合计：C.02a 的「原币 vs BKD 合计」勾稽要引用这一格。
        # 注意：这一列混着 USD/HKD/EUR…，合计**不是有意义的金额**，只作完整性控制；
        # 真正的金额口径看本币列与折算人民币列。
        if col.get("origEnd"):
            Lo = get_column_letter(col["origEnd"])
            ws.cell(sum_row, col["origEnd"],
                    "=SUM(%s%d:%s%d)" % (Lo, data_start, Lo, last_data_row))
        # 受限资金列同样要合计（Lead 的披露行引用这一格）
        Rr = get_column_letter(c_res)
        sum_res = ws.cell(sum_row, c_res, "=SUM(%s%d:%s%d)" % (Rr, data_start, Rr, last_data_row))
        res["restrictedSumCell"] = "%s%d" % (Rr, sum_row)
        # 合计行整行统一：字体 Arial 10、右对齐、金额格式 MONEY_FMT。
        # 模板合计行用的是会计格式（_(* #,##0_)），和明细行的 #,##0.00 不是一套（2026-11-12 起统一成 ACCT_FMT），
        # 同一列上下两种格式，复制粘贴出去就变成文本；新加的三列更不能一边加粗一边不加粗。
        # 边框保持模板原样（合计行本来就没有竖线），只统一样式不重新画框。
        for c in (col.get("origEnd"), col.get("localEnd"), c_rmb, c_res):
            if not c:
                continue
            cell = ws.cell(sum_row, c)
            cell.font = _BKD_FONT
            cell.number_format = MONEY_FMT
            cell.alignment = _BKD_ALIGN_R
        if c_rate:
            cell = ws.cell(sum_row, c_rate)
            cell.font = _BKD_FONT
            cell.number_format = "0.0000"
            cell.alignment = _BKD_ALIGN_R
        # 勾稽行（合计行 + 1 是 Rx，+2 是勾稽）：
        # **按 Lead 的实际布局重算引用**，不再沿用模板里写的 'C.00 Lead'!F41 ——
        # 模板的 F41 是 V7 的行号，客户模板插过行（瑞云合计在 46 行）就会指到错的行，
        # 勾稽看着不平、实际是公式指错了。只有定位失败时才退回模板原值。
        # 有外币主体时本币列合计 ≠ Lead（Lead 是人民币口径），勾稽必须用折算人民币列。
        chk_cell = ws.cell(sum_row + 2, money_col)
        old = chk_cell.value
        ref = ""
        _lead = wb["C.00 Lead"] if "C.00 Lead" in wb.sheetnames else None
        if _lead is not None:
            _l1 = (index_lead_layout(_lead) or {}).get("lead1") or {}
            _tot = _l1.get("total") or (locate_lead_rows(_lead) or {}).get("total")
            _bcol = (_l1.get("cols") or {}).get("book")
            if _tot and _bcol:
                ref = "'C.00 Lead'!%s%d" % (get_column_letter(_bcol), _tot)
        if not ref and isinstance(old, str) and "C.00 Lead" in old:
            m = re.search(r"'C\.00 Lead'!\$?([A-Z]{1,3})\$?(\d+)", old)
            if m:
                ref = "'C.00 Lead'!%s%s" % (m.group(1), m.group(2))
        if ref:
            chk_cell.value = "=%s%d-%s" % (Lrm, sum_row, ref)
            if any(abs(_rate_of(x) - 1.0) > 1e-12 for x in rows):
                chk_cell.comment = Comment(
                    "本币列合计已按主体折算率折成人民币后与 Lead 勾稽（%s 列 %s%d）。"
                    % (FX_RMB_KEY, Lrm, sum_row),
                    "PTW")
        # 勾稽行新增列补边框，保持"整块表格"的观感
        for c in (c_rate, c_rmb, c_res, c_idres):
            cell = ws.cell(sum_row + 2, c)
            cell.border = _BKD_BORDER
            cell.font = _BKD_FONT

    # ---- 明细区一律不上色（表头以下整张表白底）----
    # 之前用浅黄标「本年新增」、亮黄标「缺资料」。插了折算两列、又按行重刷样式之后，
    # 两种黄经常落在同一行里（新增行 + 该行有缺资料格），反复看就是"花"。
    # 现在统一白底；「哪里要补」由「检查」表的 gaps 清单逐格列出
    # （页签、行、列、字段、原因、账号），信息一点没少，只是不往表格上涂色。
    # 表头（第 8~9 行）的模板配色保持原样，不在此循环内。
    for r in range(data_start, ws.max_row + 1):
        for c in range(first_col, last_col + 1):
            try:
                cell = ws.cell(r, c)
                if cell.fill is not None and cell.fill.fill_type:
                    cell.fill = _BKD_NOFILL
            except Exception:
                pass

    # ---- 恢复模板的「空白格高亮」（条件格式），铺到整个明细区 ----
    # 只铺到 Notes 列为止：受限资金列空 = 不受限，不该被当成"没填"而整列变黄。
    res["blankHighlight"] = _restore_blank_highlight(
        ws, data_start, last_data_row, col.get("notes") or last_col)

    # 表格最外侧那圈深蓝：必须在"一律白底"的清理之后画，否则会被刷白
    res["frame"] = _paint_bkd_frame(ws)


# 三个一级科目 → 总账科目编码（标准会计科目编码）
LEAD_CODE_BY_SUBJECT = {"库存现金": "1001", "银行存款": "1002", "其他货币资金": "1012"}


#: Lead 里要填的三个一级科目（表1 与表2 都出现一次）
LEAD_SUBJECTS = ("库存现金", "银行存款", "其他货币资金")
#: 表1 合计行、表2 合计行、境外资金行、日期行在 V7 模板里的位置（**只在定位失败时兜底**）
LEAD_FALLBACK = {"rows": {"库存现金": 37, "银行存款": 38, "其他货币资金": 39},
                 "total": 41, "b2": {"库存现金": 62, "银行存款": 63, "其他货币资金": 64},
                 "b2Total": 66, "overseas": 68, "date": 35}


def _lead_scan(lead, subject, limit=120):
    """找出写着这个科目的所有行（B 列或 D 列，忽略空格）。"""
    out = []
    for r in range(1, min(lead.max_row, limit) + 1):
        b = T(lead.cell(r, 2).value).replace(" ", "").replace("\u3000", "")
        d = T(lead.cell(r, 4).value).replace(" ", "").replace("\u3000", "")
        if b == subject or d == subject:
            out.append(r)
    return out


def _lead_find_label(lead, keys, start, limit=120, cols=(2, 3, 4, 5, 6)):
    """从 start 往下找第一个含 keys 的行。

    要扫到 E 列：模板的「合计」两个字写在 E41（B/D 都是空的），
    只扫 B/D 会找不到合计行。扫到 F 就够——再往右是金额和公式，误命中风险大。
    """
    for r in range(max(1, start), min(lead.max_row, limit) + 1):
        txt = "".join(T(lead.cell(r, c).value).replace(" ", "") for c in cols)
        if any(k in txt for k in keys):
            return r
    return None


def locate_lead_rows(lead):
    """按科目名称定位 C.00 Lead 的行号 —— **不写死 37/38/39**。

    为什么必须定位：V7 模板是表1 37/38/39、表2 62/63/64、境外资金 68、日期 35。
    但客户模板多插一行空行、加一行「存放财务公司款项」、或者换一版 SWP，
    这些行号全变。写死之后程序**不会报错**，只会静默错位：

      · 表1 的 SUMIF 打到空行上，真正的科目行反而空着
      · 表2 的 =J37 引用到表头行
      · 境外资金的 SUMPRODUCT 落进「合计」行，把合计公式覆盖掉

    数看着都有，实际全错，而且 ok 仍然返回 True。

    返回 {"rows": {科目: 行}, "total": 行, "b2": {科目: 行}, "b2Total": 行,
          "overseas": 行, "date": 行, "source": "text"/"default", "warnings": [...]}
    """
    out = {"rows": {}, "total": None, "b2": {}, "b2Total": None,
           "overseas": None, "date": None, "source": "text", "warnings": []}

    for subj in LEAD_SUBJECTS:
        hits = _lead_scan(lead, subj)
        if hits:
            out["rows"][subj] = hits[0]
            if len(hits) > 1:
                out["b2"][subj] = hits[1]

    # 定位失败（科目名被改过 / 模板完全不同）→ 用 V7 的写死值兜底，但要留告警。
    # 兜底而不报错，是因为绝大多数底稿就是 V7；但必须让人看得见。
    if len(out["rows"]) < 3 or len(out["b2"]) < 3:
        # 先记下实际找到几个，再覆盖成兜底值（否则告警里永远是 3/3）
        n1, n2 = len(out["rows"]), len(out["b2"])
        out["source"] = "default"
        out.update({"rows": dict(LEAD_FALLBACK["rows"]),
                    "total": LEAD_FALLBACK["total"],
                    "b2": dict(LEAD_FALLBACK["b2"]),
                    "b2All": dict(LEAD_FALLBACK["b2"]),
                    "b2Total": LEAD_FALLBACK["b2Total"],
                    "overseas": LEAD_FALLBACK["overseas"],
                    "date": LEAD_FALLBACK["date"]})
        miss = []
        if n1 < 3:
            miss.append("表1 %d/3 个科目" % n1)
        if n2 < 3:
            miss.append("表2 %d/3 个科目" % n2)
        out["warnings"].append(
            "LEAD_ROWS_DEFAULT：Lead 里没找全「库存现金/银行存款/其他货币资金」"
            "（%s），已按 V7 模板的 37/38/39、62/63/64、68 写入。"
            "若这份底稿不是 V7 结构，Lead 的数字会对错行，请核对。"
            % "，".join(miss))
        return out

    # 表1 合计行：表1 最后一个科目行往下第一个「合计」
    r1_max = max(out["rows"].values())
    out["total"] = _lead_find_label(lead, ("合计",), r1_max + 1, r1_max + 8)
    # 表2 合计行
    r2_max = max(out["b2"].values())
    out["b2Total"] = _lead_find_label(lead, ("合计",), r2_max + 1, r2_max + 8)
    # 表2 的**全部**科目行：模板里除了三个固定科目，还有「存放财务公司款项」这类行。
    # 只认固定三个的话，这些行的期末数没有链接、合计也漏算它们（本次实测 D68 少了它）。
    b2_all = {}
    r2_lo = min(out["b2"].values())
    r2_hi = (out["b2Total"] or r2_max + 1) - 1
    for r in range(r2_lo, r2_hi + 1):
        subj = T(lead.cell(r, 2).value)
        subj = subj.replace(" ", "")
        if subj and subj != "合计" and "其中" not in subj:
            b2_all[subj] = r
    out["b2All"] = b2_all or dict(out["b2"])
    # 境外资金行
    out["overseas"] = _lead_find_label(lead, ("存放在境外", "境外"), r2_max + 1)
    # 日期行：模板里 F 列写的是 =C3（资产负债表日）
    for r in range(1, min(lead.max_row, 60) + 1):
        if T(lead.cell(r, 6).value).upper() in ("=C3", "=$C$3"):
            out["date"] = r
            break

    # 表1 行必须严格在表2 行之上，否则说明定位把它们搞混了
    if min(out["b2"].values()) <= r1_max:
        out["warnings"].append(
            "LEAD_ROWS_ORDER：表1 与表2 的科目行顺序异常（表1 到 %d 行，表2 从 %d 行起），"
            "请核对 Lead 结构。" % (r1_max, min(out["b2"].values())))
    return out


def _restore_blank_highlight(ws, data_start, last_data_row, last_col):
    """把模板自带的「空白格高亮」条件格式铺到整个明细区。

    模板里有一条 containsBlanks 规则（dxf 底色 FFFFCC 浅黄），原来只覆盖
    `B10:K15` 和一个 Notes 框。插了两列、写了上百行之后：
      - 新插的 I/J 两列不在 K 左边，盖不到；
      - 第 16 行往下的数据全在范围外，盖不到。
    结果就是"没填的格子"看起来一片白，模板原本的机制等于失效了。

    这里做三件事，一件都不自己发明颜色：
      1. 把原规则的 dxf 原样 deepcopy 出来（颜色、图案都照模板）；
      2. 公式锚点改成新范围的左上角（模板是 `LEN(TRIM(B10))=0`）；
      3. 主范围铺到 `B{start}:{last}{end}`，Notes 合并区另加一格。

    注意 CF 与硬写底纹的本质区别：**空白才显示，填上值就自动消失**。
    之前那版用 cell.fill 直接涂，所以填了值还是黄，看着才假。
    """
    from openpyxl.formatting.formatting import ConditionalFormattingList
    rules = []
    for cf in ws.conditional_formatting:
        for r in cf.rules:
            rules.append(_copy.deepcopy(r))
    if not rules:
        return 0
    # 只把行号换成新起点，列字母保留（数据区都从 B 列起）
    for r in rules:
        if getattr(r, "formula", None):
            r.formula = [re.sub(r"(\$?[A-Z]{1,3})\$?\d+",
                                lambda m: "%s%d" % (m.group(1), data_start), f, count=1)
                         for f in r.formula]
    ws.conditional_formatting = ConditionalFormattingList()
    rng = "B%d:%s%d" % (data_start, get_column_letter(last_col), last_data_row)
    for r in rules:
        ws.conditional_formatting.add(rng, r)
    # 明细区下方的 Notes 输入框（合并区）也保留一格空白高亮
    for mc in sorted(ws.merged_cells.ranges, key=lambda x: x.min_row):
        if mc.min_row > last_data_row and (mc.max_col - mc.min_col) >= 2:
            for r in rules:
                ws.conditional_formatting.add(
                    "%s%d" % (get_column_letter(mc.min_col), mc.min_row),
                    _copy.deepcopy(r))
            break
    return len(rules)




PRIOR_LEAD_SHEET = "上年 C.00 Lead"


def _lead_norm(v):
    """表头/科目名归一：去空格、全角空格、换行。"""
    return str(v or "").replace(" ", "").replace("\u3000", "").replace("\n", "")


#: 表1 的列：语义键 -> (包含词, 排除词)。按文字认列，不写死列号。
LEAD1_HEADS = (
    ("entity", ("账套名称", "账套编码"), ()),
    ("code", ("总账科目编码", "科目编码"), ()),
    ("subject", ("科目名称",), ()),
    ("index", ("索引号",), ()),
    ("book", ("期末账面数",), ()),
    ("adjBook", ("账表调整数",), ()),
    ("unaudited", ("期末未审数",), ()),
    ("auditAdj", ("审计调整数",), ()),
    ("audited", ("期末审定数",), ("上期", "上年")),
    # 上期列的写法各所不一：瑞云是「上期末审定数」，V7 模板表2 写「上期未审数」。
    # 只认一种写法时，另一种模板会认不到这一列 → 上年数静默不链接（2026-11-10 踩过）。
    ("prior", ("上期末审定数", "上期未审数", "上年数", "上期末数", "上期数"), ()),
    ("delta", ("变动金额",), ()),
    ("deltaPct", ("变动%", "变动率"), ()),
    ("notes", ("notes", "备注"), ()),
)
#: 表2 的列
LEAD2_HEADS = (
    ("item", ("报表科目",), ()),
    ("current", ("期末审定数",), ("上期", "上年")),
    # 同上：V7 模板表2 的上期列叫「上期未审数」，写死「上期末审定数」会整块认不到。
    ("prior", ("上期末审定数", "上期未审数", "上年数", "上期末数", "上期数"), ()),
)


def _lead_head_row(ws, must, skip=(), any_of=()):
    """找同时含 must 各关键词的表头行。找不到返回 None。

    不写死行号：上年 Lead 的封面、认定区、波动说明占几行是因所而异，
    这张真实样本的表头在 37 行，模板是 36 行。

    any_of：给一组同义写法，命中任意一个即可。表2 的上期列在不同底稿里
    写作「上期末审定数」或「上期未审数」，必须两种都认，否则整块表2 认不到。
    """
    lim = min(ws.max_row, 400)
    for r in range(1, lim + 1):
        texts = [_lead_norm(ws.cell(r, c).value)
                 for c in range(1, min(ws.max_column, 40) + 1)]
        joined = "|".join(t for t in texts if t).lower()
        if any(_lead_norm(x).lower() in joined for x in skip):
            continue
        if all(any(_lead_norm(k).lower() in t.lower() for t in texts if t) for k in must) \
                and (not any_of or any(any(_lead_norm(k).lower() in t.lower()
                                            for t in texts if t) for k in any_of)):
            return r
    return None


def _lead_head_cols(ws, hdr, spec):
    """按表头文字映射列号；返回 {语义: 列号}。"""
    out, used = {}, set()
    for key, inc, exc in spec:
        for c in range(1, min(ws.max_column, 40) + 1):
            if c in used:
                continue
            t = _lead_norm(ws.cell(hdr, c).value).lower()
            if not t:
                continue
            if any(_lead_norm(k).lower() in t for k in inc) and \
                    not any(_lead_norm(x).lower() in t for x in exc):
                out[key] = c
                used.add(c)
                break
    return out


def _lead_block(ws, hdr, cols, subj_key="subject"):
    """表头下方的数据区：从 hdr+1 到「合计」前，返回 (首行, 末行, 合计行, [(科目, 行)])。

    结束条件按内容判断，不看固定行数：
      · 科目列出现「合计」→ 到此为止，这一行就是合计行；
      · 连续两行科目列为空 → 视为表结束（防止一路吃到表2）。
    """
    sc = cols.get(subj_key)
    if not sc:
        return None
    rows, blank = [], 0
    r = hdr + 1
    total = None
    lim = min(ws.max_row, hdr + 300)
    while r <= lim:
        # 「合计」不一定写在科目列（真实样本把它放在 E 列），所以整行扫一遍
        line = "|".join(_lead_norm(ws.cell(r, c).value)
                        for c in range(2, min(ws.max_column, 40) + 1))
        if "合计" in line:
            total = r
            break
        subj = _lead_norm(ws.cell(r, sc).value)
        if not subj:
            blank += 1
            if blank >= 2 and rows:
                break
            r += 1
            continue
        blank = 0
        rows.append((subj, r))
        r += 1
    if not rows:
        return None
    return rows[0][1], rows[-1][1], total, rows


def index_lead_layout(ws):
    """识别一张 Lead 里表1、表2 的位置，全部按表头文字与内容边界现算。

    返回 {"lead1": {...}, "lead2": {...} | None}，每块含：
      hdr 表头行 / cols 列映射 / first,last 数据行 / total 合计行 / rows 科目->行
    """
    out = {"lead1": None, "lead2": None}

    h1 = _lead_head_row(ws, ("科目名称", "上期末审定数"), skip=("报表科目",))
    if h1:
        c1 = _lead_head_cols(ws, h1, LEAD1_HEADS)
        b1 = _lead_block(ws, h1, c1)
        if b1:
            out["lead1"] = {"hdr": h1, "cols": c1, "first": b1[0], "last": b1[1],
                            "total": b1[2], "rows": dict(b1[3])}

    h2 = _lead_head_row(ws, ("报表科目",),
                        any_of=("上期末审定数", "上期未审数", "上年数",
                                "上期末数", "上期数"))
    if not h2:
        # 表2 的上期列表头在 V7 模板里是公式（D61 = =K36，指向表1 的「上期末审定数」）。
        # openpyxl 以 data_only 读一份**没被 Excel 打开过**的文件时取不到缓存值，
        # 表头就是空白 → 按文字永远认不到表2，表2 与披露区的上年数会静默不链接
        # （2026-11-10 实测：本年 Lead 表2 D 列只剩 =K37，披露区整格空）。
        # 这时退到结构定位：B 列写「报表科目」，右侧两列依次是本期/上期。
        h2 = _lead_head_row(ws, ("报表科目",))
    if h2:
        c2 = _lead_head_cols(ws, h2, LEAD2_HEADS)
        # 表头读不到文字时按结构补齐（科目列右一列=本期、右二列=上期）
        if "item" in c2:
            c2.setdefault("current", c2["item"] + 1)
            c2.setdefault("prior", c2["item"] + 2)
        b2 = _lead_block(ws, h2, c2, subj_key="item")
        if b2:
            out["lead2"] = {"hdr": h2, "cols": c2, "first": b2[0], "last": b2[1],
                            "total": b2[2], "rows": dict(b2[3])}
    return out


#: Lead 里属于「金额」的语义列。变动率、备注、编码、名称一律不刷。
LEAD_MONEY_KEYS = ("book", "adjBook", "unaudited", "auditAdj", "audited", "prior", "delta")


def _paint_money_fmt(ws, first, last, cols):
    """只设数字格式，不碰字体/边框/底色/对齐。

    用户 2026-11-12 口径：模板自带的金额格可以改成会计格式（零显示 -），
    但**不许顺手改字体粗细、颜色、边框**——模板观感是项目组的，程序只统一"数字怎么显示"。
    """
    last = min(last, ws.max_row)
    for r in range(max(first, 1), last + 1):
        for c in cols:
            if c:
                ws.cell(r, c).number_format = ACCT_FMT


def _paint_lead_money_fmt(ws):
    """把 Lead 表1 / 合计 / 勾稽 / 披露区 / 表2 的金额格统一成会计格式。

    范围全部由 index_lead_layout 现算，不写死行列：
      · 表1 块：数据首行 → 表2 表头前一行（中间夹合计、勾稽、境外/受限披露行，
        它们与表1 用的是同一组金额列）；
      · 表2 块：数据首行 → 末行，列取「本期/上期」再往右两列（模板的 Check 列）。
    """
    lay = index_lead_layout(ws)
    l1, l2 = lay.get("lead1"), lay.get("lead2")

    def _l2_cols():
        base = (l2.get("cols") or {}).get("current")
        if not base:
            return []
        # C/D 本期/上期 + 右侧两个 Check 列；超出模板列数就算了
        return [c for c in range(base, base + 4) if c <= ws.max_column]

    if l1 and l2 and l2["hdr"] < l1["hdr"]:          # 表2 在上（少见）
        _paint_money_fmt(ws, l2["first"], l1["first"] - 1, _l2_cols())
        _paint_money_fmt(ws, l1["first"], ws.max_row,
                         [l1["cols"].get(k) for k in LEAD_MONEY_KEYS])
        return
    if l1:
        stop = (l2["hdr"] - 1) if l2 else ws.max_row
        _paint_money_fmt(ws, l1["first"], stop,
                         [l1["cols"].get(k) for k in LEAD_MONEY_KEYS])
    if l2:
        _paint_money_fmt(ws, l2["first"], ws.max_row, _l2_cols())


def _copy_cell_fmt(sc, dc):
    """跨工作簿复制单元格格式（逐属性复制）。

    不能写 `dc._style = sc._style`：StyleArray 里的索引只在本簿的样式表内有效，
    跨簿照搬会串成别的字体/边框。这里逐项拷贝字体、填充、边框、对齐、数字格式。
    """
    from copy import copy as _cp
    try:
        if sc.has_style:
            dc.font = _cp(sc.font)
            dc.fill = _cp(sc.fill)
            dc.border = _cp(sc.border)
            dc.alignment = _cp(sc.alignment)
            dc.protection = _cp(sc.protection)
            dc.number_format = sc.number_format
    except Exception:
        pass


def copy_prior_lead_sheet(wb, data, title=None):
    """把上年 Lead **整表**复制成本年底稿里的一张 sheet：保留格式，只留数值。

    与旧做法（只拼两列科目+金额）的区别：封面、认定区、波动说明、表1、调整汇总表、
    表2 全部原样带过来，复核人能在本年文件里直接翻到上年底稿的样子；
    但所有公式都换成计算后的数值，不把上年文件的公式链带进来。

    返回给 _fill_lead 用的索引：t1Subj/t1Amt/t2Subj/t2Amt 是**区域内**的绝对引用
    （不带表名，调用方自己拼），t1 是科目清单（用于本年插行）。
    """
    import io as _io
    from copy import copy as _cp

    title = title or PRIOR_LEAD_SHEET
    if not data:
        return {"error": "没有上年底稿内容"}
    try:
        sv = load_workbook(_io.BytesIO(data), data_only=True)
        sf = load_workbook(_io.BytesIO(data), data_only=False)
    except Exception as e:
        return {"error": "读取上年 Lead 失败：%s" % e}

    # 选表：先找真的有「科目名称 + 上期末审定数」表头的那张，再退到名字带 Lead 的
    pick = None
    for nm in sf.sheetnames:
        if _lead_head_row(sf[nm], ("科目名称", "上期末审定数"), skip=("报表科目",)):
            pick = nm
            break
    if not pick:
        for nm in sf.sheetnames:
            if "lead" in nm.lower():
                pick = nm
                break
    if not pick:
        return {"error": "上年文件里没有像 Lead 的工作表（%s）" % "、".join(sf.sheetnames[:6])}

    src_f, src_v = sf[pick], sv[pick]
    if title in wb.sheetnames:
        del wb[title]
    dst = wb.create_sheet(title)

    def _resolve_simple_ref(fx_cell, sheet_v, sheet_f):
        """把「=J36」这种单格引用就地解成值。

        为什么要它：模板 Lead 表2 的表头是公式（C61==J36、D61==K36），
        一份没被 Excel 打开过的文件没有缓存值，直接搬过来表头就是空白，
        上年 Lead 那张表看着像少了一行标题。
        """
        v = fx_cell.value
        if not isinstance(v, str) or not v.startswith("="):
            return None
        m = re.match(r"^=\s*\$?([A-Za-z]{1,3})\$?(\d+)\s*$", v)
        if not m:
            return None
        col = 0
        for ch in m.group(1).upper():
            col = col * 26 + (ord(ch) - 64)
        rr = int(m.group(2))
        got = sheet_v.cell(rr, col).value
        if got is None:
            got = sheet_f.cell(rr, col).value
            if isinstance(got, str) and got.startswith("="):
                return None
        return got

    max_r = min(src_f.max_row, 400)
    max_c = min(src_f.max_column, 40)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            sc = src_f.cell(r, c)
            value = src_v.cell(r, c).value
            # 公式没有缓存值（文件从未被 Excel 打开过）→ 先试单格引用，解不出再留空，
            # 绝不把公式搬进来。
            if value is None and isinstance(sc.value, str) and sc.value.startswith("="):
                value = _resolve_simple_ref(sc, src_v, src_f)
                if value is None:
                    value = ""
            dc = dst.cell(r, c, value)
            _copy_cell_fmt(sc, dc)

    for m in src_f.merged_cells.ranges:
        try:
            dst.merge_cells(str(m))
        except Exception:
            pass
    # 列宽/行高：源文件是**按区间**存的（如 C~J 共用一个 16.77、T 一直到 16384）。
    # 只按字典键复制，一个区间只会复制到第一列，其余列退回默认宽 → 数字挤成 ####。
    # 所以按 min..max 展开成单列再设。
    for dim in src_f.column_dimensions.values():
        try:
            if not dim.width or not dim.min:
                continue
            lo = int(dim.min)
            hi = int(dim.max or lo)
            if hi > 60:                     # 表格只用得到前几十列，别把 16384 列都铺开
                hi = 60
            for c in range(lo, hi + 1):
                d = dst.column_dimensions[get_column_letter(c)]
                # 只设 width：openpyxl 保存时会自动带 customWidth="1"。
                # 显式赋 customWidth 会抛 AttributeError（该属性是只读的），
                # 一旦抛异常，整个区间的循环就断在第一列上——D~J 会退回默认宽、数字挤成 ####。
                d.width = dim.width
                if dim.hidden:
                    d.hidden = True
        except Exception:
            pass
    for dim in src_f.row_dimensions.values():
        try:
            if not dim.height or not dim.min:
                continue
            lo = int(dim.min)
            hi = int(dim.max or lo)
            for r in range(lo, hi + 1):
                dst.row_dimensions[r].height = dim.height
                if dim.hidden:
                    dst.row_dimensions[r].hidden = True
        except Exception:
            pass
    try:                                        # 默认行高/列宽也照抄，否则空行会矮一截
        dst.sheet_format.defaultRowHeight = src_f.sheet_format.defaultRowHeight
        if src_f.sheet_format.defaultColWidth:
            dst.sheet_format.defaultColWidth = src_f.sheet_format.defaultColWidth
    except Exception:
        pass
    try:
        dst.freeze_panes = src_f.freeze_panes
    except Exception:
        pass
    try:
        dst.sheet_view.showGridLines = src_f.sheet_view.showGridLines
        dst.sheet_view.zoomScale = src_f.sheet_view.zoomScale
    except Exception:
        pass
    try:
        dst.page_setup.orientation = src_f.page_setup.orientation
        dst.page_setup.paperSize = src_f.page_setup.paperSize
        dst.page_setup.fitToWidth = src_f.page_setup.fitToWidth
        dst.page_setup.fitToHeight = src_f.page_setup.fitToHeight
        if src_f.sheet_properties.pageSetUpPr is not None:
            dst.sheet_properties.pageSetUpPr = _cp(src_f.sheet_properties.pageSetUpPr)
    except Exception:
        pass

    # 条件格式：规则里的 dxf 是"工作簿级"的，直接搬对象不会进本簿样式表，
    # 所以逐项重建一份 DifferentialStyle
    try:
        from openpyxl.formatting.rule import Rule
        from openpyxl.styles.differential import DifferentialStyle
        for rng in src_f.conditional_formatting:
            for rule in rng.rules:
                dxf = None
                if rule.dxf is not None:
                    d = rule.dxf
                    dxf = DifferentialStyle(
                        font=_cp(d.font) if d.font else None,
                        fill=_cp(d.fill) if d.fill else None,
                        border=_cp(d.border) if d.border else None,
                        numFmt=_cp(d.numFmt) if d.numFmt else None,
                        alignment=_cp(d.alignment) if d.alignment else None,
                        protection=_cp(d.protection) if d.protection else None)
                nr = Rule(type=rule.type, formula=list(rule.formula or []),
                          dxf=dxf, stopIfTrue=rule.stopIfTrue)
                if rule.type == "containsBlanks" and not nr.formula:
                    nr.formula = ["LEN(TRIM(%s))=0" % str(rng.sqref).split()[0]]
                dst.conditional_formatting.add(str(rng.sqref), nr)
    except Exception:
        pass

    # 数据验证
    try:
        from openpyxl.worksheet.datavalidation import DataValidation
        for dv in src_f.data_validations.dataValidation:
            ndv = DataValidation(type=dv.type, formula1=dv.formula1, formula2=dv.formula2,
                                 allow_blank=dv.allow_blank,
                                 showErrorMessage=dv.showErrorMessage,
                                 showInputMessage=dv.showInputMessage,
                                 error=dv.error, errorTitle=dv.errorTitle,
                                 prompt=dv.prompt, promptTitle=dv.promptTitle)
            ndv.sqref = str(dv.sqref)
            dst.add_data_validation(ndv)
    except Exception:
        pass

    lay = index_lead_layout(dst)
    L1, L2 = lay.get("lead1"), lay.get("lead2")

    def rng(block, key):
        if not block or key not in block["cols"]:
            return None
        col = get_column_letter(block["cols"][key])
        return "$%s$%d:$%s$%d" % (col, block["first"], col, block["last"])

    out = {"sheet": title, "source": pick, "layout": lay,
           "t1Rows": len(L1["rows"]) if L1 else 0,
           "t2Rows": len(L2["rows"]) if L2 else 0,
           "warnings": []}
    if L1:
        out.update({"t1Subj": rng(L1, "subject"), "t1Amt": rng(L1, "audited"),
                    "t1Code": rng(L1, "code"),
                    "t1": [{"subject": k} for k in L1["rows"].keys()]})
    else:
        out["warnings"].append("上年 Lead 里没认到表1（表头缺「科目名称/上期末审定数」）")
    if L2:
        out.update({"t2Subj": rng(L2, "item"), "t2Amt": rng(L2, "current"),
                    "t2": [{"item": k, "current": None} for k in L2["rows"].keys()]})
        # 合计行**下面**还有两行：「其中：存放在境外的款项总额」「所有权或使用权受到限制的
        # 货币资金」。_lead_block 见到「合计」就停，这两行不在它的范围里，可本年表2 恰恰
        # 要用它们的上期数。所以再算一条"从首行一直到最后一条科目文字"的宽区间。
        # 区间里夹着「合计」「Rx」等文字行无所谓——SUMIF 按科目全名精确匹配，不会被它们带偏。
        _sc = L2["cols"].get("item")
        _last = L2["last"]
        if _sc:
            for _r in range((L2.get("total") or L2["last"]) + 1, min(dst.max_row, 300) + 1):
                if _lead_norm(dst.cell(_r, _sc).value):
                    _last = _r
        if _sc and _last > L2["last"]:
            _cl = get_column_letter(_sc)
            out["t2SubjExt"] = "$%s$%d:$%s$%d" % (_cl, L2["first"], _cl, _last)
            if L2["cols"].get("current"):
                _al = get_column_letter(L2["cols"]["current"])
                out["t2AmtExt"] = "$%s$%d:$%s$%d" % (_al, L2["first"], _al, _last)
    elif L1:
        # 认不到表2 时以前是「静默」：表2 上期数只剩内部 =K37，披露区上年数整格留空，
        # 表面上数字都对、实际与上年 Lead 的链断了。这里必须留痕，让人知道去查表头写法。
        out["warnings"].append(
            "上年 Lead 里没认到表2（表头需含「报表科目」+ 上期列，如「上期末审定数」"
            "或「上期未审数」）：表2 与披露区的上年数不会链接到上年 Lead")
    return out


def _lead_insert_disclosure(lead, lr, items):
    """在表2 合计行之前插入上年有、本年没有的报表科目。

    表1 插了、表2 不插，上年数就没处落，表2 合计也和上年对不上。
    插完要重新定位（合计行、Rx 行、境外行、受限行都会下移），
    条件格式与数据验证同样靠 _shift_marks_after_insert 平移。
    """
    if not items:
        return 0
    b2 = lr.get("b2All") or lr.get("b2") or {}
    b2t = lr.get("b2Total")
    if not b2 or not b2t:
        return 0
    last = max(b2.values())
    n = len(items)
    moved = []
    for mc in list(lead.merged_cells.ranges):
        if mc.min_row > last:
            moved.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
            lead.unmerge_cells(str(mc))
    heights = {r: lead.row_dimensions[r].height
               for r in list(lead.row_dimensions.keys()) if r > last}
    lead.insert_rows(last + 1, n)
    _shift_marks_after_insert(lead, last, n)
    for r in sorted(heights, reverse=True):
        try:
            lead.row_dimensions[r + n].height = heights[r]
        except Exception:
            pass
    for (r1, r2, c1, c2) in moved:
        try:
            lead.merge_cells(start_row=r1 + n, start_column=c1,
                             end_row=r2 + n, end_column=c2)
        except Exception:
            pass
    for i, item in enumerate(items):
        r = last + 1 + i
        for c in range(2, 8):
            try:
                lead.cell(r, c)._style = _copy_style(lead.cell(last, c)._style)
            except Exception:
                pass
        lead.cell(r, 2, item)
        lead.cell(r, 3, None)
        lead.cell(r, 4, None)
        try:
            lead.row_dimensions[r].height = lead.row_dimensions[last].height
        except Exception:
            pass
    return n


def _shift_marks_after_insert(ws, anchor, n):
    """插行后平移条件格式与数据验证的区域。

    openpyxl 的 `insert_rows` **只搬单元格**，条件格式（sqref）与数据验证
    （dataValidation 的 sqref）一个都不动。后果全是"看得见、但不报错"的错位：

      · 淡黄/浅色空白高亮还停在原来那三行上 → 新增的两行没有高亮，合计行反而有；
      · 合计行的红框（O41/P41）指到了第一个科目行；
      · 「调整类型」下拉还是 B54:B57 → 表头没有下拉、最后两行没有下拉。

    规则：区域整体在插入点之下 → 整体 +n；区域的末行正好等于插入点（就是被撑开的
    科目块，如 37:39）→ 末行 +n，把新行纳进来。
    """
    from openpyxl.worksheet.cell_range import CellRange

    def move(r):
        if r >= anchor:
            return r + n
        return r

    # 1) 条件格式
    try:
        from openpyxl.formatting.formatting import ConditionalFormattingList
        old = []
        for cf in ws.conditional_formatting:
            old.append(([cr for cr in cf.sqref.ranges], list(cf.rules)))
        new = ConditionalFormattingList()
        for ranges, rules in old:
            moved = []
            for cr in ranges:
                # 只有"跨在插入点上"的块才延伸；整块在下方的一律 +n
                end = cr.max_row + n if cr.max_row >= anchor else cr.max_row
                start = cr.min_row + n if cr.min_row > anchor else cr.min_row
                if start > end:
                    start = end
                moved.append(CellRange(min_col=cr.min_col, min_row=start,
                                       max_col=cr.max_col, max_row=end).coord)
            for rule in rules:
                new.add(" ".join(moved), rule)
        ws.conditional_formatting = new
    except Exception:
        pass

    # 2) 数据验证
    try:
        for dv in ws.data_validations.dataValidation:
            parts = []
            for cr in dv.sqref.ranges:
                end = cr.max_row + n if cr.max_row >= anchor else cr.max_row
                start = cr.min_row + n if cr.min_row > anchor else cr.min_row
                if start > end:
                    start = end
                parts.append(CellRange(min_col=cr.min_col, min_row=start,
                                       max_col=cr.max_col, max_row=end).coord)
            dv.sqref = " ".join(parts)
    except Exception:
        pass


def _lead_insert_subjects(lead, lr, extra):
    """在表1 末尾（合计行之前）插入上年有、本年没有的科目行。

    货币资金本年常常没有「定期存款」这类科目，而上年有；不插行的话，
    上年数没处落，变动分析也就断了。

    openpyxl 的 insert_rows **不搬合并区**，下面的「波动说明」和表2 会一起错位，
    所以先记下所有受影响的合并区、拆掉、插完再按 +n 重挂。
    """
    if not extra:
        return 0
    rows = lr.get("rows") or {}
    if not rows:
        return 0
    last = max(rows.values())
    n = len(extra)
    moved = []
    for mc in list(lead.merged_cells.ranges):
        if mc.min_row > last:
            moved.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
            lead.unmerge_cells(str(mc))
    heights = {r: lead.row_dimensions[r].height for r in list(lead.row_dimensions.keys()) if r > last}
    lead.insert_rows(last + 1, n)
    # 条件格式 / 数据验证的区域不会跟着搬，必须手工平移
    _shift_marks_after_insert(lead, last, n)
    for r in sorted(heights, reverse=True):
        try:
            lead.row_dimensions[r + n].height = heights[r]
        except Exception:
            pass
    for (r1, r2, c1, c2) in moved:
        try:
            lead.merge_cells(start_row=r1 + n, start_column=c1, end_row=r2 + n, end_column=c2)
        except Exception:
            pass
    # 样式照抄最后一个科目行
    for i in range(n):
        r = last + 1 + i
        for c in range(2, 15):
            try:
                lead.cell(r, c)._style = _copy_style(lead.cell(last, c)._style)
            except Exception:
                pass
        lead.cell(r, 4, extra[i])
        lead.cell(r, 2, "=C2")
        lead.row_dimensions[r].height = lead.row_dimensions[last].height
    return n



def _fix_lead_after_insert(lead, lr, res):
    """插行后重算「表2 合计」与「Diff 行」。

    openpyxl 插行只搬格子，不动公式里的引用：表2 的 `=C62+C63+C64+C65` 会指到表头，
    算出来 #VALUE!；Diff 的 `=H41-H44` 会指向错位后的合计与 A3。两处都得整行重写。
    """
    rows = lr.get("rows") or {}
    if not rows:
        return
    # --- 表2 合计：范围覆盖到最后一行科目，Check 指向正确的合计行
    _rewrite_lead_disclosure(lead, lr)

    # --- Diff 行：H/J/K = 合计行 − A3 行
    a3 = None
    for r in range(1, min(lead.max_row, 80) + 1):
        if T(lead.cell(r, 7).value).replace(" ", "").upper() == "A3":
            a3 = r
            break
    tot = lr.get("total")
    if a3 and tot:
        diff = a3 + 1
        if T(lead.cell(diff, 7).value).replace(" ", "").lower() != "diff":
            diff = a3 + 1
        for col in (8, 10, 11):
            L = get_column_letter(col)
            cell = lead.cell(diff, col, "=%s%d-%s%d" % (L, tot, L, a3))
            cell.number_format = ACCT_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")


def _lead_te_ref(lead):
    """找 TE（可容忍误差）所在单元格，返回 "$C$4" 这种绝对引用。

    模板的认定阈值公式写死了 $C$4；换版模板 TE 行一挪，VLOOKUP 乘错格子，
    算出来的阈值还是"看着像数字"的错值。按标签现找，找不到再退回 C4。
    """
    from openpyxl.utils import get_column_letter as _L
    for r in range(1, min(lead.max_row, 14) + 1):
        for c in range(1, 5):
            t = T(lead.cell(r, c).value).replace(" ", "")
            if t and "可容忍误差" in t:
                for cc in range(c + 1, c + 5):
                    if lead.cell(r, cc).value not in (None, ""):
                        return "$%s$%d" % (_L(cc), r)
    return "$C$4"


def _fill_lead_cra(lead, res):
    """把 CRA 等级写进 Lead 的认定区，并把百分比烘进阈值公式。

    模板原公式 `=IF(C14="Minimal",$C$4*100%,IF(C14="Low",$C$4*75%,...))`
    把百分比固化在公式里。现在按平台配置重新生成同一串 IF，只换百分比数字：
    改配置 → 重新生成 → 公式里的百分比跟着变，**不再需要任何中间表**。
    """
    cfg = (res or {}).get("craCfg") or {}
    pcts = cfg.get("pcts") or {}
    try:
        try:
            from . import baseinfo as BI
        except ImportError:
            import baseinfo as BI
    except Exception:
        return

    te_ref = _lead_te_ref(lead)

    def _p(k, dflt):
        v = pcts.get(k)
        return dflt if v is None else v

    def _pct(k, dflt):
        return "%.6g%%" % (float(_p(k, dflt)) * 100)

    grades = cfg.get("assertions") or {}
    done = {}
    for key, kw, _label in BI.ASSERTIONS:
        row = None
        for r in range(10, min(lead.max_row, 40) + 1):
            t = T(lead.cell(r, 2).value).replace(" ", "")
            if t and kw in t:
                row = r
                break
        if not row:
            continue
        if grades.get(key):
            c = lead.cell(row, 3, grades[key])
            c.alignment = Alignment(horizontal="center", vertical="center")
        # 百分比按配置写死进公式；C 空 → 空，不造假阈值
        f = ('=IF($C{r}="","",IF($C{r}="N/A","-",'
             'IF($C{r}="Minimal",{te}*{m},IF($C{r}="Low",{te}*{l},'
             'IF($C{r}="Moderate",{te}*{mo},IF($C{r}="High",{te}*{h},""))))))').format(
                 r=row, te=te_ref,
                 m=_pct("Minimal", 1.00), l=_pct("Low", 0.75),
                 mo=_pct("Moderate", 0.50), h=_pct("High", 0.25))
        cell = lead.cell(row, 4, f)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        done[key] = row
    res["leadCra"] = {"rows": done, "teRef": te_ref,
                      "pcts": {k: _p(k, d) for k, d in
                               (("Minimal", 1.0), ("Low", 0.75),
                                ("Moderate", 0.5), ("High", 0.25))}}


#: 调整汇总表的列（B 起连续五列）——表头文字是给复核人看的，定位一律按文字找
ADJ_HEADERS = ("调整类型", "主体", "调整说明", "科目名称", "调整金额")


def _fill_lead_adj(lead, res, lr):
    """改造「调整汇总表」：五列结构 + 表1 审计调整数按科目名称汇总。

    模板原来只有 B 列一个「调整类型」表头，金额无处可填，表1 的 I 列（审计调整数）
    只能靠人手工抄。现在改成 调整类型/主体/调整说明/科目名称/调整金额 五列，
    表1 的 I 列改成 SUMIFS 按科目名称取数——填一笔调整，审定数自动跟着动。

    条件只用「科目名称」：调整类型（CLA / AA / SAD）是给人判断用的标签，
    决定这笔钱落在哪一行的是科目。
    """
    title = None
    for r in range(40, min(lead.max_row, 100) + 1):
        if "调整汇总表" in T(lead.cell(r, 2).value):
            title = r
            break
    if not title:
        return
    hdr = title + 1

    for i, txt in enumerate(ADJ_HEADERS[1:], start=3):
        src = lead.cell(hdr, 2)
        cell = lead.cell(hdr, i)
        if not T(cell.value):
            cell.value = txt
        try:
            cell._style = _copy_style(src._style)
        except Exception:
            pass
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    d0 = d1 = None
    for dv in lead.data_validations.dataValidation:
        ref = str(dv.sqref)
        if not ref.startswith("B"):
            continue
        if "管理层账表调整" in str(dv.formula1 or ""):
            part = ref.split(":")[-1]
            try:
                d0 = hdr + 1
                d1 = int("".join(ch for ch in part if ch.isdigit()))
            except ValueError:
                pass
            break
    if not d0:
        d0, d1 = hdr + 1, hdr + 4
    for r in range(d0, min(d1 + 8, lead.max_row + 1)):
        if "表2" in T(lead.cell(r, 2).value):
            d1 = max(d0, r - 1)
            break

    if not (lr and lr.get("rows")):
        return
    for subject, row in lr["rows"].items():
        f = "=SUMIFS($F${a}:$F${b},$E${a}:$E${b},$D{r})".format(a=d0, b=d1, r=row)
        cell = lead.cell(row, 9, f)
        cell.number_format = ACCT_FMT
        cell.alignment = Alignment(horizontal="right", vertical="center")

    res["leadAdj"] = {"title": title, "header": hdr, "rows": [d0, d1]}


def _lead_cols(lead):
    """按表头文字定位表1 的列号，返回 {语义: 列号}。

    为什么不写死：模板 V7 是 F..N 连续 9 列，但客户模板可能多插一列（比如「索引号」
    拆成两列），写死之后公式全部右移一格，数字看着有、其实全错位。
    """
    hdr = None
    for r in range(30, min(lead.max_row, 60) + 1):
        if T(lead.cell(r, 4).value).replace(" ", "") == "科目名称":
            hdr = r
            break
    want = (("book", ("期末账面数",)), ("adjBook", ("账表调整数",)),
            ("unaudited", ("期末未审数",)), ("adj", ("审计调整数",)),
            ("audited", ("期末审定数",)), ("prior", ("上期末审定数",)),
            ("delta", ("变动金额",)), ("deltaPct", ("变动%", "变动率")),
            ("notes", ("notes", "备注")))
    out = {}
    if hdr:
        for key, kws in want:
            for c in range(2, min(lead.max_column, 22) + 1):
                t = T(lead.cell(hdr, c).value).replace(" ", "").lower()
                if t and any(w.lower() in t for w in kws):
                    out[key] = c
                    break
    # 拿不到表头就按 V7 的列位
    dflt = {"book": 6, "adjBook": 7, "unaudited": 8, "adj": 9, "audited": 10,
            "prior": 11, "delta": 12, "deltaPct": 13, "notes": 14}
    for k, v in dflt.items():
        out.setdefault(k, v)
    return out


def _lead_flag_col(lead, row):
    """找「是否重大波动」那一列（模板无表头，只能按公式特征认）。

    模板里它写的是 =IF(AND(ABS(L37)>=$C$30,ABS(M37)>=$C$31),"是","否")，
    认「公式里同时有 是 和 否」那一格。认不到退回变动% 右邻第二列。
    """
    for c in range(2, min(lead.max_column, 22) + 1):
        v = T(lead.cell(row, c).value)
        if v.startswith("=") and "是" in v and "否" in v:
            return c
    return _lead_cols(lead).get("deltaPct", 13) + 2


def _lead_threshold_refs(lead):
    """找「波动幅度（金额）」和「波动幅度（%）」两个阈值所在的格子，返回绝对引用。

    模板把这两个值放在 C30 / C31，公式里写成 `$C$30` / `$C$31`。写死的话，
    客户模板在「波动范围」上面多写一行说明，阈值就整体下移，判断公式就指到空格，
    空白 > 0 恒不成立 → 全表一律「否」，而且看不出错。
    按 B 列标签现找，值取标签右边第一个非空格。
    """
    amp = pct = None
    for r in range(1, min(lead.max_row, 60) + 1):
        t = T(lead.cell(r, 2).value).replace(" ", "").replace("\u3000", "")
        if not t or "波动" not in t:
            continue
        for c in range(3, 9):
            if lead.cell(r, c).value not in (None, ""):
                if "%" in t:
                    pct = pct or "$%s$%d" % (get_column_letter(c), r)
                else:
                    amp = amp or "$%s$%d" % (get_column_letter(c), r)
                break
    return (amp or "$C$30", pct or "$C$31")


def _lead_flag_formula(lead, row, cols, thr):
    """「基于波动幅度判断，是否进一步调查？」的公式。

    L/M 任一为空 → 直接「否」（没有数就谈不上波动，留空反而让复核人以为漏填）；
    否则两个阈值同时触顶才算「是」。
    """
    Ld = get_column_letter(cols["delta"])
    Lp = get_column_letter(cols["deltaPct"])
    return ('=IF(OR({d}{r}="",{p}{r}=""),"否",'
            'IF(AND(ABS({d}{r})>={a},ABS({p}{r})>={b}),"是","否"))'
            ).format(d=Ld, p=Lp, r=row, a=thr[0], b=thr[1])


def _lead_code_formula(pl, row):
    """总账科目编码（C 列）链接上年 Lead 的编码列，按科目名精确取。

    IFERROR 兜「上年没这个科目」：本年多出来的科目（如大额存单）编码宁可留空，
    也不要塞一个 1002 让两行看着像同一个科目。
    """
    sheet = (pl or {}).get("sheet")
    subj = (pl or {}).get("t1Subj")
    code = (pl or {}).get("t1Code")
    if not (sheet and subj and code):
        return None
    q = "'%s'" % str(sheet).replace("'", "''")
    return ('=IFERROR(INDEX({q}!{c},MATCH($D{r},{q}!{s},0)),"")'
            ).format(q=q, c=code, s=subj, r=row)


def _lead_label_row(lead, keys, start=1, end=None):
    """从 start 往下找第一个含 keys 的行（只扫 B~F，避免打到公式里的文字）。"""
    end = end or min(lead.max_row, 300)
    for r in range(max(1, start), end + 1):
        for c in range(2, 7):
            t = T(lead.cell(r, c).value).replace(" ", "").replace("\u3000", "")
            if t and any(k in t for k in keys):
                return r
    return None


def _link_prior_disclosure(lead, rows, pl):
    """披露区几行的上期数（D 列）链到上年 Lead 表2 的同名行。

    用**宽区间**（t2SubjExt / t2AmtExt，跨过「合计」一直到表2 最后一条科目文字），
    因为「其中：存放在境外」「所有权或使用权受到限制的货币资金」都写在合计行下面，
    index_lead_layout 的表2 块到「合计」就停了，用窄区间永远匹配不到。
    """
    sheet = (pl or {}).get("sheet")
    subj = (pl or {}).get("t2SubjExt") or (pl or {}).get("t2Subj")
    amt = (pl or {}).get("t2AmtExt") or (pl or {}).get("t2Amt")
    if not (sheet and subj and amt):
        return 0
    q = "'%s'" % str(sheet).replace("'", "''")
    n = 0
    for r in rows or ():
        label = T(lead.cell(r, 2).value).replace(" ", "")
        if not label:
            continue
        c = lead.cell(r, 4, "=SUMIF(%s!%s,$B%d,%s!%s)" % (q, subj, r, q, amt))
        c.number_format = ACCT_FMT
        c.alignment = Alignment(horizontal="right", vertical="center")
        _clear_input_fill(c)
        n += 1
    return n


def _clear_seq_row_marks(lead, lr):
    """去掉「合计前那一行空行」和「合计行」在判断列上的淡黄空白高亮。

    模板给 O/P 两列铺了 containsBlanks 规则（空白格显示浅黄），本意是提示人填。
    但插行后这条例外被撑到了合计行上：合计行的 P 列本来就是空的（留给复核人定性判断），
    于是永远显示黄色；合计前那行空行也被涂上。用户要的是"干净"：
      · 只有真正的科目行（37~44）保留黄底提示；
      · 空行与合计行（含 O 列的有公式格）一律无底色。
    条件格式是"算出来的颜色"，光把 fill 设成 None 没用，必须动 sqref。
    """
    rows = lr.get("rows") or {}
    total = lr.get("total")
    if not rows or not total:
        return 0
    last = max(rows.values())
    c_flag = _lead_flag_col(lead, min(rows.values()))
    band = (last + 1, total)                      # 空行 .. 合计行
    cols = (c_flag, c_flag + 1)                   # O = 波动判断，P = 定性判断

    from openpyxl.formatting.formatting import ConditionalFormattingList
    from openpyxl.worksheet.cell_range import CellRange

    old = []
    for cf in lead.conditional_formatting:
        old.append(([cr for cr in cf.sqref.ranges], list(cf.rules)))
    new = ConditionalFormattingList()
    kept = 0
    for ranges, rules in old:
        pieces = []
        for cr in ranges:
            hit_col = cr.min_col <= cols[1] and cr.max_col >= cols[0]
            hit_row = cr.max_row >= band[0] and cr.min_row <= band[1]
            if not (hit_col and hit_row):
                pieces.append(cr.coord)
                continue
            # 与禁区重叠：拆成禁区上方、下方两段（本模板的 CF 都是纵向单列，足够）
            if cr.min_row < band[0]:
                pieces.append(CellRange(min_col=cr.min_col, min_row=cr.min_row,
                                        max_col=cr.max_col,
                                        max_row=min(cr.max_row, band[0] - 1)).coord)
            if cr.max_row > band[1]:
                pieces.append(CellRange(min_col=cr.min_col,
                                        min_row=max(cr.min_row, band[1] + 1),
                                        max_col=cr.max_col, max_row=cr.max_row).coord)
        if not pieces:
            continue
        for rule in rules:
            new.add(" ".join(pieces), rule)
        kept += 1
    lead.conditional_formatting = new

    # 双保险：把这几格的硬填充也清掉（模板若有手工涂色，一并还原成"无底色"）
    from openpyxl.styles import PatternFill
    none = PatternFill(fill_type=None)
    for r in range(band[0], band[1] + 1):
        for c in cols:
            lead.cell(r, c).fill = none
    return kept


def _rewrite_lead_total(lead, rows, total):
    """重写表1 合计行：F~变动金额 逐列 SUM，变动% 与「是否」按模板口径重算。

    插行后模板的 `=SUM(F37:F39)` 覆盖不到新行；而如果连列号也写错（把 变动%
    写到「变动金额」列上），就会变成 `L43 = IF(K43<>0, L43/K43, 1)`——
    **自己引用自己**，Excel 一打开就弹「循环引用」，整张表算不出数。
    """
    if not total or not rows:
        return
    r1, r2 = min(rows.values()), max(rows.values())
    cols = _lead_cols(lead)
    c_delta = cols["delta"]
    c_pct = cols["deltaPct"]
    c_flag = _lead_flag_col(lead, r1)
    Ld, Lp, Lk = (get_column_letter(c_delta), get_column_letter(c_pct),
                  get_column_letter(cols["prior"]))
    for col in range(cols["book"], c_delta + 1):
        L = get_column_letter(col)
        cell = lead.cell(total, col, "=SUM(%s%d:%s%d)" % (L, r1, L, r2))
        cell.number_format = ACCT_FMT
        cell.alignment = Alignment(horizontal="right", vertical="center")
    cell = lead.cell(total, c_pct,
                     '=IF(%s%d<>0,%s%d/%s%d,"")' % (Lk, total, Ld, total, Lk, total))
    cell.number_format = '0.0%'
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell = lead.cell(total, c_flag,
                     _lead_flag_formula(lead, total, cols, _lead_threshold_refs(lead)))
    cell.alignment = Alignment(horizontal="center", vertical="center")


def _rewrite_lead_disclosure(lead, lr):
    """重写表2 合计与 Check：范围覆盖到最后一行科目，Check 指向正确的合计行。

    模板 `=C62+C63+C64+C65`、`=C66-J41` 都是写死行号；插行后前者可能少算一行
    （实测「存放财务公司款项」没进合计），后者会指到表1 的错误行。
    """
    b2t = lr.get("b2Total")
    tot = lr.get("total")
    if not b2t or not tot:
        return
    r1 = min((lr.get("b2All") or lr.get("b2") or {b2t: 0}).values())
    for col, L in ((3, "C"), (4, "D")):
        cell = lead.cell(b2t, col, "=SUM(%s%d:%s%d)" % (L, r1, L, b2t - 1))
        cell.number_format = ACCT_FMT
        cell.alignment = Alignment(horizontal="right", vertical="center")
    cell = lead.cell(b2t, 5, "=C%d-J%d" % (b2t, tot))
    cell.number_format = ACCT_FMT
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell = lead.cell(b2t, 6, "=D%d-K%d" % (b2t, tot))
    cell.number_format = ACCT_FMT
    cell.alignment = Alignment(horizontal="right", vertical="center")


def _fill_lead(wb, res=None):
    """往 C.00 Lead 表打链接公式 + 填编码/账套名称：
    行号一律由 locate_lead_rows() 按科目名定位，不写死 —— V7 是 37/38/39，
    换一版模板或客户多插一行就全变，写死会静默错位。
    0. B 列（账套名称）= 客户名称 C2；C 列（总账科目编码）= 1001/1002/1012；
    1. F 列（期末账面数）用 SUMIF 按科目名称汇总 BKD 人民币列；
    2. 表2 For Disclosure 在 Lead 内部链接到表1的期末/上期审定数（J/K 列）；
    3. 境外资金行用 BKD 中非人民币账户的金额汇总；
    4. 日期行（F/K）链接到资产负债表日 C3（上期 = 上年末）。
    5. 受限资金汇总到 Lead（表1 有科目行做调整，没有则落披露行），见 _fill_lead_restricted。
    """
    if "C.00 Lead" not in wb.sheetnames or "C.00 BKD" not in wb.sheetnames:
        return
    lead = wb["C.00 Lead"]
    bkd_name = "C.00 BKD"

    # Lead 的行号（按科目名定位，不写死 37/38/39 —— 客户模板多一行就全错位）
    lr = locate_lead_rows(lead)

    # 上年 Lead 里有、本年模板没有的科目（定期存款 / 存放财务公司款项 之类）：
    # 先插行，再重新定位——不插的话上年数没处落，变动分析直接断。
    _pl = (res or {}).get("priorLead") or {}
    _extra_rows = {}
    if _pl.get("t1") and not _pl.get("error"):
        _have = set(lr["rows"].keys())
        _extra = [it["subject"] for it in _pl["t1"]
                  if it.get("subject") and it["subject"] not in _have
                  and "合计" not in it["subject"]]
        if _extra:
            _n = _lead_insert_subjects(lead, lr, _extra)
            if _n:
                lr = locate_lead_rows(lead)
                _last = max(lr["rows"].values())
                for i, subj in enumerate(_extra):
                    _extra_rows[subj] = _last + 1 + i   # 新行插在 last 之后，别再差一行
                res["priorLeadInserted"] = _n
                res["priorLeadRows"] = dict(_extra_rows)
                _fix_lead_after_insert(lead, lr, res)

    # 表2 同步补行：上年有的报表科目（定期存款、大额存单…）本年模板里没有，
    # 不插的话上年数没处落，表2 合计也会小于上年。
    if _pl.get("t2") and not _pl.get("error"):
        _have2 = set((lr.get("b2All") or lr.get("b2") or {}).keys())
        _extra2 = [it["item"] for it in _pl["t2"]
                   if it.get("item") and it["item"] not in _have2]
        if _extra2:
            _n2 = _lead_insert_disclosure(lead, lr, _extra2)
            if _n2:
                lr = locate_lead_rows(lead)
                res["priorLeadDiscInserted"] = _n2
                _fix_lead_after_insert(lead, lr, res)

    if res is not None:
        res["leadRows"] = {k: v for k, v in lr.items() if k != "warnings"}
        for w in lr.get("warnings") or []:
            res.setdefault("warnings", []).append(w)

    # BKD 的列位置与数据行范围（不写死 C/H/$500：换一份底稿列就变了）
    _bkd = wb[bkd_name]
    _bc = _locate_bkd_cols(_bkd) or {}
    _c_subj = get_column_letter(_bc.get("subject") or 3)
    _c_money = get_column_letter(_bc.get("localEnd") or 8)
    _c_cur = get_column_letter(_bc.get("currency") or 6)
    # 有「折算人民币」列就以它为准（本位币为外币的主体必须先折成人民币再上 Lead），
    # 没有折算列（旧文件/模板）就退回本币列，行为与以前一致。
    _hdr = _bc.get("_header_row") or BKD_HEADER_ROW
    _fx_rate, _fx_rmb = _locate_fx_cols(_bkd, _hdr)
    if _fx_rmb and "人民币" in T(_bkd.cell(_hdr, _fx_rmb).value):
        _c_money = get_column_letter(_fx_rmb)
    _dstart = (_bc.get("_header_row") or BKD_HEADER_ROW) + 1
    _srow = _find_sum_row(_bkd, _dstart, _bc.get("localEnd") or 8)
    _dend = (_srow - 1) if _srow and _srow - 1 >= _dstart else 500
    # 合计行上面若还有一行长文本说明，别把它算进 SUMIF 区间
    if _srow and len(T(_bkd.cell(_srow - 1, _bc.get("company") or 2).value)) > 12:
        _dend = max(_dstart, _srow - 2)

    # 插进来的上年科目并进 lr["rows"]：后面的 F 列 SUMIF、I 列 SUMIFS、合计都按它走
    lr["rows"].update(_extra_rows)

    # 0) 账套名称（B 列）= 客户名称（C2，若为空则留空让用户后填）；总账科目编码（C 列）
    #    编码优先"链接上年 Lead 的编码列"（按科目名精确取），拿不到上年 Lead 才退回内置的
    #    1001/1002/1012。写死编码的问题：客户改过科目编码、或上年新增了「存放财务公司款项」
    #    这类科目，本年就跟着错，而且看不出来。
    for subject, row in lr["rows"].items():
        bcell = lead.cell(row, 2, "=C2")            # B 列 = 账套名称 = 客户名称
        bcell.alignment = Alignment(horizontal="left", vertical="center")
        f_code = _lead_code_formula(_pl, row) if not _pl.get("error") else None
        if f_code:
            ccell = lead.cell(row, 3, f_code)
        else:
            # 上年带过来的科目（定期存款等）没有固定编码，留空让项目组填，
            # 不要硬塞一个 1002 —— 那会让两行看起来是同一个科目。
            ccell = lead.cell(row, 3, LEAD_CODE_BY_SUBJECT.get(subject, ""))
        ccell.alignment = Alignment(horizontal="center", vertical="center")

    # 1) 表1 期末账面数：F 列按科目名称 SUMIF BKD 本币
    for subject, row in lr["rows"].items():
        formula = "=SUMIF('%s'!$%s$%d:$%s$%d,\"%s\",'%s'!$%s$%d:$%s$%d)" % (
            bkd_name, _c_subj, _dstart, _c_subj, _dend, subject,
            bkd_name, _c_money, _dstart, _c_money, _dend)
        cell = lead.cell(row, 6, formula)   # F 列 = 期末账面数
        cell.number_format = ACCT_FMT
        cell.alignment = Alignment(horizontal="right", vertical="center")

    # 1.5) 科目行公式补全：插进来的上年科目只复制了样式、没有公式，
    #      不补的话「变动金额/变动%」整行空白，变动分析就断了；
    #      模板已有的行重写一遍也无害（都是同一套模板公式）。
    _cols = _lead_cols(lead)
    _flag_c = _lead_flag_col(lead, min(lr["rows"].values()))
    _thr = _lead_threshold_refs(lead)
    for subject, row in lr["rows"].items():
        _Lf = get_column_letter(_cols["book"])
        _Lg = get_column_letter(_cols["adjBook"])
        _Lh = get_column_letter(_cols["unaudited"])
        _Li = get_column_letter(_cols["adj"])
        _Lj = get_column_letter(_cols["audited"])
        _Lk = get_column_letter(_cols["prior"])
        _Ll = get_column_letter(_cols["delta"])
        _Lm = get_column_letter(_cols["deltaPct"])
        for c, f, fmt in (
                (_cols["unaudited"], "=%s%d+%s%d" % (_Lf, row, _Lg, row), ACCT_FMT),
                (_cols["audited"], "=%s%d+%s%d" % (_Lh, row, _Li, row), ACCT_FMT),
                (_cols["delta"], "=%s%d-%s%d" % (_Lj, row, _Lk, row), ACCT_FMT),
                # 上期为 0 时留空：模板原本回落成 1（100%），
                # 本期上期都是 0 的行会显示「增长 100%」，是假数字。
                (_cols["deltaPct"], '=IF(%s%d<>0,%s%d/%s%d,"")'
                 % (_Lk, row, _Ll, row, _Lk, row), '0.0%'),
                (_flag_c, _lead_flag_formula(lead, row, _cols, _thr), None)):
            cell = lead.cell(row, c, f)
            if fmt:
                cell.number_format = fmt
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # 2) 表2 For Disclosure：Lead 内部链接到表1的审定数（J=期末审定，K=上期末审定）
    #    表1科目在 D37/D38/D39，表2科目在 B62/B63/B64
    for subject, r2 in (lr.get("b2All") or lr["b2"]).items():
        r1 = lr["rows"].get(subject)
        if not r1:
            # 本年没有这个科目：表2 这一行的本期数填 0，**不留空**。
            # 上年有「存放财务公司款项」、本年没有，是"确实为 0"，不是"不知道"；
            # 留空和填 0 在对账时长得一样，但留空会让 Table 2 的合计与表1 差一截，
            # 复核人得逐个猜哪一格是漏填。上期数照旧由下面的 SUMIF 去上年 Lead 取。
            c_zero = lead.cell(r2, 3, 0)
            c_zero.number_format = ACCT_FMT
            c_zero.alignment = Alignment(horizontal="right", vertical="center")
            _clear_input_fill(c_zero)
            continue
        # C 列 = 期末审定数 = 表1 J 列
        c_cur = lead.cell(r2, 3, "=J%d" % r1)
        c_cur.number_format = ACCT_FMT
        c_cur.alignment = Alignment(horizontal="right", vertical="center")
        _clear_input_fill(c_cur)
        # D 列 = 上期审定数 = 表1 K 列
        c_prior = lead.cell(r2, 4, "=K%d" % r1)
        c_prior.number_format = ACCT_FMT
        c_prior.alignment = Alignment(horizontal="right", vertical="center")
        _clear_input_fill(c_prior)

    # 2.5) 上年数：K 列与表2 D 列都链到「上年 C.00 Lead」，按科目名 SUMIF
    if _pl.get("t1Subj") and not _pl.get("error"):
        _sheet = _pl.get("sheet") or PRIOR_LEAD_SHEET
        _sr, _ar = _pl["t1Subj"], _pl["t1Amt"]
        for subject, row in lr["rows"].items():
            c = lead.cell(row, 11, "=SUMIF('%s'!%s,$D%d,'%s'!%s)" % (_sheet, _sr, row, _sheet, _ar))
            c.number_format = ACCT_FMT
            c.alignment = Alignment(horizontal="right", vertical="center")
        if _pl.get("t2Subj"):
            _s2, _a2 = _pl["t2Subj"], _pl["t2Amt"]
            for subject, r2 in (lr.get("b2All") or lr["b2"]).items():
                c = lead.cell(r2, 4, "=SUMIF('%s'!%s,$B%d,'%s'!%s)" % (_sheet, _s2, r2, _sheet, _a2))
                c.number_format = ACCT_FMT
                c.alignment = Alignment(horizontal="right", vertical="center")
                _clear_input_fill(c)


    # 3) 披露区两行：「其中：存放在境外的款项总额」「所有权或使用权受到限制的货币资金」
    #    本期数（C 列）一律留空——本年没有可靠的取数口径，宁缺勿造；
    #    上期数（D 列）链接上年 Lead 表2 的同名行，口径与本表其他行一致。
    _disc_rows = []
    _ov = lr.get("overseas")
    if _ov:
        lead.cell(_ov, 3, None)
        _disc_rows.append(_ov)
    _rs = _lead_label_row(lead, ("受到限制的货币资金",), (lr.get("b2Total") or 0) + 1)
    if _rs:
        lead.cell(_rs, 3, None)
        _disc_rows.append(_rs)
    _link_prior_disclosure(lead, _disc_rows, _pl)

    # 4) 日期单元格：F35 本期 = C3（资产负债表日）；K35 上期 = 上年 12/31
    _drow = lr["date"] or 35
    lead.cell(_drow, 6, "=C3")                       # F 本期期末余额日期
    lead.cell(_drow, 11, "=DATE(YEAR($C$3)-1,12,31)")  # K 上期期末余额日期

    # 4.5) A3 行（审计确认数）：**留空，由项目组人工填**（用户 2026-11-07 口径）。
    #      以前这里默认链到合计行（H←未审、J←审定、K←上期审定），本意是让下面
    #      Diff 行归零好看；但那等于程序替审计师签了"这三个数我确认过了"。
    #      Diff 公式保留：人工填完自动归零，没填时显示"合计 − 空"，就是一条待办。
    #      只设数字格式，方便直接输入；已有值（人工填过再生成）不动。
    for _r in range(1, min(lead.max_row, 60) + 1):
        if T(lead.cell(_r, 7).value).replace(" ", "").upper() == "A3":
            for _c in (8, 10, 11):          # H / J / K
                _tgt = lead.cell(_r, _c)
                _tgt.number_format = ACCT_FMT
                _tgt.alignment = Alignment(horizontal="right", vertical="center")
            break

    # 5) CRA 认定区：等级写进 C 列，阈值公式改成查 CRA 表（百分比可配）
    _fill_lead_cra(lead, res)

    # 5.5) 调整汇总表：五列结构 + 表1 审计调整数按科目名称 SUMIFS
    _fill_lead_adj(lead, res, lr)

    # 5.8) 合计行收尾：不管有没有上年数据都要重算——模板的 =SUM(F37:F39)
    #      固定三行，V7 模板里表2 还有「存放财务公司款项」那一行不在旧公式里。
    _rewrite_lead_total(lead, lr["rows"], lr.get("total"))
    _rewrite_lead_disclosure(lead, lr)

    # 6) 受限资金：从 BKD 汇总到 Lead（表1 有科目行就做调整行，没有就落披露行）
    _fill_lead_restricted(lead, _bkd, _bc, _dstart, _dend, res)

    # 7) 收尾：合计前那行空行与合计行的 O/P 列不要淡黄底——只有真科目行才该有提示
    _clear_seq_row_marks(lead, lr)

    # 8) 金额格统一会计格式（两位小数、零显示「-」）。只改数字格式，
    #    字体/边框/底色/对齐保持模板原样——模板观感归项目组，程序只统一数字怎么显示。
    _paint_lead_money_fmt(lead)

    # 9) 回填 BKD 勾稽行的 Lead 引用。
    #    必须等 Lead 写完再算：BKD 先填时 Lead 的科目行/合计行还没落，布局识别不到，
    #    只能沿用模板写死的 F41（V7 的行号），客户模板插过行就指错。
    _fix_bkd_lead_check(wb)

    # 10) 蓝框补齐：插行出来的新行不继承邻行填充，右侧蓝竖条会在那几行断开，
    #     看上去就是"边缘蓝格被截断"。这里按当前实际行数把右竖边与底横边补满。
    _paint_outer_frame(lead)


def _fix_bkd_lead_check(wb):
    """把 C.00 BKD 勾稽行的 'C.00 Lead'!Fx 改成**按当前 Lead 布局**算出来的格。

    勾稽 = BKD 折算人民币合计 − Lead 期末账面数合计；两边都是人民币口径。
    模板里的 F41 是 V7 行号，客户模板多一行（瑞云合计在 46 行）就指到别的科目上，
    结果看着"勾稽不平"，其实是公式指错了格。
    """
    if "C.00 BKD" not in wb.sheetnames or "C.00 Lead" not in wb.sheetnames:
        return
    lead, bkd = wb["C.00 Lead"], wb["C.00 BKD"]
    lay = (index_lead_layout(lead) or {}).get("lead1") or {}
    tot = lay.get("total")
    bcol = (lay.get("cols") or {}).get("book")
    if not (tot and bcol):
        lr = locate_lead_rows(lead) or {}
        tot = tot or lr.get("total")
        bcol = bcol or 6                       # 定位失败：账面数列在 F（模板固定语义）
    if not tot:
        return
    col = _locate_bkd_cols(bkd)
    if not col:
        return
    srow = _find_sum_row(bkd, (col.get("_header_row") or 9) + 1, col["localEnd"])
    if not srow:
        return
    _r0, c_rmb0 = _locate_fx_cols(bkd, col["_header_row"])
    mcol = c_rmb0 or col["localEnd"]
    Lrm = get_column_letter(mcol)
    # 勾稽格不一定在折算列：模板常把它放在「本币」列，公式里却去减 Lead。
    # 所以在勾稽行上按内容找那一格，找到就原地重写，不按列猜。
    row = srow + 2
    for c in range(1, (bkd.max_column or 1) + 1):
        cell = bkd.cell(row, c)
        if isinstance(cell.value, str) and "C.00 Lead" in cell.value:
            cell.value = "=%s%d-'C.00 Lead'!%s%d" % (
                Lrm, srow, get_column_letter(bcol), tot)
            break


#: 受限资金在 Lead 表1 里可能用的科目名（模板语义固定，行号不固定）
RESTRICTED_SUBJECT_NAMES = ("受限资金", "受限货币资金", "使用受限的货币资金")


def _fill_lead_restricted(lead, bkd_ws, bkd_col, dstart, dend, res):
    """把 BKD「受限资金」列汇总到 Lead。

    Skill §4.3 的口径是「Lead 用审计调整从其他货币资金重分类到受限资金」，
    但那要求模板表1 里有「受限资金」科目行。V7 模板只有 37/38/39 三行，
    **没有这一行**——按 Skill §7「先识别模板形态，再决定写入内容」，
    这种形态只能在披露区（「所有权或使用权受到限制的货币资金」）列示。

    两种形态都实现，按模板实际结构选，并在 res 里记明用了哪种——
    否则复核人看不出「为什么这张底稿没有重分类调整」。
    """
    if res is None:
        return
    c_res = _locate_restricted_col(bkd_ws, bkd_col["_header_row"])
    sum_cell = res.get("restrictedSumCell") or ""
    ref = ("'C.00 BKD'!%s" % sum_cell) if sum_cell else \
          "SUM('C.00 BKD'!%s%d:%s%d)" % (get_column_letter(c_res), dstart,
                                         get_column_letter(c_res), dend)
    mode, target = "", None

    # ---- 形态 A：表1 里有「受限资金」科目行 → 做审计调整 ----
    subj_row = None
    for r in range(36, min(lead.max_row, 60) + 1):
        nm = T(lead.cell(r, 4).value).replace(" ", "")       # D 列 = 科目名称
        if nm and any(k == nm or k in nm for k in RESTRICTED_SUBJECT_NAMES):
            subj_row = r
            break
    if subj_row:
        mode = "reclass"
        lead.cell(subj_row, 6, 0).number_format = ACCT_FMT      # F 账面数：不是独立账面科目
        c_adj = lead.cell(subj_row, 9, "=" + ref)                 # I 审计调整：重分类进来
        c_adj.number_format = ACCT_FMT
        c_adj.alignment = Alignment(horizontal="right", vertical="center")
        # 调整汇总表**不写任何文字**：那是给项目组手工用的，程序自动塞一条
        # 「重分类 / 受限资金 / 金额 / 说明」进去，五列会串成一行看不懂的东西，
        # 而且表1 的 I 列已直接链接 BKD，不需要再借调整表转一手。
        target = "表1 第 %d 行（受限资金科目行）" % subj_row
    else:
        # ---- 形态 B：表1 无该行 → 落披露区 ----
        drow = _lead_label_row(lead, ("受到限制的货币资金",), 60)
        if drow is None:
            res["restrictedMode"] = "none"
            res["restrictedNote"] = ("Lead 里既没有「受限资金」科目行，也没有"
                                     "「所有权或使用权受到限制的货币资金」披露行，受限金额未上 Lead")
            return
        mode = "disclosure"
        # 披露行的本期数保持空白（用户口径：本期不填，上期数链上年 Lead），
        # 这里只登记"该行的数应落到哪里"，不往 C 列写公式。
        target = "披露区 第 %d 行（所有权或使用权受到限制的货币资金）" % drow

    res["restrictedMode"] = mode
    res["restrictedTarget"] = target
    res["restrictedCell"] = ref
    res["restrictedTotal"] = r2(res.get("restrictedTotal") or 0)


def _find_cutoff_sheet(wb):
    """找 C.03 Cutoff 工作表（容忍空格/大小写/后缀差异）。"""
    if CUTOFF_SHEET in wb.sheetnames:
        return CUTOFF_SHEET
    for n in wb.sheetnames:
        k = re.sub(r"[\s　]", "", n).lower()
        if k.startswith("c.03") or "cutoff" in k:
            return n
    return None


def _cutoff_blocks(ws):
    """探测 C.03 里的样本数据块：返回 [(表头行, 数据起始行)]，按模板顺序（期末前、期末后）。

    模板表头行的 B 列固定写着「样本编号」，据此刻位，不依赖硬编码行号。
    """
    out = []
    last = min(ws.max_row, 300)
    for r in range(1, last + 1):
        if T(ws.cell(r, 2).value).replace(" ", "") == "样本编号":
            out.append((r, r + 1))
    return out


def _shift_row_dims(ws, insert_at, n):
    """insert_rows 只搬单元格、不搬行高，这里把 insert_at 及以下的行高整体下移 n 行。"""
    items = []
    for r, d in list(ws.row_dimensions.items()):
        if r >= insert_at and (d.height is not None or d.hidden):
            items.append((r, d.height, d.hidden))
    for r, _, _ in items:
        d = ws.row_dimensions[r]
        d.height = None
        d.hidden = False
    for r, h, hidden in sorted(items, reverse=True):
        nd = ws.row_dimensions[r + n]
        nd.height = h
        nd.hidden = hidden


def _grow_cutoff_block(ws, header_row, need):
    """保证表头下方至少 need 行数据位：超出模板预置行数时在块尾插入行并复制样式。

    返回插入的行数。
    """
    extra = int(need) - CUTOFF_PRESET_ROWS
    if extra <= 0:
        return 0
    insert_at = header_row + 1 + CUTOFF_PRESET_ROWS      # 预置数据的下一行
    # 插入点之后的合并单元格需要手动下移（本表合并区在上方，稳妥处理）
    stale = []
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row >= insert_at:
            stale.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
            ws.unmerge_cells(str(mc))
    ws.insert_rows(insert_at, extra)
    for (r0, r1, c0, c1) in stale:
        ws.merge_cells(start_row=r0 + extra, start_column=c0,
                       end_row=r1 + extra, end_column=c1)
    _shift_row_dims(ws, insert_at, extra)
    # 以该块最后一行数据位为样板，复制样式（边框/字体/对齐/数字格式）与行高
    src = insert_at - 1
    max_col = min(ws.max_column or CUTOFF_CHECK_COL, 20)
    for i in range(extra):
        dst = insert_at + i
        h = ws.row_dimensions[src].height
        if h:
            ws.row_dimensions[dst].height = h
        for col in range(1, max_col + 1):
            s = ws.cell(src, col)
            d = ws.cell(dst, col)
            if s.has_style:
                d._style = _copy_style(s._style)
    return extra


def _retarget_cutoff_zones(ws, block_map, shifts):
    """同步条件格式 / 数据验证的覆盖范围。

    block_map: [(旧首行, 旧末行, 新首行, 新末行)]，正好是数据块的范围时整块替换；
    shifts:    [(插入点, 插入行数)]，其余区域按其起始行整体下移。
    """
    def _map_row(r):
        for at, d in shifts:
            if r >= at:
                r += d
        return r

    def _remap(rng_text):
        cr = CellRange(rng_text)
        for o0, o1, n0, n1 in block_map:
            if cr.min_row == o0 and cr.max_row == o1:
                return "%s%d:%s%d" % (get_column_letter(cr.min_col), n0,
                                      get_column_letter(cr.max_col), n1)
        delta = _map_row(cr.min_row) - cr.min_row
        if delta:
            cr.shift(row_shift=delta)
        return str(cr)

    # 条件格式：ConditionalFormatting 以 sqref 参与 hash，直接改 sqref 会让 dict 键失效，
    # 所以先快照、清空，再按新范围重新注册。
    cf = ws.conditional_formatting
    snap = [(str(key.sqref), list(rules)) for key, rules in list(cf._cf_rules.items())]
    cf._cf_rules.clear()
    prio = 1
    for sqref_text, rules in snap:
        for part in sqref_text.split():
            new_part = _remap(part)
            for rule in rules:
                try:
                    nr = _copy.deepcopy(rule)
                    nr.priority = prio
                    prio += 1
                    cf.add(new_part, nr)
                except Exception:
                    pass
    # 数据验证（N 列「是否跨期」下拉）
    for dv in (ws.data_validations.dataValidation if ws.data_validations else []):
        parts = str(dv.sqref).split()
        try:
            dv.sqref = MultiCellRange(" ".join(_remap(p) for p in parts))
        except Exception:
            pass


def _fill_cutoff(wb, cutoff):
    """往 C.03 Cutoff 填截止性测试结果。

    表1.1：B7 测试阈值（覆盖模板公式）、B8 截止期间（资产负债表日）、B11 测试说明；
    表1.2：期末前样本、期末后样本。样本数超过模板预置的 5 行时，自动在块尾插入行，
           并把插入行的样式、行高、P 列 Check（=M+I）、N 列下拉、条件格式同步补齐，
           保证「全部结果」都写进底稿，不再截断。
    """
    if not cutoff:
        return
    sheet = _find_cutoff_sheet(wb)
    if not sheet:
        return
    ws = wb[sheet]
    st = cutoff.get("stat") or {}

    # ---- 表1.1：阈值 / 截止期间 / 说明
    thr = st.get("threshold")
    if thr and float(thr) > 0:
        c = ws.cell(7, 2, r2(float(thr)))
        c.number_format = ACCT_FMT
        c.alignment = Alignment(horizontal="right", vertical="center")
    end_d = _as_date(st.get("endDate"))
    if end_d:
        c = ws.cell(8, 2, end_d)
        c.number_format = 'yyyy-mm-dd'
        c.alignment = Alignment(horizontal="center", vertical="center")

    if st and cutoff.get("source") != "flow":
        # 样本应取自银行流水。本期没有流水数据源 → 只写参数与缺口，不编样本。
        txt = (
            "我们以 %s 为资产负债表日，选取资产负债表日前后各 %d 天 / %d 天"
            "（%s 至 %s）作为截止期间，以单笔收付金额 %s 元为测试阈值。\n"
            "本期尚未取得银行流水，本表样本区留空。截止性测试的样本必须取自银行流水"
            "（对账单口径的收付款双方、账号与到账日期），不从序时账推导：序时账只有科目与金额，"
            "靠凭证号或金额凑出的配对是推测，无法证明资金实际进出银行的时间。\n"
            "取得流水后按上述期间与阈值重新生成即可自动填充。\n"
            "（请在此补充所使用截止期间的理由，以及识别被认为异常或意料之外的银行收付的依据。）"
        ) % (
            st.get("endDate", ""), st.get("beforeDays", 0), st.get("afterDays", 0),
            st.get("windowStart", ""), st.get("windowEnd", ""),
            format(st.get("threshold", 0), ",.2f"),
        )
        c = ws.cell(11, 2, txt)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c.font = Font(name="Arial", size=9)
    elif st:
        # 将来接入流水后的口径：样本来自流水，不再提序时账配对
        txt = (
            "我们以 %s 为资产负债表日，选取其前后各 %d 天 / %d 天（%s 至 %s）作为截止期间，"
            "以单笔收付金额 %s 元为测试阈值，从银行流水中识别截止日前后的大额收付与跨期收付"
            "共 %d 笔（期末前 %d 笔、期末后 %d 笔，其中跨期 %d 笔）。\n"
            "（请在此补充所使用截止期间的理由，以及识别被认为异常或意料之外的银行收付的依据。）"
        ) % (
            st.get("endDate", ""), st.get("beforeDays", 0), st.get("afterDays", 0),
            st.get("windowStart", ""), st.get("windowEnd", ""),
            format(st.get("threshold", 0), ",.2f"),
            st.get("transferPairs", 0), st.get("beforeCount", 0), st.get("afterCount", 0),
            st.get("crossCount", 0),
        )
        c = ws.cell(11, 2, txt)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c.font = Font(name="Arial", size=9)

    # ---- 表1.2：样本行
    money_fmt = ACCT_FMT
    txt_fmt = '@'
    date_fmt = 'yyyy-mm-dd'

    blocks = _cutoff_blocks(ws)
    before_rows = cutoff.get("before") or []
    after_rows = cutoff.get("after") or []
    nb, na = len(before_rows), len(after_rows)

    if blocks:
        b_hdr, b_start = blocks[0]
        a_hdr, a_start = blocks[1] if len(blocks) > 1 else (None, None)
    else:
        b_hdr, b_start, a_hdr, a_start = CUTOFF_BEFORE_START - 1, CUTOFF_BEFORE_START, None, None

    # ---- 列按表头文字定位，不按列号 ----
    # 写死 B~P 的后果：模板多插一列，收款方就写进付款方；
    # 而 Check 列 =M+I 仍然配平（一正一负相加还是 0），账面全对、错得没有声音。
    # 定位不到必需列时**中止写样本**，宁可空着让人看见，也不写到隔壁列。
    cmap, cmiss = (locate_cols(ws, b_hdr, C03_COLS) if b_hdr else ({}, [c[0] for c in C03_COLS]))
    col_bad = [k for k in C03_REQUIRED if k not in cmap]
    if col_bad:
        st["colError"] = ("C.03 表头识别失败，未写入样本：找不到 %s（表头行第 %s 行）"
                          % ("、".join(col_bad), b_hdr))
        if b_start:
            c = ws.cell(b_start, 2, "（C.03 表头识别失败，未写入样本）%s" % st["colError"])
            c.font = Font(name="Arial", size=9, italic=True, color="C00000")
        return

    # ---- 按需插行（先期末前、后期末后，行号随插入位移）
    old_b_end = b_start + CUTOFF_PRESET_ROWS - 1
    old_a_end = (a_start + CUTOFF_PRESET_ROWS - 1) if a_start else None
    ins_b = _grow_cutoff_block(ws, b_hdr, nb) if nb > CUTOFF_PRESET_ROWS else 0
    if a_start:
        a_hdr += ins_b
        a_start += ins_b
    ins_a = _grow_cutoff_block(ws, a_hdr, na) if (a_hdr and na > CUTOFF_PRESET_ROWS) else 0

    new_b_end = b_start + max(nb, CUTOFF_PRESET_ROWS) - 1
    new_a_start = a_start
    new_a_end = (a_start + max(na, CUTOFF_PRESET_ROWS) - 1) if a_start else None
    if ins_b or ins_a:
        zone_map = [(b_start, old_b_end, b_start, new_b_end)]
        if a_start and old_a_end:
            # 旧坐标（模板原始）= a_start 插入前的位置
            zone_map.append((a_start - ins_b, old_a_end, new_a_start, new_a_end))
        shifts = []
        if ins_b:
            shifts.append((b_hdr + 1 + CUTOFF_PRESET_ROWS, ins_b))
        if ins_a:
            shifts.append((a_hdr + 1 + CUTOFF_PRESET_ROWS, ins_a))
        _retarget_cutoff_zones(ws, zone_map, shifts)

    def _write(rows, start_row):
        if not start_row:
            return
        c_pay, c_recv = cmap.get("payAmt"), cmap.get("recvAmt")
        c_acct1, c_acct2 = cmap.get("payAcct"), cmap.get("recvAcct")
        c_d1, c_d2 = cmap.get("payTime"), cmap.get("recvDate")
        c_ctr = (cmap.get("no"), cmap.get("currency"), cmap.get("cross"), cmap.get("check"))
        for i, s in enumerate(rows or []):
            r = start_row + i
            vals = [
                (cmap.get("no"), s.get("no", "")),
                (cmap.get("company"), s.get("company", "")),
                (cmap.get("payBank"), s.get("payBank", "")),
                (cmap.get("payAcct"), s.get("payAcct", "")),
                (cmap.get("payTime"), _as_date(s.get("payDate"))),
                (cmap.get("vno"), s.get("vno", "")),
                (cmap.get("currency"), s.get("currency", "")),
                (cmap.get("payAmt"), s.get("payAmt")),
                (cmap.get("recvBank"), s.get("recvBank", "")),
                (cmap.get("recvAcct"), s.get("recvAcct", "")),
                (cmap.get("recvDate"), _as_date(s.get("recvDate"))),
                (cmap.get("recvAmt"), s.get("recvAmt")),
                (cmap.get("cross"), s.get("cross", "")),
                (cmap.get("note"), s.get("note", "")),
                (cmap.get("check"), "=%s%d+%s%d" % (get_column_letter(c_recv), r,
                                                    get_column_letter(c_pay), r)),
            ]
            for col, v in vals:
                if col is None:
                    continue
                if v is None:
                    v = ""
                cell = ws.cell(r, col, v)
                if col in (c_pay, c_recv):
                    cell.number_format = money_fmt
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col in (c_acct1, c_acct2):
                    cell.number_format = txt_fmt        # 账号保持文本，防前导零丢失
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif col in (c_d1, c_d2):
                    cell.number_format = date_fmt
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col in c_ctr:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    _write(before_rows, b_start)
    _write(after_rows, a_start)

    # 插入行后，模板原有 P 列公式（=M20+I20 之类）行号不会自动跟随，整块重写一遍
    _chk = cmap.get("check")
    _cpay, _crecv = cmap.get("payAmt"), cmap.get("recvAmt")
    for _start, _cnt in ((b_start, max(nb, CUTOFF_PRESET_ROWS) if b_start else 0),
                         (a_start, max(na, CUTOFF_PRESET_ROWS) if a_start else 0)):
        if not _start or not _chk or not _cpay or not _crecv:
            continue
        for i in range(_cnt):
            r = _start + i
            cell = ws.cell(r, _chk)
            cell.value = "=%s%d+%s%d" % (get_column_letter(_crecv), r,
                                         get_column_letter(_cpay), r)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    if st:
        st["shownBefore"] = nb
        st["shownAfter"] = na
        st["insertedRows"] = ins_b + ins_a

    # 一笔样本都没抓到：在样本区首行写明原因，避免"生成了却不知道为什么不填数"。
    # 但「还没接入流水」这种情况不往表里写——样本区保持干净，理由写在 B11 说明里，
    # 否则交出去的底稿样本表第一行挂着一句红字，看着像出了错。
    if not (before_rows or after_rows) and cutoff.get("source") == "flow":
        reason = cutoff.get("warn") or _cutoff_empty_reason(st)
        c = ws.cell(b_start, 2, "（本次未识别到可填入的转账样本）" + reason)
        c.font = Font(name="Arial", size=9, italic=True, color="C00000")
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _cutoff_no_flow_reason(st):
    """没有银行流水时的提示：说清缺什么、怎么补，不暗示"期间填错了"。"""
    st = st or {}
    return ("本表样本需取自银行流水，本期未取得流水数据，样本区留空"
            "（截止期间 %s 至 %s，单笔金额 ≥ %s 元）。取得流水后重跑即自动填充。"
            % (st.get("windowStart", "?"), st.get("windowEnd", "?"),
               format(st.get("threshold", 0) or 0, ",.2f")))


def _cutoff_empty_reason(st):
    """样本为空时给出可执行的排查提示。"""
    st = st or {}
    jr = int(st.get("journalRows") or 0)
    jf, jt = st.get("journalFrom"), st.get("journalTo")
    _rng = ("序时账实际只覆盖 %s 至 %s。" % (jf, jt)) if (jf and jt) else ""
    if jr == 0:
        return ("截止期间（%s 至 %s）内没有筛到任何货币资金科目（库存现金/银行存款/其他货币资金）的序时账行。%s"
                "请确认序时账已导入完整期间、日期字段映射正确、截止日与前后天数覆盖到实际业务日期。" % (
                    st.get("windowStart", "?"), st.get("windowEnd", "?"), _rng))
    pairs = int(st.get("transferPairs") or 0)
    if pairs == 0:
        return ("期间内货币资金科目共 %d 行，但没有配对成「一借一贷」的银行间转账"
                "（需同一凭证号且金额相等，或日期分处截止日两侧且金额相等）："
                "请确认序时账的凭证号、借贷金额已正确映射。" % jr)
    return ("期间内识别到 %d 笔转账，但全部低于测试阈值 %s 元，未纳入样本："
            "请下调阈值或扩大测试期间。" % (pairs, format(st.get("threshold", 0), ",.2f")))


def _legacy_build_workbook(res, tpl_headers=None, tpl_widths=None, meta=None, highlight_new=True, mark_new=True):
    """旧版：从零生成 BKD + 检查表（保留作回退）。"""
    meta = meta or {}
    cols = _resolve_tpl_cols(tpl_headers, tpl_widths)
    wb = Workbook()
    ws = wb.active
    ws.title = "BKD"

    # 样式
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_font = Font(name="宋体", size=10, bold=True)
    head_fill = PatternFill("solid", fgColor="D9E2F3")
    txt_font = Font(name="宋体", size=10)
    new_fill = PatternFill("solid", fgColor="FFF2CC")
    miss_fill = PatternFill("solid", fgColor="FCE4E4")
    sect_font = Font(name="宋体", size=10, bold=True, color="1F3864")
    sect_fill = PatternFill("solid", fgColor="F2F2F2")
    title_font = Font(name="宋体", size=15, bold=True)
    sub_font = Font(name="宋体", size=9, color="666666")
    tot_font = Font(name="宋体", size=10, bold=True)

    header_row = 1
    # 列宽
    for i, (_, h, w) in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w if (isinstance(w, (int, float)) and w >= 10) else 14
    # 表头
    ws.row_dimensions[header_row].height = 42
    for c, (_, h, w) in enumerate(cols, start=1):
        cell = ws.cell(header_row, c, T(h) or " ")
        cell.font = head_font
        cell.fill = head_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 数据
    for idx, row in enumerate(res["rows"]):
        r = header_row + 1 + idx
        is_new = not row["_matched"] and highlight_new
        for c, (field, h, w) in enumerate(cols, start=1):
            cell = ws.cell(r, c)
            cell.border = border
            cell.font = txt_font
            if is_new:
                cell.fill = new_fill
            val = ""
            isnum = False
            if field == "company":
                val = row["company"]
            elif field == "subject":
                val = row["subject"]
            elif field == "bankName":
                val = row["bankName"]
            elif field == "acctNo":
                val = row["acctNo"]
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif field == "currency":
                val = row["currency"]
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif field in ("origEnd", "endAny"):
                val = row["orig"]
                isnum = True
            elif field == "localEnd":
                val = row["local"]
                isnum = True
            elif field == "nature":
                val = row["nature"]
            elif field == "confirm":
                val = row["confirm"]
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif field == "notes":
                val = row["notes"] or ""
                if not row["_matched"] and mark_new:
                    val = "【新增·上年底稿无此账户】" + ((" " + val) if val else "")
            elif field == "jNote":
                val = "N/a" if row["confirm"] == "是" else ""
            if isnum:
                cell.value = val
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.value = val
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    ws.freeze_panes = "A%d" % (header_row + 1)
    if res["rows"]:
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cols)), header_row + len(res["rows"]))

    # ---- 检查 sheet ----
    _build_check_sheet(wb, res, meta, rate_threshold=0.5, amt_threshold=100000)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _build_check_sheet(wb, res, meta, rate_threshold, amt_threshold, cutoff=None):
    st = res["stats"]
    ws = wb.create_sheet("检查")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_font = Font(name="宋体", size=10, bold=True)
    head_fill = PatternFill("solid", fgColor="D9E2F3")
    txt_font = Font(name="宋体", size=10)
    num_fmt = MONEY_FMT
    title_font = Font(name="宋体", size=15, bold=True)
    sub_font = Font(name="宋体", size=9, color="666666")
    sect_font = Font(name="宋体", size=10, bold=True, color="1F3864")
    sect_fill = PatternFill("solid", fgColor="F2F2F2")
    tot_font = Font(name="宋体", size=10, bold=True)
    new_fill = PatternFill("solid", fgColor="FFF2CC")
    miss_fill = PatternFill("solid", fgColor="FCE4E4")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    widths = [38, 26, 24, 22, 20, 10, 18, 16, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 1
    c = ws.cell(r, 1, "货币资金 BKD 生成检查表")
    c.font = title_font
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    ws.row_dimensions[r].height = 30
    r += 1
    meta_txt = "生成时间：%s　　本年科目余额表：PTW平台自动取数（%s）" % (
        meta.get("time", ""), meta.get("tbDb", ""))
    c = ws.cell(r, 1, meta_txt)
    c.font = sub_font
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    r += 2

    def sect(txt):
        nonlocal r
        c = ws.cell(r, 1, txt)
        c.font = sect_font
        c.fill = sect_fill
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        r += 1

    def head(cells):
        nonlocal r
        for i, t in enumerate(cells, start=1):
            c = ws.cell(r, i, t)
            c.font = head_font
            c.fill = head_fill
            c.border = border
            c.alignment = center
        ws.row_dimensions[r].height = 30
        r += 1

    # 一、汇总
    sect("一、匹配情况汇总")
    head(["项目", "数量", "金额（本币）", "", "", "说明", "", "", ""])
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=5)
    ws.merge_cells(start_row=r - 1, start_column=6, end_row=r - 1, end_column=9)
    sum_rows = [
        ("上年BKD底稿账户数", st["priorCount"], st["sumPrior"], "去重后账户数" + ("" if not st["dupPrior"] else "（原始%d条，重复%d条已合并）" % (st["priorRawCount"], st["dupPrior"]))),
        ("本年科目余额表账户数", st["tbCount"], st["sumTb"], "PTW平台自动取数"),
        ("匹配成功账户数", st["matched"], None, "其中：唯一主键一致 %d 户；账号一致 %d 户；按公司+科目 %d 户"
         % (st.get("byKey", 0), st.get("byAcct", 0), st.get("bySubject", 0))),
        ("本年新增（上年底稿未找到）", st["newCount"], r2(sum(x["local"] for x in res["news"])), "已列示于第三部分，需补录底稿信息"),
        ("上年有、本年无", st["missCount"], r2(sum(x["local"] for x in res["miss"])), "已列示于第四部分，需核实是否销户/漏记"),
    ]
    for row in sum_rows:
        c = ws.cell(r, 1, row[0]); c.font = txt_font; c.border = border
        c = ws.cell(r, 2, row[1]); c.font = txt_font; c.border = border; c.alignment = center
        c = ws.cell(r, 3, row[2] if row[2] is not None else ""); c.border = border
        if row[2] is not None:
            c.number_format = num_fmt; c.alignment = right
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        c = ws.cell(r, 6, row[3]); c.font = txt_font; c.border = border; c.alignment = left
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=9)
        r += 1

    # 二、金额核对
    r += 1
    sect("二、金额核对（期末余额本币合计）")
    head(["项目", "上年BKD合计", "本年余额表合计", "差异", "说明", "", "", "", ""])
    ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1, end_column=9)
    c = ws.cell(r, 1, "期末余额（本币）合计"); c.font = txt_font; c.border = border
    c = ws.cell(r, 2, st["sumPrior"]); c.number_format = num_fmt; c.border = border; c.alignment = right
    c = ws.cell(r, 3, st["sumTb"]); c.number_format = num_fmt; c.border = border; c.alignment = right
    c = ws.cell(r, 4, st["diff"]); c.number_format = num_fmt; c.border = border; c.alignment = right
    c = ws.cell(r, 5, "一致，无差异" if abs(st["diff"]) < 0.005 else "存在差异，主要系本年新增/销户账户及上年账户余额变动所致")
    c.font = txt_font; c.border = border; c.alignment = left
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
    r += 1

    # 三、本年有、上年无
    r += 1
    sect("三、本年科目余额表有、上年BKD底稿未找到的账户（需补充底稿信息）")
    head(["公司名称", "账号", "科目名称", "银行/机构名称", "币种", "期末余额（本币）", "期末余额（原币）", "备注", ""])
    ws.merge_cells(start_row=r - 1, start_column=8, end_row=r - 1, end_column=9)
    if not res["news"]:
        c = ws.cell(r, 1, "无此类账户。"); c.font = txt_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9); r += 1
    else:
        for row in res["news"]:
            for i, v in enumerate([row["company"], row["acctNo"], row["subject"], row["bankName"], row["currency"]], start=1):
                c = ws.cell(r, i, v); c.font = txt_font; c.border = border; c.fill = miss_fill
            c = ws.cell(r, 6, row["local"]); c.number_format = num_fmt; c.border = border; c.fill = miss_fill; c.alignment = right
            c = ws.cell(r, 7, row["orig"]); c.number_format = num_fmt; c.border = border; c.fill = miss_fill; c.alignment = right
            note = "上年底稿中未找到该账号，请确认是否为新增账户" if row["acctNo"] else "未识别到银行账号，请检查余额表或手工补录"
            c = ws.cell(r, 8, note); c.font = txt_font; c.border = border; c.fill = miss_fill
            ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
            r += 1
        c = ws.cell(r, 1, "合计（%d 户）" % len(res["news"])); c.font = tot_font; c.border = border
        for i in range(2, 6):
            ws.cell(r, i).border = border
        c = ws.cell(r, 6, r2(sum(x["local"] for x in res["news"]))); c.number_format = num_fmt; c.border = border; c.alignment = right; c.font = tot_font
        c = ws.cell(r, 7, r2(sum(x["orig"] for x in res["news"]))); c.number_format = num_fmt; c.border = border; c.alignment = right; c.font = tot_font
        ws.cell(r, 8).border = border
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        r += 1

    # 四、上年有、本年无
    r += 1
    sect("四、上年BKD底稿有、本年科目余额表未找到的账户（需核实是否已销户或漏记）")
    head(["公司名称", "账号", "银行/存款机构名称", "科目名称", "币种", "上年余额（本币）", "上年余额（原币）", "备注", ""])
    ws.merge_cells(start_row=r - 1, start_column=8, end_row=r - 1, end_column=9)
    if not res["miss"]:
        c = ws.cell(r, 1, "无此类账户。"); c.font = txt_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9); r += 1
    else:
        for row in res["miss"]:
            for i, v in enumerate([row["company"], row.get("acctNo", ""), row["bank"], row["subject"], row["currency"]], start=1):
                c = ws.cell(r, i, v); c.font = txt_font; c.border = border
            c = ws.cell(r, 6, row["local"]); c.number_format = num_fmt; c.border = border; c.alignment = right
            c = ws.cell(r, 7, row["orig"]); c.number_format = num_fmt; c.border = border; c.alignment = right
            c = ws.cell(r, 8, "本年余额表未找到该账号：如已销户请确认余额已转出；如未销户请核实是否漏取数据")
            c.font = txt_font; c.border = border; c.alignment = left
            ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
            r += 1
        c = ws.cell(r, 1, "合计（%d 户）" % len(res["miss"])); c.font = tot_font; c.border = border
        for i in range(2, 6):
            ws.cell(r, i).border = border
        c = ws.cell(r, 6, r2(sum(x["local"] for x in res["miss"]))); c.number_format = num_fmt; c.border = border; c.alignment = right; c.font = tot_font
        c = ws.cell(r, 7, r2(sum(x["orig"] for x in res["miss"]))); c.number_format = num_fmt; c.border = border; c.alignment = right; c.font = tot_font
        ws.cell(r, 8).border = border
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        r += 1

    # 五、非精确匹配
    r += 1
    sect("五、非精确匹配（按账号或银行名称匹配，需人工复核）")
    head(["本年公司名称", "本年账号", "匹配方式", "上年底稿公司名称", "上年底稿账号", "币种", "本年余额（本币）", "上年余额（本币）", "差异"])
    ne = [x for x in res["nonexact"] if x["_matched"]]
    if not ne:
        c = ws.cell(r, 1, "无（全部账户均按「公司名称+账号」精确匹配）。"); c.font = txt_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9); r += 1
    else:
        for row in ne:
            vals = [row["company"], row["acctNo"], row["_how"], row["priorCompany"], row["priorAcct"], row["currency"]]
            for i, v in enumerate(vals, start=1):
                c = ws.cell(r, i, v); c.font = txt_font; c.border = border
                if i in (2, 3, 5, 6):
                    c.alignment = center
            c = ws.cell(r, 7, row["local"]); c.number_format = num_fmt; c.border = border; c.alignment = right
            c = ws.cell(r, 8, row["priorLocal"] if row["priorLocal"] is not None else ""); c.number_format = num_fmt; c.border = border; c.alignment = right
            c = ws.cell(r, 9, r2(row["local"] - row["priorLocal"]) if row["priorLocal"] is not None else ""); c.number_format = num_fmt; c.border = border; c.alignment = right
            r += 1

    # 六、大额变动
    r += 1
    sect("六、金额变动较大账户（变动率≥%d%%，且变动额≥%s 元）" % (round(rate_threshold * 100), format(amt_threshold, ",")))
    head(["公司名称", "账号", "银行/机构名称", "币种", "上年余额（本币）", "本年余额（本币）", "变动额", "变动率", "备注"])
    if not res["changes"]:
        c = ws.cell(r, 1, "无。"); c.font = txt_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9); r += 1
    else:
        for x in res["changes"]:
            row = x["r"]
            for i, v in enumerate([row["company"], row["acctNo"], row["bankName"], row["currency"]], start=1):
                c = ws.cell(r, i, v); c.font = txt_font; c.border = border
                if i in (2, 4):
                    c.alignment = center
            c = ws.cell(r, 5, row["priorLocal"]); c.number_format = num_fmt; c.border = border; c.alignment = right
            c = ws.cell(r, 6, row["local"]); c.number_format = num_fmt; c.border = border; c.alignment = right
            c = ws.cell(r, 7, x["delta"]); c.number_format = num_fmt; c.border = border; c.alignment = right
            c = ws.cell(r, 8, x["rate"]); c.number_format = "0.0%"; c.border = border; c.alignment = center
            c = ws.cell(r, 9, "余额发生较大变动，建议取得对账单/询证函支持"); c.font = txt_font; c.border = border
            r += 1

    # 七、Cutoff 截止性测试明细（完整清单；C.03 已按实际笔数自动插行，全部写入）
    if cutoff and (cutoff.get("before") or cutoff.get("after")):
        cs = cutoff.get("stat") or {}
        r += 1
        sect("七、C.03 Cutoff 截止性测试明细（阈值 %s 元　期间 %s 至 %s　共 %d 笔）" % (
            format(cs.get("threshold", 0), ",.2f"),
            cs.get("windowStart", ""), cs.get("windowEnd", ""),
            cs.get("beforeCount", 0) + cs.get("afterCount", 0)))
        head(["编号", "公司名称", "付款银行 / 账号", "收款银行 / 账号", "付款日期", "收款日期",
              "金额（本币）", "跨期", "备注"])
        ordered = (cutoff.get("before") or []) + (cutoff.get("after") or [])
        if not ordered:
            c = ws.cell(r, 1, "期间内未识别到超过阈值的银行间转账。"); c.font = txt_font
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
            r += 1
        else:
            for s in ordered:
                vals = [
                    s.get("no", ""), s.get("company", ""),
                    "%s / %s" % (s.get("payBank", ""), s.get("payAcct", "") or "未识别"),
                    "%s / %s" % (s.get("recvBank", ""), s.get("recvAcct", "") or "未识别"),
                    s.get("payDate", ""), s.get("recvDate", ""),
                ]
                for i, v in enumerate(vals, start=1):
                    c = ws.cell(r, i, v); c.font = txt_font; c.border = border
                    if i in (1, 5, 6):
                        c.alignment = center
                    else:
                        c.alignment = left
                c = ws.cell(r, 7, s.get("amt")); c.number_format = num_fmt
                c.border = border; c.alignment = right; c.font = txt_font
                c = ws.cell(r, 8, s.get("cross", "")); c.font = txt_font; c.border = border; c.alignment = center
                if s.get("cross") == "是":
                    c.fill = miss_fill
                c = ws.cell(r, 9, s.get("note", "")); c.font = txt_font; c.border = border; c.alignment = left
                r += 1
            c = ws.cell(r, 1, "合计（%d 笔，跨期 %d 笔）" % (
                len(ordered), cs.get("crossCount", 0)))
            c.font = tot_font; c.border = border
            for i in range(2, 7):
                ws.cell(r, i).border = border
            c = ws.cell(r, 7, r2(sum(x.get("amt", 0) for x in ordered)))
            c.number_format = num_fmt; c.border = border; c.alignment = right; c.font = tot_font
            ws.cell(r, 8).border = border; ws.cell(r, 9).border = border
            r += 1
        if cs.get("belowThreshold"):
            c = ws.cell(r, 1, "另有 %d 笔转账低于测试阈值 %s 元，未纳入测试范围。" % (
                cs["belowThreshold"], format(cs.get("threshold", 0), ",.2f")))
            c.font = sub_font
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
            r += 1
        if cs.get("singleCount"):
            c = ws.cell(r, 1, "另有 %d 笔单边记录（只转出未见转入、或只转入未见转出）未配对，请复核是否漏记到账。明细见「C.03 Cutoff」测试说明。" % cs["singleCount"])
            c.font = sub_font
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
            r += 1

    r += 1
    c = ws.cell(r, 1, "说明：本表由「货币资金BKD生成器」自动生成。BKD表中的「期末账面金额」取自PTW平台清洗后的科目余额表；其余字段沿自上年BKD底稿。请对第三、四、五、六部分逐项复核。")
    c.font = sub_font
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    ws.row_dimensions[r].height = 44


# ============================================================ 账面信息表（银行存款账面模板）
# 从 tb 取银行存款账户，按「主体 × 末级科目 × 币别」填进官方「1.银行存款-账面模板.xlsx」。
#
# 为什么不用 openpyxl 另存：
#   该模板的下拉菜单是 x14 扩展（extLst/x14:dataValidation），另存后 openpyxl 会丢掉
#   x14 数据验证、xl/woinfos.xml、customXml 等 8 个包内部件（实测 19 -> 11），
#   模板就废了。所以这里走 zip 级手术：只重写 sheet1.xml 的 sheetData，
#   其余部件原样复制；文本用 inlineStr，不碰 sharedStrings。
BOOK_TPL_NAME = "1.银行存款-账面模板.xlsx"
BOOK_TPL_LEGACY = os.path.join(APP_DIR, "_tpl", BOOK_TPL_NAME)
BOOK_HEADER_ROW = 2                 # 表头在第 2 行，数据从第 3 行起
BOOK_FIRST_DATA_ROW = 3
# B~K 十列；值对应「账面信息」表头
BOOK_COLS = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
BOOK_DEFAULT_STYLES = {"B": "18", "C": "19", "D": "19", "E": "19", "F": "20",
                       "G": "19", "H": "20", "I": "20", "J": "21", "K": "21"}


def _add_gap_section(ws, res):
    """在检查表末尾追加「待确认清单」——所有标黄的格子逐行列出。

    这张清单是「上数工具」对使用者的承诺：凡是工具没填的，都在这里交代清楚，
    人不用去底稿里逐格找黄色。
    """
    gaps = res.get("gaps") or []
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    txt_font = Font(name="宋体", size=10)
    tot_font = Font(name="宋体", size=10, bold=True)
    sect_font = Font(name="宋体", size=10, bold=True, color="1F3864")
    sect_fill = PatternFill("solid", fgColor="F2F2F2")
    gap_fill = PatternFill("solid", fgColor="FFFF00")
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    r = ws.max_row + 2
    c = ws.cell(r, 1, "五、待确认清单（底稿中标黄的格子，共 %d 处）" % len(gaps))
    c.font = sect_font
    c.fill = sect_fill
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    r += 1

    if not gaps:
        c = ws.cell(r, 1, "无。本次所有需要的资料都已取到。")
        c.font = txt_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        return

    # 按原因分组，先给一张汇总，再给明细
    by_reason = {}
    for g in gaps:
        by_reason[g["reason"]] = by_reason.get(g["reason"], 0) + 1
    for reason, n in sorted(by_reason.items(), key=lambda x: -x[1]):
        c = ws.cell(r, 1, "· %s —— %d 处" % (reason, n))
        c.font = txt_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        r += 1

    r += 1
    heads = ["工作表", "单元格", "字段", "公司名称", "科目名称", "账号", "原因", "", ""]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(r, i, h)
        c.font = tot_font
        c.border = border
        c.fill = PatternFill("solid", fgColor="D9E2F3")
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
    r += 1

    for g in gaps:
        cell_ref = "%s%d" % (get_column_letter(g["col"]), g["row"])
        vals = [g["sheet"], cell_ref, g["field"], g["company"], g["subject"], g["acct"], g["reason"]]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(r, i, v)
            c.font = txt_font
            c.border = border
            c.alignment = left
            if i == 2:
                c.fill = gap_fill
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
        r += 1


def probe_prior_structure(wb, max_sheets=20):
    """给「上年 WP」做结构探测：每张表找最可能的表头行并打分，输出可交给 AI 的 markdown。

    为什么要这一步：上年 WP 是客户/去年的产出，sheet 名和列位每个项目都不一样，
    正则硬猜会静默错位。这里把「表长什么样」摊开给 AI 看，由 AI 判断哪张是明细表、
    表头在第几行、各列是什么，比让代码猜可靠得多。
    """
    sheets = []
    for name in wb.sheetnames[:max_sheets]:
        ws = wb[name]
        scan = min(ws.max_row or 1, 40)
        ncol = min(ws.max_column or 1, 25)
        rows = [[ws.cell(r, c).value for c in range(1, ncol + 1)] for r in range(1, scan + 1)]
        try:
            best = detect_header(rows, BKD_SPEC, limit=scan)
        except Exception:
            best = {"i": 0, "score": 0}
        sheets.append({
            "name": name, "rows": ws.max_row or 0, "cols": ws.max_column or 0,
            "score": best.get("score", 0), "header_row": (best.get("i", 0) + 1),
            "preview": [[None if v is None else str(v)[:60] for v in row] for row in rows[:12]],
        })
    sheets.sort(key=lambda x: -x["score"])

    md = ["# 上年 WP 结构探测包", "",
          "> 由 PTW 平台自动生成。用途：让 AI 认出「哪张表是 BKD 明细、表头在第几行」，",
          "> 因为上年底稿的 sheet 名与列位每个项目都不一样，代码硬猜会静默错位。", "",
          "| # | 工作表 | 行数 | 列数 | 表头匹配分 | 推测表头行 |", "|---|---|---|---|---|---|"]
    for i, s in enumerate(sheets, 1):
        md.append("| %d | %s | %s | %s | %s | %s |" % (
            i, s["name"], s["rows"], s["cols"], s["score"], s["header_row"]))
    md += ["",
           "匹配分满分 21（公司 3 / 账号 3 / 科目 2 / 银行 2 / 原币 2 / 本币 2 / 期末金额 2 / 币种 1 / 性质 1 / 函证 1 / 理由 1 / Notes 1）。",
           "≥16 基本就是 BKD 明细表；<8 通常是说明页或参数页。注意分数高不等于就是 BKD——",
           "C.01 Confirmations 这类表也有账号和金额列，最终靠「有没有逐账户明细」判断。", ""]
    for s in sheets[:4]:
        md += ["## %s（推测表头第 %d 行，分 %d）" % (s["name"], s["header_row"], s["score"]), "```"]
        for r, row in enumerate(s["preview"], 1):
            cells = ["" if v is None else str(v).replace("\n", " ") for v in row]
            while cells and cells[-1] == "":
                cells.pop()
            if cells:
                md.append("r%-3d| %s" % (r, " | ".join(cells)))
        md += ["```", ""]
    md += ["## 请按这个格式回一份 JSON", "```json", "{",
           '  "sheet": "C.00 BKD",',
           '  "header_row": 9,',
           '  "data_start": 10,',
           '  "data_end": 135,',
           '  "note": "为什么判断它是明细表（一句话）"',
           "}", "```"]
    return {"sheets": sheets, "markdown": "\n".join(md),
            "best": sheets[0] if sheets else None}


def book_tpl_path():
    """账面模板路径：优先 data/bkd_tpl，其次模块内置 _tpl（自动复制过去）。"""
    p = os.path.join(TPL_DIR, BOOK_TPL_NAME)
    if os.path.exists(p):
        return p
    if os.path.exists(BOOK_TPL_LEGACY):
        try:
            import shutil
            os.makedirs(TPL_DIR, exist_ok=True)
            shutil.copy2(BOOK_TPL_LEGACY, p)
            return p
        except Exception:
            return ""
    return ""


def _split_bank_acct(v):
    """把「工商银行:6222xxx」这类辅助核算值拆成 (银行名, 账号)。"""
    t = T(v)
    if not t:
        return "", ""
    m = re.match(r"^([^:：]{2,24})[:：]\s*(.+)$", t)
    if m and pick_acct(m.group(2)):
        return m.group(1).strip(), pick_acct(m.group(2))
    return "", pick_acct(t)


def _book_currency(v):
    """币种归一：'美元(USD)' / 'USD' / '美元' 一律给「美元」，保证同表口径一致。"""
    cn = {"CNY": "人民币", "RMB": "人民币", "USD": "美元", "HKD": "港币", "EUR": "欧元",
          "JPY": "日元", "GBP": "英镑", "AUD": "澳元", "CAD": "加拿大元", "SGD": "新加坡元",
          "CHF": "瑞士法郎", "TWD": "新台币", "KRW": "韩元", "THB": "泰铢", "VND": "越南盾",
          "MYR": "马来西亚林吉特", "NZD": "新西兰元", "RUB": "卢布", "INR": "印度卢比"}
    t = T(v)
    if not t:
        return ""
    m = re.match(r"^(.+?)[（(]([A-Za-z]{2,4})[)）]$", t)
    if m:
        return cn.get(m.group(2).upper()) or norm_cur(m.group(1)) or t
    return norm_cur(t) or t


def _bank_like(v):
    """判断一个值像不像银行/开户行名称（用于从辅助核算里挑银行名）。"""
    t = T(v)
    if not t or len(t) > 30:
        return ""
    if re.search(r"银行|农商|信用社|邮政|支行|分行|营业部|信用社|储蓄所|信用合作", t):
        return t
    return ""


def _pick_acct_any(*vals):
    """账号取值：先按数字规则提，提不到就用原文（很多系统账号是「账号010」这种）。"""
    for v in vals:
        t = T(v)
        if not t:
            continue
        a = pick_acct(t)
        if a:
            return a
        # 去掉「账号/账户/卡号」这类前缀后，剩下的整体当账号
        s = re.sub(r"^(银行账号|银行帐号|账号|帐号|账户|帐户|卡号)[:：\s]*", "", t).strip()
        if s and len(s) <= 40:
            return s
    return ""


def read_book_rows(conn=None, keyword="银行存款"):
    """取银行存款账户，返回 (rows, meta)。

    行粒度 = 主体 × 科目 × 币别。余额表里的「综合本位币」是各外币的折算合计，
    单独出会和外币行重复，所以排除；只有当某个科目只有综合本位币行时才用它兜底。
    """
    close_after = False
    if conn is None:
        if not os.path.exists(KZ_DB):
            return [], {"error": "找不到PTW平台数据库：%s" % KZ_DB}
        try:
            con = _open_db_readonly(KZ_DB)
            close_after = True
        except RuntimeError as e:
            return [], {"error": str(e)}
    else:
        con = conn
    try:
        tabs = [r[0] for r in con.execute("SHOW TABLES").fetchall()]
        if "tb" not in tabs:
            return [], {"error": "PTW平台里还没有「科目余额表」（tb 表），请先导入并清洗科目余额表。"}
        kw = T(keyword).replace(" ", "")
        want = ("1002", "银行存款") if kw in ("银行存款", "1002") else (kw,)
        in_list = ",".join("'%s'" % x.replace("'", "''") for x in want)
        sql = """
        SELECT coalesce(nullif(trim(company_name),''), nullif(trim(book_name),''), '(未标注主体)') AS company,
               coalesce(nullif(trim(acct_code),''), '') AS code,
               coalesce(nullif(trim(acct_name),''), '') AS name,
               coalesce(nullif(trim(currency),''), '')   AS cur,
               coalesce(nullif(trim(bank_account), ''), '') AS bank_acct,
               sum(end_bal) AS end_bal, sum(end_debit) AS end_dr, sum(end_credit) AS end_cr,
               sum(end_bal_fx) AS end_bal_fx, sum(end_debit_fx) AS end_dr_fx,
               sum(end_credit_fx) AS end_cr_fx,
               max(nullif(trim(aux_text), '')) AS aux_text,
               max(nullif(trim(aux1), '')) AS aux1,
               max(nullif(trim(aux2), '')) AS aux2,
               max(nullif(trim(aux3), '')) AS aux3
        FROM tb
        WHERE nullif(trim(acct_l1), '') IN ({L})
           OR (nullif(trim(acct_l1), '') IS NULL
               AND (trim(coalesce(acct_code,'')) IN ({L})
                    OR trim(coalesce(acct_name,'')) IN ({L})))
        GROUP BY 1, 2, 3, 4, 5
        ORDER BY 1, 2, 3, 4, 5
        """.format(L=in_list)
        raw = con.execute(sql).fetchall()
    finally:
        if close_after:
            con.close()

    # 先按 (主体, 科目) 归组，组内区分「综合本位币」与真实币别
    groups = {}
    for (company, code, name, cur, bank_acct, end_bal, end_dr, end_cr,
         bal_fx, dr_fx, cr_fx, aux_text, aux1, aux2, aux3) in raw:
        # 分组键必须带账号：同一家公司同一个「银行存款」科目下常挂多个账号，
        # 只按公司+科目归组会把不同账号的币别行混成一行。
        company = re.sub(r"[(（]\d{2,6}[)）]$", "", T(company)).strip() or T(company)
        key = (company, T(code), T(name), T(bank_acct))
        g = groups.setdefault(key, {"fx": [], "sum_row": None})
        item = {"cur": T(cur), "end_bal": end_bal, "end_dr": end_dr, "end_cr": end_cr,
                "bal_fx": bal_fx, "dr_fx": dr_fx, "cr_fx": cr_fx, "bank_acct": T(bank_acct),
                "aux1": T(aux1), "aux2": T(aux2), "aux3": T(aux3), "aux_text": T(aux_text)}
        if "综合本位币" in item["cur"]:
            g["sum_row"] = item
        else:
            g["fx"].append(item)

    rows, no_bank, no_acct = [], 0, 0
    for (company, code, name, _acct_key), g in groups.items():
        picked = g["fx"] or ([dict(g["sum_row"], cur="人民币(CNY)")] if g["sum_row"] else [])
        for it in picked:
            bal = it["end_bal"]
            local = float(bal) if bal is not None else (
                float(it["end_dr"] or 0) - float(it["end_cr"] or 0))
            fxv = it["bal_fx"]
            orig = float(fxv) if fxv is not None else (
                float(it["dr_fx"] or 0) - float(it["cr_fx"] or 0))
            if not any(x is not None for x in (it["bal_fx"], it["dr_fx"], it["cr_fx"])):
                orig = local                      # 没有外币列时，原币＝本币
            bank, acct_aux = _split_bank_acct(it["bank_acct"])
            if not bank:
                bank = derive_bank(name)
            if not bank:                       # 辅助核算里常带开户行（「工商银行」等）
                for v in (it["aux1"], it["aux2"], it["aux3"], it["aux_text"]):
                    b = _bank_like(v)
                    if b:
                        bank = b
                        break
            # 账号：bank_account 原文优先（脱敏数据可能是「账号010」这种），
            # 其次从辅助核算/科目名称里按数字规则提
            acct = (_pick_acct_any(it["bank_acct"]) or acct_aux
                    or pick_acct(it["aux_text"]) or pick_acct(name))
            if not bank:
                no_bank += 1
            if not acct:
                no_acct += 1
            rows.append({
                "company": company, "subject": name or code, "bank": bank,
                "acctNo": acct, "currency": _book_currency(it["cur"]) or it["cur"],
                "orig": r2(orig), "local": r2(local),
                "confirm": "是",                  # 按约定：全部默认「是」
                "note": "",                       # 判断理由留空给人工
                "_missBank": not bank, "_missAcct": not acct,
            })
    meta = {"rows": len(rows), "groups": len(groups), "missBank": no_bank,
            "missAcct": no_acct, "keyword": kw,
            "balanced": round(sum(r["local"] for r in rows), 2)}
    return rows, meta


def _xesc(s):
    return (str(s if s is not None else "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _cell_xml(ref, style, value, kind):
    """按类型生成一个 <c>。kind: empty/text/num/date"""
    s = ' s="%s"' % style if style else ""
    if kind == "empty" or value is None or value == "":
        return '<c r="%s"%s/>' % (ref, s)
    if kind == "text":
        return ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                % (ref, s, _xesc(value)))
    if kind == "date":
        return '<c r="%s"%s><v>%d</v></c>' % (ref, s, value)
    return '<c r="%s"%s><v>%s</v></c>' % (ref, s, value)


def _excel_serial(d):
    return (d - _date(1899, 12, 30)).days


def build_book_xlsx(rows, balance_date=None, tpl_path=None):
    """把 rows 填进账面模板，返回 (bytes, 模板路径)。保留模板全部部件与下拉。"""
    import zipfile
    tp = tpl_path or book_tpl_path()
    if not tp:
        return None, ""
    bd = _as_date(balance_date) if balance_date else None

    with zipfile.ZipFile(tp) as zin:
        sheet_name = None
        for n in zin.namelist():
            if re.match(r"^xl/worksheets/sheet1\.xml$", n):
                sheet_name = n
                break
        if not sheet_name:
            return None, tp
        xml = zin.read(sheet_name).decode("utf-8")
        others = [(i, zin.read(i.filename)) for i in zin.infolist() if i.filename != sheet_name]

    m = re.search(r"<sheetData>(.*)</sheetData>", xml, re.S)
    if not m:
        return None, tp
    body = m.group(1)

    # 取模板第 3 行的行属性与每列样式，逐行克隆（模板改版也不怕）
    row3 = re.search(r'<row r="3"([^>]*)>(.*?)</row>', body, re.S)
    styles = dict(BOOK_DEFAULT_STYLES)
    row_attrs = ' spans="2:11" s="8" customFormat="1" ht="47" customHeight="1"'
    if row3:
        row_attrs = row3.group(1)
        for col, st in re.findall(r'<c r="([A-Z]+)3"[^>]*\bs="(\d+)"', row3.group(2)):
            styles[col] = st

    keep_rows = re.findall(r'<row r="([12])"[^>]*>.*?</row>', body, re.S)
    kept = [r for r in re.findall(r'<row r="[12]"[^>]*>.*?</row>', body, re.S)]

    out = list(kept)
    for i, rec in enumerate(rows):
        rn = BOOK_FIRST_DATA_ROW + i
        vals = {
            "B": (_excel_serial(bd) if bd else None, "date"),
            "C": (rec.get("company") or "", "text"),
            "D": (rec.get("subject") or "", "text"),
            "E": (rec.get("bank") or "", "text"),
            "F": (rec.get("acctNo") or "", "text"),
            "G": (rec.get("currency") or "", "text"),
            "H": (("%.2f" % rec["orig"]) if rec.get("orig") is not None else None, "num"),
            "I": (("%.2f" % rec["local"]) if rec.get("local") is not None else None, "num"),
            "J": (rec.get("confirm") or "", "text"),
            "K": (rec.get("note") or "", "text"),
        }
        cells = "".join(_cell_xml("%s%d" % (c, rn), styles.get(c), vals[c][0], vals[c][1])
                        for c in BOOK_COLS)
        out.append('<row r="%d"%s>%s</row>' % (rn, row_attrs, cells))

    new_xml = xml[:m.start()] + "<sheetData>" + "".join(out) + "</sheetData>" + xml[m.end():]
    last = BOOK_FIRST_DATA_ROW + max(len(rows) - 1, 0)
    new_xml = re.sub(r'<dimension ref="([A-Z]+1):([A-Z]+)\d+"/>',
                     lambda mm: '<dimension ref="%s:%s%d"/>' % (mm.group(1), mm.group(2), last),
                     new_xml, count=1)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in others:
            zout.writestr(info, data)
        zout.writestr(sheet_name, new_xml.encode("utf-8"))
    return buf.getvalue(), tp
