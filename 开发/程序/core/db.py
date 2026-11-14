# -*- coding: utf-8 -*-
"""DuckDB 数据层：导入 / 表头识别 / 字段映射 / 合并 / 查询 / 导出 / 分享包"""
import csv
import difflib
import hashlib
import io
import json
import os
import re
import shutil
import sys
import threading
import time
import uuid
import zipfile
from datetime import date, datetime
from decimal import Decimal

import duckdb

from . import clean, htmlread, recon as recon_mod, recipe, xlsread, xlsxinfo
from .fields import (AUX_ROLES, RECON_LABEL, ROLE_KEYS, TARGETS, auto_map,
                     cn_headers, detect_header, field_dict, role_from_header,
                     role_from_values, to_cn)

MACROS = r"""
CREATE OR REPLACE MACRO _num(s) AS
  CASE
    WHEN s IS NULL THEN NULL
    WHEN trim(s) IN ('', '-', '--', '/', '—', '无', 'NULL') THEN NULL
    -- 文字里夹数字（「合计123」「见附注123」）不能当金额用：直接吞成 123 会让
    -- 一个说明行悄悄参与汇总。只放行「纯数字串 + 计量单位后缀」这种正常写法。
    WHEN regexp_matches(trim(s), '[^\d\s,，.．\-+()¥￥$%Ee－—−　]')
         AND NOT regexp_full_match(trim(s),
             '[\d\s,，.．\-+()¥￥$%Ee－—−　]*(元|万元|万|美元|港币|欧元|日元|%|CR|DR|cr|dr)')
      THEN NULL
    -- 欧式千分位 1.234,56（点作千分位、逗号作小数点）：德法系统导出
    WHEN regexp_full_match(trim(s), '[-+]?\d{1,3}(\.\d{3})+,\d+') THEN
      try_cast(replace(replace(trim(s), '.', ''), ',', '.') AS DECIMAL(20,4))
    WHEN regexp_full_match(trim(s), '\(.*\)') THEN
      -1 * try_cast(NULLIF(regexp_replace(translate(trim(s), '－—−　', '-  '), '[^0-9.]', '', 'g'), '') AS DECIMAL(20,4))
    ELSE COALESCE(
      try_cast(NULLIF(regexp_replace(translate(trim(s), '－—−　', '-  '), '[ ,，¥￥$]', '', 'g'), '') AS DECIMAL(20,4)),
      try_cast(NULLIF(regexp_replace(translate(trim(s), '－—−　', '-  '), '[^0-9.\-]', '', 'g'), '') AS DECIMAL(20,4))
    )
  END;

CREATE OR REPLACE MACRO _to_date(s) AS
  CASE
    WHEN s IS NULL OR trim(s) = '' THEN NULL
    WHEN regexp_full_match(trim(s), '\d{5}') THEN
      CASE WHEN try_cast(trim(s) AS INTEGER) BETWEEN 20000 AND 80000
           THEN (DATE '1899-12-30' + try_cast(trim(s) AS INTEGER)) ELSE NULL END
    WHEN regexp_full_match(trim(s), '\d{8}') THEN try_strptime(trim(s), '%Y%m%d')::DATE
    WHEN regexp_full_match(trim(s), '\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?([ T].*)?') THEN
      try_strptime(regexp_replace(regexp_replace(regexp_replace(substr(trim(s), 1, 10), '[年月./]', '-', 'g'), '日', ''), '-+', '-', 'g'), '%Y-%m-%d')::DATE
    WHEN regexp_full_match(trim(s), '\d{4}[-/.年]\d{1,2}月?') THEN
      try_strptime(regexp_replace(regexp_replace(substr(trim(s), 1, 8), '[年月/.]', '-', 'g'), '-+', '-', 'g') || '-01', '%Y-%m-%d')::DATE
    ELSE NULL
  END;

CREATE OR REPLACE MACRO _text(s) AS NULLIF(trim(s), '');

-- 辅助核算取值：值形如「客户:阿里巴巴/部门:销售部」时按类型前缀取名字。
-- plain=1 时，没有冒号的整列值原样取（该列已被明确指定角色）；
-- plain=0 时只认带类型前缀的片段（用于「混合」列，避免整列值被塞进所有语义列）。
CREATE OR REPLACE MACRO _aux_pick(s, keys, plain) AS
  CASE
    WHEN s IS NULL OR trim(s) = '' THEN NULL
    WHEN regexp_matches(trim(s), '[:：]') THEN
      NULLIF(trim(regexp_extract(trim(s), '(' || keys || ')[:：]\s*([^/、,，；|]+)', 2)), '')
    WHEN plain = 1 THEN trim(s)
    ELSE NULL
  END;

CREATE OR REPLACE MACRO _period_norm(s) AS
  CASE
    WHEN s IS NULL OR trim(s) = '' THEN NULL
    -- 年份限定 19xx/20xx/21xx：科目编码「1121.01.01」长得像日期，
    -- 不限年份就会被当成 1121 年 1 月。会计期间的年份一定在这个区间内。
    -- 完整日期 → 归到月。财务系统导出「会计期间」时被 Excel 转成日期是常态，
    -- 不归一的话 '2026-10-31' 和 '2026-10' 会被当成两个不同期间。
    WHEN regexp_full_match(trim(s),
         '(19|20|21)\d{2}[-/.年]\d{1,2}[-/.月]\d{1,2}日?([ T].*)?') THEN
      left(trim(s), 4) || '-' ||
      lpad(regexp_extract(trim(s), '[-/.年](\d{1,2})', 1), 2, '0')
    -- 八位纯数字：20261031。月和日必须合法，否则 10537050 这种流水号
    -- 会被当成「1053 年 70 月」，比不归一更糟。
    WHEN regexp_full_match(trim(s),
         '(19|20|21)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])') THEN
      left(trim(s), 4) || '-' || substr(trim(s), 5, 2)
    WHEN regexp_full_match(trim(s), '(19|20|21)\d{2}[-年./]\d{1,2}(月|期)?') THEN
      left(trim(s), 4) || '-' || lpad(regexp_extract(trim(s), '[-年./](\d{1,2})', 1), 2, '0')
    WHEN regexp_full_match(trim(s), '\d{6}') THEN left(trim(s), 4) || '-' || right(trim(s), 2)
    WHEN regexp_full_match(trim(s), '\d{1,2}') THEN lpad(trim(s), 2, '0')
    ELSE trim(s)
  END;

CREATE OR REPLACE MACRO _book_norm(s) AS
  CASE
    WHEN s IS NULL OR trim(s) = '' THEN NULL
    ELSE NULLIF(trim(regexp_replace(trim(s), '\s*\([0-9]+\)\s*$', '')), '')
  END;
"""

# 文件名当账簿名称时，尾巴上那截「_科目余额表」「-序时账」不是账簿名的一部分，
# 留着反而让「账簿对照表」匹配不上（对照表里写的是公司名）。
# 只剥"表类型"词，且剥完得有像样的内容（不能只剩个「2026年」）。
_LEDGER_TAIL = ("固定资产台账", "多账簿科目余额表", "科目余额表", "银行存款余额表",
                "余额表", "序时账", "明细账", "日记账", "总账", "凭证", "账簿",
                "银行流水", "流水", "台账")


# ---------------------------------------------------------------- 主体名归一
# 账簿名称和对照表里的公司名称经常是「同一家、不同写法」：
#   （南海店）佛山固生堂中医门诊部有限公司(21101)
#   佛山固生堂中医门诊部有限公司
# 平台只做「去掉尾部 (数字)」的归一（_book_norm），剩下的要靠推断 + 留痕。
_BOOK_SUFFIX = ("股份有限公司", "有限责任公司", "有限公司", "个体工商户", "个体户")


def _book_key(s, strip_prefix=False):
    """归一后的比对键。

    strip_prefix=False（保守）：只去尾部 (数字)、空格、公司类型后缀。
    strip_prefix=True （激进）：再去掉开头「（南海店）」这类门店前缀。
    **两套分开用**——门店前缀是真信息，「（南海店）」和「（东湖店）」是两家，
    激进归一只能当兜底，命中后必须标「需人工确认」。
    """
    t = str(s or "").replace("　", " ").strip()
    if not t:
        return ""
    t = re.sub(r"\s*[（(]\d{3,8}[）)]\s*$", "", t)          # 尾部 (21101)
    if strip_prefix:
        t = re.sub(r"^[（(][^）)]{1,10}[）)]\s*", "", t)     # 开头 （南海店）
    t = re.sub(r"\s+", "", t)
    for suf in _BOOK_SUFFIX:
        if t.endswith(suf) and len(t) >= len(suf) + 3:
            t = t[:-len(suf)]
            break
    return t


def _common_prefix(a, b):
    n = 0
    for x, y in zip(a, b):
        if x != y:
            break
        n += 1
    return n


def _pick_book(name, cands):
    """给定一个账簿名和候选池，判断它最可能是哪家公司。

    cands: [(保守键, 激进键, 公司名, 对照表原账簿名), ...]
    返回  (score, method, company, evidence, need_confirm)
    method 由强到弱：norm / norm_prefix / contain / prefix / similar / none。
    只有 norm 能自动认定；其余都要人看一眼（门店前缀和近似名最容易错配）。
    """
    k1, k2 = _book_key(name), _book_key(name, True)
    if not k1:
        return (0.0, "none", None, "账簿名为空或全被归一掉了", 1)
    # 1) 归一后完全相同
    for kl, kt, cn, bn in cands:
        if kl and k1 == kl:
            return (1.0, "norm", cn,
                    "去掉尾码/空格/公司类型后缀后，与对照表账簿「%s」完全相同" % bn, 0)
    # 2) 去掉门店前缀后相同
    for kl, kt, cn, bn in cands:
        if kt and k2 == kt:
            m = re.match(r"^[（(]([^）)]{1,10})[）)]", str(name).strip())
            return (0.95, "norm_prefix", cn,
                    "去掉前缀「%s」后与「%s」相同——门店不同，需确认是否同一主体"
                    % (m.group(1) if m else "—", bn), 1)
    best = None
    # 3) 互相包含
    for kl, kt, cn, bn in cands:
        if not kl or len(k1) < 4:
            continue
        if k1 in kl or kl in k1:
            sc = min(len(k1), len(kl)) * 1.0 / max(len(k1), len(kl))
            sc = min(sc, 0.94)
            if not best or sc > best[0]:
                best = (sc, "contain", cn,
                        "与对照表账簿「%s」互相包含（短的一方完全出现在长的一方里）" % bn,
                        1 if sc < 0.85 else 0)
    if best:
        return best
    # 4) 共同开头
    for kl, kt, cn, bn in cands:
        cp = _common_prefix(k1, kl) if kl else 0
        if cp >= 4:
            sc = min(cp * 1.0 / max(len(k1), len(kl)) + 0.2, 0.9)
            if not best or sc > best[0]:
                best = (sc, "prefix", cn,
                        "与「%s」共享开头 %d 个字「%s」" % (bn, cp, k1[:cp]), 1)
    # 5) 文字相似度兜底
    for kl, kt, cn, bn in cands:
        if not kl:
            continue
        r = difflib.SequenceMatcher(None, k1, kl).ratio()
        if r >= 0.6 and (not best or r > best[0]):
            best = (r, "similar", cn, "与「%s」文字相似度 %.0f%%" % (bn, r * 100), 1)
    return best or (0.0, "none", None, "对照表里没有相近的主体名，需人工指定", 1)



def _strip_ledger_tail(name):
    s = str(name or "").strip()
    if not s:
        return s
    for w in _LEDGER_TAIL:
        if not s.endswith(w):
            continue
        head = s[:-len(w)].rstrip(" _-—()（）·.")
        # 剥完至少留 4 个字、且不能只剩年份/数字，否则宁可不剥。
        # 命中一个尾缀就定论：不能因为这次不剥，又去拿更短的词再剥一遍
        # （「2026年科目余额表」曾被剥成「2026年科目」）。
        if len(head) >= 4 and not re.fullmatch(r"[\d\s年月份\-_.]*", head):
            return head
        return s
    return s

CANON = {
    "journal": [
        ("src", "VARCHAR"), ("src_row", "BIGINT"), ("vdate", "DATE"),
        ("period", "VARCHAR"), ("vtype", "VARCHAR"), ("vno", "VARCHAR"),
        ("voucher_key", "VARCHAR"),
        ("summary", "VARCHAR"), ("acct_code", "VARCHAR"), ("acct_name", "VARCHAR"),
        ("acct_full", "VARCHAR"),
        ("acct_l1", "VARCHAR"),
        ("debit", "DECIMAL(20,2)"), ("credit", "DECIMAL(20,2)"),
        ("debit_fx", "DECIMAL(20,2)"), ("credit_fx", "DECIMAL(20,2)"),
        ("direction", "VARCHAR"), ("qty", "DECIMAL(20,4)"), ("currency", "VARCHAR"),
        ("book_name", "VARCHAR"), ("company_name", "VARCHAR"),
        ("short_name", "VARCHAR"),
        ("dept", "VARCHAR"), ("counterparty", "VARCHAR"), ("person", "VARCHAR"),
        ("checker", "VARCHAR"), ("poster", "VARCHAR"), ("attach", "VARCHAR"),
        # 辅助核算：与科目余额表一致，3 个位置槽 + 语义列 + 拼接检索列
        ("aux1", "VARCHAR"), ("aux2", "VARCHAR"), ("aux3", "VARCHAR"),
        ("bank_name", "VARCHAR"), ("bank_account", "VARCHAR"),
        ("counterparty_name", "VARCHAR"),
        ("dept_name", "VARCHAR"), ("project_name", "VARCHAR"),
        ("person_name", "VARCHAR"), ("item_name", "VARCHAR"),
        ("aux_text", "VARCHAR"),
    ],
    "tb": [
        ("src", "VARCHAR"), ("src_row", "BIGINT"),
        ("company_name", "VARCHAR"), ("book_name", "VARCHAR"),
        ("short_name", "VARCHAR"),
        ("acct_code", "VARCHAR"), ("acct_name", "VARCHAR"),
        ("acct_full", "VARCHAR"),
        ("acct_l1", "VARCHAR"),
        ("level", "VARCHAR"), ("currency", "VARCHAR"), ("direction", "VARCHAR"),
        ("open_debit", "DECIMAL(20,2)"), ("open_credit", "DECIMAL(20,2)"),
        ("open_bal", "DECIMAL(20,2)"),
        ("cur_debit", "DECIMAL(20,2)"), ("cur_credit", "DECIMAL(20,2)"),
        ("ytd_debit", "DECIMAL(20,2)"), ("ytd_credit", "DECIMAL(20,2)"),
        ("end_debit", "DECIMAL(20,2)"), ("end_credit", "DECIMAL(20,2)"),
        ("end_bal", "DECIMAL(20,2)"),
        # 外币（原币）金额：与本币金额一一对应，选填
        ("open_debit_fx", "DECIMAL(20,2)"), ("open_credit_fx", "DECIMAL(20,2)"),
        ("open_bal_fx", "DECIMAL(20,2)"),
        ("cur_debit_fx", "DECIMAL(20,2)"), ("cur_credit_fx", "DECIMAL(20,2)"),
        ("ytd_debit_fx", "DECIMAL(20,2)"), ("ytd_credit_fx", "DECIMAL(20,2)"),
        ("end_debit_fx", "DECIMAL(20,2)"), ("end_credit_fx", "DECIMAL(20,2)"),
        ("end_bal_fx", "DECIMAL(20,2)"),
        # 辅助核算：3 个位置槽 + 7 个语义列 + 1 个拼接检索列
        # bank_name 是「开户行」，排在账号前面：货币资金 BKD 的银行列优先吃它
        ("aux1", "VARCHAR"), ("aux2", "VARCHAR"), ("aux3", "VARCHAR"),
        ("bank_name", "VARCHAR"), ("bank_account", "VARCHAR"),
        ("counterparty_name", "VARCHAR"),
        ("dept_name", "VARCHAR"), ("project_name", "VARCHAR"),
        ("person_name", "VARCHAR"), ("item_name", "VARCHAR"),
        ("aux_text", "VARCHAR"),
    ],
    # ---------------------------------------------------------------- 银行流水
    # 前两段是原样承接银行导出；第三段（acct_key 起）是 merge 之后的派生列，
    # 由 link_accounts() 回填——先建好列，配对可随时重跑，不用重新导入流水。
    "flow": [
        ("src", "VARCHAR"), ("src_row", "BIGINT"),
        ("acct_no", "VARCHAR"), ("acct_name", "VARCHAR"), ("bank_name", "VARCHAR"),
        ("company_name", "VARCHAR"), ("short_name", "VARCHAR"),
        ("txn_date", "DATE"), ("txn_time", "VARCHAR"), ("serial", "VARCHAR"),
        ("direction", "VARCHAR"),
        ("income", "DECIMAL(20,2)"), ("expense", "DECIMAL(20,2)"),
        ("amount", "DECIMAL(20,2)"), ("balance", "DECIMAL(20,2)"),
        ("currency", "VARCHAR"),
        ("counterparty_acct", "VARCHAR"), ("counterparty_name", "VARCHAR"),
        ("counterparty_bank", "VARCHAR"),
        ("summary", "VARCHAR"), ("channel", "VARCHAR"), ("remark", "VARCHAR"),
        # ---- 与账面账户清单的配对结果（派生，可重算）
        ("acct_key", "VARCHAR"),        # 归一后的本方账号
        ("cp_key", "VARCHAR"),          # 归一后的对方账号
        ("acct_source", "VARCHAR"),     # 本方账号来源：列 / 文件名 / 人工
        ("cp_company", "VARCHAR"),      # 对方账号对应的主体（内部转账时才有值）
        ("is_internal", "VARCHAR"),     # 1 内部 / 0 外部 / 空 = 无法判定
        ("internal_reason", "VARCHAR"),
    ],
    # ---------------------------------------------------------------- 固定资产台账
    # 客户资产清单（FA list）。期初/期末是两份文件，所以挂两个 target、两张表，
    # 合成一张会互相覆盖（见 fields.TARGETS 的注释）。
    "fa": [
        ("src", "VARCHAR"), ("src_row", "BIGINT"),
        ("company_name", "VARCHAR"), ("acct_code", "VARCHAR"),
        ("category_code", "VARCHAR"), ("category_name", "VARCHAR"),
        ("asset_group", "VARCHAR"),
        ("asset_code", "VARCHAR"), ("asset_name", "VARCHAR"),
        ("dept_name", "VARCHAR"), ("location", "VARCHAR"),
        ("start_date", "DATE"), ("disposal_date", "DATE"),
        ("life_months", "DECIMAL(20,4)"), ("residual_rate", "DECIMAL(20,4)"),
        ("orig_cost", "DECIMAL(20,2)"), ("acc_dep", "DECIMAL(20,2)"),
        ("impair", "DECIMAL(20,2)"), ("reval", "DECIMAL(20,2)"),
        ("net_book", "DECIMAL(20,2)"),
        ("status", "VARCHAR"),
    ],
    "fa_beg": [
        ("src", "VARCHAR"), ("src_row", "BIGINT"),
        ("company_name", "VARCHAR"), ("acct_code", "VARCHAR"),
        ("category_code", "VARCHAR"), ("category_name", "VARCHAR"),
        ("asset_group", "VARCHAR"),
        ("asset_code", "VARCHAR"), ("asset_name", "VARCHAR"),
        ("dept_name", "VARCHAR"), ("location", "VARCHAR"),
        ("start_date", "DATE"), ("disposal_date", "DATE"),
        ("life_months", "DECIMAL(20,4)"), ("residual_rate", "DECIMAL(20,4)"),
        ("orig_cost", "DECIMAL(20,2)"), ("acc_dep", "DECIMAL(20,2)"),
        ("impair", "DECIMAL(20,2)"), ("reval", "DECIMAL(20,2)"),
        ("net_book", "DECIMAL(20,2)"),
        ("status", "VARCHAR"),
    ],
}

FTYPE = {}
for _t, _meta in TARGETS.items():
    for _k, _lb, _ty, _c in _meta["fields"]:
        FTYPE[(_t, _k)] = _ty


def qi(name):
    """quote identifier"""
    return '"' + str(name).replace('"', '""') + '"'


_SQL_FROM_TB = re.compile(r"\bfrom\s+tb\b", re.I)
_SQL_FROM_JOURNAL = re.compile(r"\bfrom\s+journal\b", re.I)
_SQL_FROM_RECON = re.compile(r"\bfrom\s+_recon\b", re.I)


def _sniff_target(sql):
    """从 SQL 里看出查的是哪张标准表，用于挑对应的中文表头字典。

    只有 direction 这个列名在两张表里中文名不同（余额方向 / 借贷方向），
    但也值得认准。认不出就返回 None，走合并版字典。
    """
    s = re.sub(r"[\"'`]", "", sql or "")
    if _SQL_FROM_RECON.search(s):
        return "recon"
    if _SQL_FROM_TB.search(s):
        return "tb"
    if _SQL_FROM_JOURNAL.search(s):
        return "journal"
    return None


def ql(s):
    return "'" + str(s).replace("'", "''") + "'"


def _acct_norm(expr):
    """账号归一表达式：全角→半角、去空格与连字符、去前导零。

    配对（流水 ↔ 账户清单、流水 ↔ 流水）全靠它。不归一的话
    「7559 3264 1710 201」和「755932641710201」会被当成两个不同账户，
    内部转账就全漏了。expr 是 SQL 表达式，不是列名。
    """
    return ("nullif(regexp_replace(regexp_replace("
            "translate(trim(coalesce({e}::VARCHAR, '')), "
            "'０１２３４５６７８９　－—−', '0123456789 ---'), "
            "'[ 　—−-]', '', 'g'), '^0+', ''), '')").format(e=expr)


def _num(v):
    if isinstance(v, Decimal):
        return float(v)
    if isinstance(v, (date, datetime)):
        return v.isoformat()
    return v


_SHEET_BAD = re.compile(r"[\[\]:*?/\\]")


def _safe_sheet(name):
    """Excel 工作表名的合法化：去掉非法字符，截到 31 字，空名兜底。"""
    s = _SHEET_BAD.sub("·", str(name or "")).strip() or "结果"
    return s[:31]


class MergeBlocked(Exception):
    """合并被清洗达标门禁拦下。items = fail 级不达标项（含 title/detail/fix）。"""

    def __init__(self, target, items):
        self.target = target
        self.items = items or []
        super().__init__("清洗未达标，已阻止合并（%s）" % target)


class Ledger:
    def __init__(self, app_dir):
        self.app_dir = os.path.abspath(app_dir)
        # KZ_DATA_DIR 可指定别的数据目录（测试/多账套用），默认指向项目根目录的 data/。
        # 数据目录统一放到根目录 data/（duckdb、规则、用户导入 excel 等），与代码分离，
        # 便于整体替换某个模块文件夹时数据不受影响。
        self.data_dir = os.path.abspath(os.environ.get("KZ_DATA_DIR")
                                        or os.path.join(self.app_dir, "..", "data"))
        self.up_dir = os.path.join(self.data_dir, "uploads")
        self.work_dir = os.path.join(self.data_dir, "work")
        self.out_dir = os.path.join(self.data_dir, "exports")
        self.log_dir = os.path.join(self.data_dir, "logs")
        for d in (self.data_dir, self.up_dir, self.work_dir, self.out_dir, self.log_dir):
            os.makedirs(d, exist_ok=True)
        ext = os.path.join(self.app_dir, "runtime", "extensions")
        cfg = {
            "memory_limit": "4GB",
            "threads": str(max(2, min(8, os.cpu_count() or 4))),
            "temp_directory": os.path.join(self.work_dir, "spill"),
            "preserve_insertion_order": "true",
        }
        if os.path.isdir(ext):
            cfg["extension_directory"] = ext
        self.db_path = os.path.join(self.data_dir, "ledger.duckdb")
        self.conn = duckdb.connect(self.db_path, config=cfg)
        self.lock = threading.RLock()
        self._init()
        self._excel_ok = None
        self._load_excel()
        # 工作表清单缓存：(路径, mtime, 大小) -> list_sheets 结果。
        # inspect 已扫过一次，import 再扫一次是白花时间（无 dimension 的表要 2s+）。
        self._sheets_cache = {}
        # 导入进度：key = file_id|sheet -> {pct, phase, t0, done, rows, seconds}
        # 前端在 import 请求进行中轮询 /api/import/progress 拿百分比。
        self._progress = {}

    # ------------------------------------------------------------ init
    def _init(self):
        c = self.conn
        # 按「行尾分号」切分，避免宏体里出现分号时被截断
        for stmt in re.split(r";\s*\n", MACROS):
            if stmt.strip():
                try:
                    c.execute(stmt)
                except Exception as e:
                    sys.stderr.write("[macro] %s -> %s\n" % (stmt.strip()[:60], e))
        c.execute("""CREATE TABLE IF NOT EXISTS _files(
            file_id VARCHAR PRIMARY KEY, file_name VARCHAR, path VARCHAR,
            size BIGINT, kind VARCHAR, uploaded_at TIMESTAMP)""")
        try:
            c.execute("ALTER TABLE _files ADD COLUMN IF NOT EXISTS md5 VARCHAR")
        except Exception:
            pass
        c.execute("""CREATE TABLE IF NOT EXISTS _sheets(
            sid VARCHAR PRIMARY KEY, file_id VARCHAR, file_name VARCHAR,
            sheet_name VARCHAR, header_row INTEGER, header_rows INTEGER,
            table_name VARCHAR, col_names VARCHAR, n_rows BIGINT,
            target VARCHAR, mapping VARCHAR, imported_at TIMESTAMP)""")
        # 清洗流水线配置 + 清洗后的表名（stg_<sid>），老库自动补列
        for col, tp in (("clean", "VARCHAR"), ("stg_table", "VARCHAR")):
            try:
                c.execute("ALTER TABLE _sheets ADD COLUMN IF NOT EXISTS %s %s" % (col, tp))
            except Exception:
                pass
        # 账簿对照表：账簿名称 -> 公司名称 / 简称。合并时按账簿名称回填。
        c.execute("""CREATE TABLE IF NOT EXISTS _books(
            book_name VARCHAR PRIMARY KEY, company_name VARCHAR,
            short_name VARCHAR, src VARCHAR, imported_at TIMESTAMP)""")
        # 账簿名推断留痕：对照表没覆盖的账簿名，平台按名称相似度推断到最可能的公司。
        # 一条账簿名一行，记下方法 / 分数 / 中文依据 / 是否需人工确认，供复核与导出。
        # src='manual' 的行不会被自动推断覆盖。
        c.execute("""CREATE TABLE IF NOT EXISTS _book_match(
            book_name VARCHAR PRIMARY KEY, company_name VARCHAR,
            method VARCHAR, score DOUBLE, evidence VARCHAR,
            need_confirm INTEGER, src VARCHAR, updated_at TIMESTAMP)""")
        # 清洗方案库：命名保存的清洗动作，界面上从下拉框直接套用
        c.execute("""CREATE TABLE IF NOT EXISTS _recipes(
            name VARCHAR PRIMARY KEY, steps VARCHAR, sig VARCHAR,
            note VARCHAR, uses INTEGER, updated_at TIMESTAMP)""")
        for t, cols in CANON.items():
            ddl = ", ".join('"%s" %s' % (n, tp) for n, tp in cols)
            c.execute('CREATE TABLE IF NOT EXISTS "%s"(%s)' % (t, ddl))
            # 老库升级：表已存在但缺新列（辅助核算那批）
            for n, tp in cols:
                try:
                    c.execute('ALTER TABLE "%s" ADD COLUMN IF NOT EXISTS "%s" %s' % (t, n, tp))
                except Exception:
                    pass
        # 判断留痕表（_judgments / _todos）。横切所有底稿，晚建会让每个科目都返工。
        # 设计见 03_文档/V1.2.09设计_判断留痕与黄格分诊.md
        try:
            from . import judge as _judge
            _judge.ensure(c)
        except Exception as e:
            sys.stderr.write("[judge] init -> %s\n" % e)

    def _load_excel(self):
        if self._excel_ok is not None:
            return self._excel_ok
        try:
            self.conn.execute("LOAD excel")
            self._excel_ok = True
        except Exception:
            try:
                self.conn.execute("INSTALL excel")
                self.conn.execute("LOAD excel")
                self._excel_ok = True
            except Exception:
                self._excel_ok = False
        return self._excel_ok

    def _list_sheets(self, path):
        """带缓存的工作表清单。

        inspect 探测过后 import 还会再要一次；每次都要打开 zip 读工作表结构，
        没有 dimension 的表还要全扫 XML 数行列（10 万行约 2s）。按文件路径 +
        修改时间 + 大小做键，文件没变就直接复用。
        """
        try:
            st = os.stat(path)
            key = (path, st.st_mtime_ns, st.st_size)
        except OSError:
            return xlsxinfo.list_sheets(path)
        hit = self._sheets_cache.get(key)
        if hit is not None:
            return hit
        out = xlsxinfo.list_sheets(path)
        if len(self._sheets_cache) > 32:
            self._sheets_cache.clear()
        self._sheets_cache[key] = out
        return out

    # ------------------------------------------------------- import progress
    _PROG_MAX = 64

    def _prog_key(self, fid, sheet):
        return "%s|%s" % (fid, sheet)

    def prog_set(self, key, pct, phase, **extra):
        """更新导入进度（内存 dict，前端轮询即时可见）。"""
        d = self._progress.get(key)
        if d is None or d.get("done"):
            if len(self._progress) > self._PROG_MAX:
                for k in [k for k, v in self._progress.items() if v.get("done")]:
                    self._progress.pop(k, None)
            d = {"t0": time.time(), "rows": 0, "seconds": 0, "done": False}
            self._progress[key] = d
        d["pct"] = max(0, min(100, int(pct)))
        d["phase"] = phase
        d["elapsed"] = round(time.time() - d["t0"], 2)
        if extra:
            d.update(extra)
        return d

    def prog_get(self, key):
        d = self._progress.get(key)
        if not d:
            return {"pct": 0, "phase": "", "running": False,
                    "elapsed": 0, "rows": 0, "seconds": 0, "done": False}
        out = dict(d)
        out["running"] = not out.get("done")
        out.pop("t0", None)
        return out

    def _run_with_progress(self, prog_key, phase, lo, hi, fn, est=3.0):
        """跑一个拿不到内部进度的长操作，用时间来平滑推进百分比。

        DuckDB 的 read_xlsx 是一次性调用，中间取不到「已解析多少行」。实测它的
        耗时与解压后 XML 字节数基本成正比（约 30~40MB/s），所以用
        bytes/35MBps 估总时长，再按时间线性推到 hi，留一段给后续建表阶段。
        真正精确的数字在结束时给出（总行数 + 总耗时）。
        """
        if not prog_key:
            return fn()
        stop = threading.Event()

        def _tick():
            t_start = time.time()
            while not stop.is_set():
                el = time.time() - t_start
                pct = lo + (hi - lo) * min(1.0, el / max(0.05, est))
                self.prog_set(prog_key, pct, phase)
                stop.wait(0.1)

        th = threading.Thread(target=_tick, daemon=True)
        th.start()
        try:
            return fn()
        finally:
            stop.set()
            th.join(0.3)

    # ------------------------------------------------------------ files
    def add_file(self, name, stream, size=None):
        name = os.path.basename(name or "upload.dat")
        fid = uuid.uuid4().hex[:12]
        safe = re.sub(r"[^\w\u4e00-\u9fa5\.\-\(\) ]", "_", name)
        path = os.path.join(self.up_dir, fid + "_" + safe)
        h = hashlib.md5()
        n = 0
        with open(path, "wb") as f:
            remain = size
            while remain is None or remain > 0:
                want = (4 << 20) if remain is None else min(4 << 20, remain)
                b = stream.read(want)
                if not b:
                    break
                h.update(b)
                n += len(b)
                f.write(b)
                if remain is not None:
                    remain -= len(b)
        low = name.lower()
        if low.endswith((".csv", ".txt", ".tsv")):
            kind = "csv"
        elif low.endswith((".xlsx", ".xlsm")):
            kind = "xlsx"
        elif low.endswith(".xls"):
            kind = "xls"
        elif low.endswith((".html", ".htm")):
            # 只是"声称"是网页；到底是不是，由 _ensure_xlsx 按文件头判定
            kind = "html"
        else:
            kind = "other"
        md5 = h.hexdigest()
        with self.lock:
            # 同一个文件重复上传：复用已有 file_id，避免同一份数据被导入两遍
            dup = self.conn.execute(
                "SELECT file_id FROM _files WHERE md5=? AND size=?", [md5, n]).fetchone()
            if dup:
                try:
                    os.remove(path)
                except OSError:
                    pass
                old = self._file(dup[0])
                old.update({"size": n, "kind": kind, "md5": md5, "duplicate": True})
                return old
            self.conn.execute(
                "INSERT INTO _files (file_id,file_name,path,size,kind,uploaded_at,md5) "
                "VALUES (?,?,?,?,?,?,?)",
                [fid, name, path, n, kind, datetime.now(), md5])
        return {"file_id": fid, "file_name": name, "size": n, "kind": kind,
                "md5": md5}

    def _file(self, fid):
        r = self.conn.execute(
            "SELECT file_id,file_name,path,size,kind FROM _files WHERE file_id=?",
            [fid]).fetchone()
        if not r:
            raise ValueError("文件不存在：%s" % fid)
        return dict(zip(["file_id", "file_name", "path", "size", "kind"], r))

    # ------------------------------------------------ 压缩包展开（导入前）
    # 包内的 .txt 一般不收：ERP 打包里它是 readme/导出说明，收进来会变成一张只有几行
    # 文字的"表"混进待导入。真要导 .txt 账表，直接拖那个文件走普通入口即可。
    ZIP_INNER_EXT = (".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".html", ".htm")

    @staticmethod
    def _zip_name(info):
        """压缩包成员名。没打 UTF-8 标志的包（Windows 中文工具常见）按 GBK 还原。"""
        name = info.filename or ""
        if not (info.flag_bits & 0x800):
            try:
                name = name.encode("cp437").decode("gbk")
            except Exception:
                pass
        return name.replace("\\", "/")

    def expand_zip(self, fid, max_files=300, max_bytes=500 << 20):
        """把客户导出的压缩包在服务端解开，包内每个表格注册成一个待导入文件。

        客户 ERP 常按账/按主体导出成 zip（如 DMOL_xxx.zip 里是明细账 xlsx），
        这些**不是**平台数据包（没有 manifest.json）。以前前端一律把 .zip 归到
        「导入数据包」入口，导致这类文件两条路都进不去，只能人工解压。
        """
        f = self._file(fid)
        try:
            zf = zipfile.ZipFile(f["path"])
        except zipfile.BadZipFile:
            raise ValueError("压缩包已损坏，或不是 zip 格式（可先在本机解压确认）")
        except OSError as e:
            raise ValueError("读不到压缩包：%s" % e)

        out, skipped, total = [], [], 0
        tmp = os.path.join(self.work_dir, "unz_" + uuid.uuid4().hex[:8])
        os.makedirs(tmp, exist_ok=True)
        try:
            with zf as z:
                if "manifest.json" in (z.namelist() or []):
                    return {"data_pack": True, "files": [], "skipped": []}
                for i in z.infolist():
                    if i.is_dir():
                        continue
                    name = self._zip_name(i)
                    base = os.path.basename(name)
                    if not base:
                        continue
                    low = base.lower()
                    if base.startswith("~$") or base.startswith("._") \
                            or low in ("thumbs.db", ".ds_store", "desktop.ini"):
                        skipped.append({"name": name, "why": "Office 临时文件／系统文件"})
                        continue
                    ext = os.path.splitext(low)[1]
                    if ext == ".zip":
                        skipped.append({"name": name, "why": "压缩包里还有压缩包，请先解开这一层"})
                        continue
                    if ext not in self.ZIP_INNER_EXT:
                        skipped.append({"name": name, "why": "非表格文件，已忽略"})
                        continue
                    if len(out) >= max_files:
                        skipped.append({"name": name, "why": "包内文件过多（上限 %d 个），已跳过" % max_files})
                        continue
                    total += int(i.file_size or 0)
                    if total > max_bytes:
                        raise ValueError("压缩包解压后超过 %d MB，请分批导入" % (max_bytes >> 20))
                    dst = os.path.join(tmp, "f%04d%s" % (len(out), ext))
                    try:
                        with z.open(i) as src, open(dst, "wb") as fo:
                            shutil.copyfileobj(src, fo, 1 << 20)
                    except Exception as e:
                        # 单个成员坏了或被单独加了密码，跳过它，其余继续
                        why = "已加密，无法解压" if "encrypt" in str(e).lower() else "解压失败：%s" % str(e)[:60]
                        skipped.append({"name": name, "why": why})
                        continue
                    try:
                        with open(dst, "rb") as fh:
                            info = self.add_file(base, fh)
                    except Exception as e:
                        skipped.append({"name": name, "why": "登记失败：%s" % str(e)[:60]})
                        continue
                    info["rel"] = name
                    info["from_zip"] = True
                    out.append(info)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
        return {"files": out, "skipped": skipped, "data_pack": False}

    @staticmethod
    def _head_kind(head):
        """按文件头判真实格式：ole2 / zip / sml / html / unknown。

        只看前 4KB，够用且便宜；顺序有意义（OLE2 > zip > xml > html）。
        """
        h8 = (head or b"")[:8]
        if xlsread.is_xls(h8):
            return "ole2"
        if h8[:4] == b"PK\x03\x04":
            return "zip"
        s = (head or b"")[:4096].decode("utf-8", "ignore").lower()
        if "<?xml" in s and ("mso-application" in s
                             or "urn:schemas-microsoft-com:office:spreadsheet" in s):
            return "sml"
        if htmlread.looks_html(head):
            return "html"
        return "unknown"

    @staticmethod
    def _zip_is_xlsx(path):
        """zip 里有没有工作簿——用来区分"xlsx 改了后缀"和"docx/pptx/压缩包改了后缀"。"""
        try:
            with zipfile.ZipFile(path) as z:
                return "xl/workbook.xml" in (z.namelist() or [])
        except Exception:
            return False

    def _xls_to_xlsx(self, f):
        """老式 .xls（OLE2/BIFF8）→ .conv.xlsx，同一文件重复用不再转。"""
        src = f["path"]
        out = os.path.splitext(src)[0] + ".conv.xlsx"
        try:
            if (not os.path.exists(out)) or os.path.getmtime(out) < os.path.getmtime(src):
                with open(src, "rb") as fh:
                    xlsread.xls_to_xlsx(fh.read(), out)
        except xlsread.XlsError as e:
            raise ValueError("读取 .xls 失败：%s" % e)
        except Exception as e:
            raise ValueError("转换 .xls 失败：%s" % e)
        f["src_name"] = f.get("file_name")
        f["path"] = out
        f["kind"] = "xlsx"
        f["file_name"] = re.sub(r"\.xls$", ".xlsx", f.get("file_name") or "", flags=re.I)
        return f

    def _html_to_xlsx(self, f):
        """网页表格（ERP/网银「导出 Excel」的真实形态）→ <原名>.html.xlsx 缓存。

        转成真正的 xlsx 之后，预览 / 导入 / 清洗 / 合并全部走原有通道，一行都不用改。
        """
        src = f["path"]
        out = os.path.splitext(src)[0] + ".html.xlsx"
        tables, note = 0, ""
        try:
            if (not os.path.exists(out)) or os.path.getmtime(out) < os.path.getmtime(src):
                meta = htmlread.to_xlsx(src, out)
                tables = len(meta)
                note = next((m.get("note") for m in meta if m.get("note")), "")
        except htmlread.HtmlError as e:
            raise ValueError(str(e))
        except Exception as e:
            raise ValueError("转换网页表格失败：%s" % e)
        f["src_name"] = f.get("file_name")
        f["path"] = out
        f["kind"] = "xlsx"
        f["file_name"] = re.sub(r"\.(xls|html?|csv|tsv|txt)$", ".xlsx",
                                f.get("file_name") or "", flags=re.I)
        f["converted_from"] = "html"
        if tables:
            f["html_tables"] = tables
        if note:
            f["html_note"] = note
        return f

    def _ensure_xlsx(self, f):
        """读取前的统一入口：按**文件头**判真实格式，必要时先转成 xlsx。

        为什么不能只看后缀（三条都是现场踩过的）：
        - 加密的 xlsx 是 OLE2 复合文档（与 .xls 同头），直接进 zipfile 只会报
          "File is not a zip file"，前端翻译成「可能已损坏」，把"缺密码"带偏；
        - 大量 ERP / 网银「导出 Excel」产出的其实是 **HTML 表格**，后缀却写成 .xls，
          必须解析后转成真 xlsx，否则下游 read_xlsx 一律失败；
          （典型特征：Excel 打开是乱码/全是问号，把后缀改成 .html 就正常）
        - 反过来，真 xlsx 被改成 .xls/.html 也要认，不能靠后缀赌。

        转换缓存：.xls → .conv.xlsx、网页 → .html.xlsx，都按 mtime 判新旧，
        inspect 与 import 各调一次也只会转一遍。
        """
        kind = f.get("kind")
        src = f.get("path") or ""
        try:
            with open(src, "rb") as fh:
                head = fh.read(4096)
        except OSError as e:
            raise ValueError("读不到文件：%s" % e)
        head_kind = self._head_kind(head)

        if kind == "csv":
            # 有些系统把网页表格存成 .csv（同样是"改后缀成 .html 才看得见内容"）。
            # 按内容纠正一下，其余 CSV 一律保持原样，不碰分隔符/编码的既有逻辑。
            if head_kind == "html":
                return self._html_to_xlsx(f)
            if head_kind == "zip" and self._zip_is_xlsx(src):
                f["kind"] = "xlsx"
                return f
            return f
        suspicious = kind in ("xls", "xlsx", "html")

        if head_kind == "ole2":
            if kind in ("xlsx", "html"):
                raise ValueError("该文件已加密（受密码保护）：Excel 打开它需要密码。"
                                 "请在 Excel 里「另存为」一份不加密的 .xlsx 再导入")
            if kind == "xls":
                return self._xls_to_xlsx(f)
            return f                       # other：维持旧行为（当 CSV 读，不误伤 doc 改后缀）

        if head_kind == "zip":
            if self._zip_is_xlsx(src):
                f["kind"] = "xlsx"         # 真工作簿：不管后缀写成 .xls 还是 .html 都认
                return f
            if suspicious:
                raise ValueError("这不是 Excel 工作簿：文件是个 zip 压缩包"
                                 "（可能是 .docx / .pptx，或普通压缩包改了后缀）。"
                                 "压缩包请直接拖进「导入原始数据」，平台会自动解开")
            return f

        if head_kind == "sml":
            # SpreadsheetML 无论叫什么后缀都不能当 CSV 读（读出来全是标签），
            # 它只可能来自 Excel「另存为 XML 表格」，直接指明怎么转
            raise ValueError("这是 Excel 2003 XML 表格（SpreadsheetML），暂不支持。"
                             "请在 Excel 里打开后「另存为」.xlsx 再导入")

        if head_kind == "html":
            return self._html_to_xlsx(f)

        if suspicious:
            raise ValueError("无法识别的格式：文件内容既不是 Excel 工作簿（xlsx / OLE2），"
                             "也不是网页表格（HTML <table>）。"
                             "请用 Excel 打开后「另存为」.xlsx 再导入")
        return f

    def _csv_utf8(self, path):
        """GBK 等编码的 CSV 先转成 UTF-8，返回 (新路径, 分隔符)"""
        delim, enc = xlsxinfo.sniff_csv(path)
        if enc in ("utf-8", "utf-8-sig"):
            return path, delim
        out = path + ".utf8.csv"
        if not os.path.exists(out):
            with open(path, "r", encoding=enc, errors="replace", newline="") as fi, \
                 open(out, "w", encoding="utf-8", newline="") as fo:
                while True:
                    chunk = fi.read(8 << 20)
                    if not chunk:
                        break
                    fo.write(chunk)
        return out, delim

    # ------------------------------------------------------------ inspect
    def inspect(self, fid, target=None):
        f = self._ensure_xlsx(self._file(fid))
        if f["kind"] == "xlsx":
            if not self._load_excel():
                raise RuntimeError("Excel 扩展不可用，无法读取 xlsx")
            sheets = self._list_sheets(f["path"])
            out = []
            for s in sheets:
                item = {"sheet": s["name"], "rows": s["rows"], "cols": s["cols"]}
                # 图表页 / 宏表 / 非标准工作表：read_xlsx 读不了，标出来跳过
                if s.get("path") and not str(s["path"]).startswith("xl/worksheets/"):
                    item.update({"preview": [], "header_row": 1, "header_rows": 1,
                                 "mapping": {}, "empty": True, "note": "非数据工作表，已跳过"})
                    out.append(item)
                    continue
                try:
                    prev, cols = self._xlsx_preview(f["path"], s)
                except Exception as e:
                    prev, cols = [], 0
                    item["error"] = str(e)[:200]
                if not prev:
                    # 三种"没有预览"要分清：读失败 / 疑似损坏 / 真空表。
                    # 原来一律写「空表（没有数据行）」，把损坏和丢数据都盖掉了。
                    if item.get("error"):
                        item.update({"preview": [], "header_row": 1, "header_rows": 1,
                                     "mapping": {}, "empty": True, "broken": True,
                                     "note": "读取失败：%s" % item["error"]})
                    elif s.get("broken"):
                        item.update({"preview": [], "header_row": 1, "header_rows": 1,
                                     "mapping": {}, "empty": True, "broken": True,
                                     "note": s.get("reason") or "疑似损坏：解析不出任何单元格"})
                    else:
                        item.update({"preview": [], "header_row": 1, "header_rows": 1,
                                     "mapping": {}, "empty": True,
                                     "note": item.get("note") or "空表（没有数据行）"})
                else:
                    hr, hrows = detect_header(prev, target)
                    names = self._header_names(prev, hr, hrows)
                    item.update({"preview": prev, "header_row": hr, "header_rows": hrows,
                                 "columns": names,
                                 "mapping": auto_map(target, names) if target else {}})
                    if target == "flow":
                        self._attach_flow_acct(item, prev, hr, f["file_name"])
                out.append(item)
            return {"kind": "xlsx", "file": f, "sheets": out}
        else:
            p, delim = self._csv_utf8(f["path"])
            sql = ("SELECT * FROM read_csv(?, header=false, all_varchar=true, delim=?, "
                   "quote='\"', escape='\"', ignore_errors=true, sample_size=20480) LIMIT 30")
            cur = self.conn.execute(sql, [p, delim])
            cols = [d[0] for d in cur.description]
            rows = [[_s(v) for v in r] for r in cur.fetchall()]
            hr, hrows = detect_header(rows, target)
            return {"kind": "csv", "file": f, "delim": delim, "sheets": [{
                "sheet": f["file_name"], "rows": 0, "cols": len(cols),
                "preview": rows, "header_row": hr, "header_rows": hrows,
                "columns": self._header_names(rows, hr, hrows),
                "mapping": auto_map(target, self._header_names(rows, hr, hrows)) if target else {}}]}

    def _attach_flow_acct(self, item, rows, header_row, file_name):
        """给银行流水补「本方账号」——它极少在列里，几乎总在表头上方的键值行。

        银行的账户明细导出长这样：
            某某银行账户明细
            本方账号   755932641710201
            账户名称   某某公司
            交易日期  金额  余额     ← 表头从这行才算
        所以 auto_map 永远找不到账号列。不单独捞出来的话，整个文件的流水
        都不知道属于哪个账户，后面跟账户清单就对不上。
        """
        mp = item.get("mapping") or {}
        if mp.get("acct_no"):
            return
        hint = self._flow_acct_kv(rows, header_row) or self._flow_acct_name(file_name)
        if hint:
            if hint.get("acct_no"):
                mp["_flow_acct"] = hint["acct_no"]
                mp["_flow_acct_src"] = hint.get("from") or "表头/文件名"
            if hint.get("acct_name"):
                mp["_flow_acct_name"] = hint["acct_name"]
            item["mapping"] = mp
            item["acctHint"] = hint

    @staticmethod
    def _flow_acct_kv(rows, header_row):
        """表头区「标签 值」相邻两格里的账号 / 户名。"""
        lbl = re.compile(r"(账号|帐号|账户号码|卡号)")
        name_lbl = re.compile(r"(账户名称|户名|开户名称|客户名称)")
        other = re.compile(r"对方|收款人|付款人|收方|付方|对手")
        hit = {}
        for i in range(max(0, (header_row or 1) - 1)):
            row = [("" if c is None else str(c).strip()) for c in rows[i]]
            for j, cell in enumerate(row[:-1]):
                if not cell or other.search(cell):
                    continue
                val = next((row[k] for k in range(j + 1, min(j + 3, len(row)))
                            if row[k]), "")
                if not val:
                    continue
                if lbl.search(cell) and re.fullmatch(r"[\d\-\s]{8,}", val) \
                        and sum(ch.isdigit() for ch in val) >= 8:
                    hit.setdefault("acct_no", val)
                    hit.setdefault("from", "表头第 %d 行" % (i + 1))
                elif name_lbl.search(cell) and len(val) >= 2:
                    hit.setdefault("acct_name", val)
        return hit or None

    @staticmethod
    def _flow_acct_name(file_name):
        """文件名里的账号，如「账户明细查询_某某公司_03006017389_2026.xlsx」。"""
        base = os.path.splitext(os.path.basename(file_name or ""))[0]
        for m in re.findall(r"(?<!\d)(\d{8,24})(?!\d)", base):
            if len(m) == 8 and m.startswith("20"):
                continue          # 20260630 这种是查询期间，不是账号
            return {"acct_no": m, "from": "文件名"}
        return None

    def _xlsx_preview(self, path, s):
        rows, cols = s["rows"], s["cols"]
        # 完全空的工作表（封面、模板页、财务软件导出的空白首表）：DuckDB 的 excel
        # 扫描器会直接抛 BinderException "No rows found in xlsx file"，而不是返回 0 行。
        # 所以这里不打数据库，直接给空预览。
        if rows <= 0 or cols <= 0:
            return [], 0
        # ---- 快速路径：直读 sheet XML 开头，只解压到够 40 行 ----
        # 原来走 read_xlsx(range=A1:xx60, LIMIT 40)，但实测 range/LIMIT 并不能让
        # DuckDB 跳过解析：它仍要解压整张工作表（真实序时账的 sheet1.xml 解压后
        # 47MB），40 行预览和全表导入几乎同样耗时。预览只要前 40 行，几十 KB 就够。
        try:
            cells = xlsxinfo.preview_cells(path, s.get("path"), 40)
        except Exception:
            cells = None
        if cells:
            keys = [xlsxinfo.num_to_col(c).encode() for c in range(1, cols + 1)]
            data = []
            for rn in range(1, 41):
                d = cells.get(rn) or {}
                data.append([_s(d.get(k)) for k in keys])
            if any(any((v or "").strip() for v in r) for r in data):
                return data, cols
        # ---- 兜底：XML 直读失败（加密工作簿、异常合并区等）时回到 read_xlsx ----
        # 有维度就显式给 range；只有 1 行/1 列的表也必须给，否则同样可能报 "No rows found"
        rng = ", range=%s" % ql("A1:%s%d" % (
            xlsxinfo.num_to_col(cols), max(1, min(rows, 60))))
        sql = "SELECT * FROM read_xlsx(%s, sheet=%s, header=false, all_varchar=true%s) LIMIT 40" % (
            ql(path), ql(s["name"]), rng)
        try:
            cur = self.conn.execute(sql)
        except Exception:
            # dimension 写坏（整表只有 A1、被工具改过）时，退化成固定区域再试一次
            try:
                cur = self.conn.execute(
                    "SELECT * FROM read_xlsx(%s, sheet=%s, header=false, all_varchar=true, "
                    "range=%s) LIMIT 40" % (ql(path), ql(s["name"]), ql("A1:ZZ300")))
            except Exception:
                return [], 0
        data = [[_s(v) for v in r] for r in cur.fetchall()]
        ncols = len(cur.description)
        if not any((v or "").strip() for r in data for v in r):
            return [], 0
        return data, ncols

    def _header_names(self, preview, header_row, header_rows):
        """把表头行拼成列名；多行表头时对上方行做横向左填充，再用 _ 连接。"""
        hr = max(1, int(header_row or 1))
        n = int(header_rows or 1)
        block = preview[hr - 1: hr - 1 + n]
        if not block:
            return []
        width = max(len(r) for r in block)
        norm = []
        for i, r in enumerate(block):
            r = list(r) + [""] * (width - len(r))
            if i < len(block) - 1:  # 非最后一行左填充（合并单元格）
                last = ""
                for j in range(width):
                    if r[j].strip():
                        last = r[j].strip()
                    else:
                        r[j] = last
            norm.append(r)
        names = []
        for j in range(width):
            parts = [norm[i][j].strip() for i in range(len(norm))]
            parts = [p for p in parts if p and p not in ("-", "/", "—")]
            nm = "_".join(dict.fromkeys(parts)) if parts else ""
            names.append(nm or ("列%d" % (j + 1)))
        # 去重：必须保证「结果里每个名字唯一」，而不是只给重复项加序号。
        # 列名是 金额 / 金额_2 / 金额 时，简单计数会生成两个「金额_2」，
        # DuckDB 会自作主张改名成 金额_2_1，映射指向的列就跟用户看到的对不上了。
        out = []
        for nm in names:
            cand, k = nm, 2
            while cand in out:
                cand = "%s_%d" % (nm, k)
                k += 1
            out.append(cand)
        return out

    # ------------------------------------------------------------ import
    def _read_raw(self, f, sheet, hr, hrows, table, prog_key=None):
        """把一个工作表整表读成 <table>（全部文本，不清洗），列名取表头。

        返回列名列表。import_sheet 和账簿对照表导入共用这一段。
        prog_key 非空时，解析期间持续上报进度百分比。
        """
        f = self._ensure_xlsx(f)
        stg = table + "_src"
        s = None
        if f["kind"] == "xlsx":
            if not self._load_excel():
                raise RuntimeError("Excel 扩展不可用")
            info = {s["name"]: s for s in self._list_sheets(f["path"])}
            s = info.get(sheet)
            if s and (s["rows"] <= 0 or s["cols"] <= 0):
                raise ValueError("工作表「%s」是空表，没有可导入的数据（换一张表试试）" % sheet)
            rng = ""
            if s and s["rows"] > 0 and s["cols"] > 0:
                rng = ", range=%s" % ql("A1:%s%d" % (
                    xlsxinfo.num_to_col(s["cols"]), s["rows"]))
            # 解析一次要 1~5 秒且期间拿不到已完成行数，用解压后字节数估时长做进度
            nbytes = (s or {}).get("bytes") or 0
            est = (max(0.3, nbytes / 33e6) if nbytes
                   else max(0.3, float((s or {}).get("rows") or 0) / 30000.0))
            self._run_with_progress(
                prog_key, "解析 Excel 数据", 6, 86,
                lambda: self.conn.execute(
                    "CREATE OR REPLACE TABLE %s AS SELECT * FROM read_xlsx(%s, sheet=%s, "
                    "header=false, all_varchar=true%s)" % (stg, ql(f["path"]), ql(sheet), rng)),
                est=est)
        else:
            p, delim = self._csv_utf8(f["path"])
            self.conn.execute(
                "CREATE OR REPLACE TABLE %s AS SELECT * FROM read_csv(%s, header=false, "
                "all_varchar=true, delim=%s, quote='\"', escape='\"', ignore_errors=true)"
                % (stg, ql(p), ql(delim)))
        cur = self.conn.execute("SELECT * FROM %s LIMIT %d" % (stg, hr + hrows))
        stg_cols = [d[0] for d in cur.description]
        head = cur.fetchall()
        if not head:
            self.conn.execute("DROP TABLE IF EXISTS %s" % stg)
            raise ValueError("该工作表没有数据")
        preview = [[_s(v) for v in r] for r in head]
        names = self._header_names(preview, hr, hrows)
        ncol = len(stg_cols)
        if len(names) < ncol:
            names += ["列%d" % (i + 1) for i in range(len(names), ncol)]
        sel = ", ".join('%s AS %s' % (qi(stg_cols[i]), qi(names[i])) for i in range(ncol))
        skip = hr - 1 + hrows
        # 建标准表（含 row_number 物化）在大表上要花近 1 秒，同样用插值推进进度
        self._run_with_progress(
            prog_key, "生成标准列", 87, 92,
            lambda: self.conn.execute(
                "CREATE OR REPLACE TABLE %s AS SELECT %s, row_number() OVER () AS _src_row "
                "FROM %s LIMIT 1000000000 OFFSET %d" % (table, sel, stg, skip)),
            est=max(0.2, float((s or {}).get("rows") or 0) / 110000.0))
        self.conn.execute("DROP TABLE IF EXISTS %s" % stg)
        return names

    def import_sheet(self, fid, sheet, header_row=1, header_rows=1, target=None, mode="replace"):
        f = self._file(fid)
        sid = hashlib.md5((fid + "|" + sheet).encode()).hexdigest()[:10]
        # mode=new 时给一个全新 sid，避免覆盖旧配置（同文件同 sheet 重导默认仍是 replace）
        if mode == "new":
            base = sid
            k = 2
            while self.conn.execute("SELECT count(*) FROM _sheets WHERE sid=?", [sid]).fetchone()[0]:
                sid = hashlib.md5(("%s|#%d" % (base, k)).encode()).hexdigest()[:10]
                k += 1
        raw = "raw_" + sid
        hr = max(1, int(header_row or 1))
        hrows = max(1, int(header_rows or 1))
        t0 = time.time()
        pkey = self._prog_key(fid, sheet)
        self.prog_set(pkey, 1, "准备导入")
        try:
            with self.lock:
                self.prog_set(pkey, 4, "读取工作表结构")
                names = self._read_raw(f, sheet, hr, hrows, raw, prog_key=pkey)
                self.prog_set(pkey, 93, "统计行数")
                n = self.conn.execute("SELECT count(*) FROM %s" % raw).fetchone()[0]

                mapping = auto_map(target, names) if target else {}
                # 银行流水：表头键值行里的本方账号要跟着一起落库，否则合并时
                # 整张表都不知道自己属于哪个账户（inspect 阶段的推断到这里会丢）。
                # 注意 raw 表的列名取自表头行，表头区那几行不在里面，必须回原始文件读。
                if target == "flow" and not mapping.get("acct_no"):
                    try:
                        meta = next((x for x in self._list_sheets(f["path"])
                                     if x["name"] == sheet), None)
                        if meta:
                            prev, _ = self._xlsx_preview(f["path"], meta)
                            _tmp = {"mapping": mapping}
                            self._attach_flow_acct(_tmp, prev, hr, f["file_name"])
                            mapping = _tmp["mapping"]
                    except Exception:
                        pass
                self.prog_set(pkey, 96, "登记工作表")
                overwrote = bool(self.conn.execute(
                    "SELECT count(*) FROM _sheets WHERE sid=?", [sid]).fetchone()[0])
                self.conn.execute("DELETE FROM _sheets WHERE sid=?", [sid])
                # 重新导入：旧的清洗表和剔除记录一并清掉，避免脏表残留
                for t in ("stg_" + sid, "_dropped_" + sid):
                    self.conn.execute('DROP TABLE IF EXISTS "%s"' % t)
                self.conn.execute(
                    "INSERT INTO _sheets (sid,file_id,file_name,sheet_name,header_row,header_rows,"
                    "table_name,col_names,n_rows,target,mapping,imported_at,clean,stg_table) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    [sid, fid, f["file_name"], sheet, hr, hrows, raw,
                     json.dumps(names, ensure_ascii=False), n, target,
                     json.dumps(mapping, ensure_ascii=False), datetime.now(),
                     # 清洗动作一律留空：默认配方会「猜」使用者的意图，猜错的代价
                     # （把真实分录当重复删、把父科目当明细留）比省下的一次点击大得多。
                     # 正确做法是让人或 AI 按「清洗达标标准」自己配，或套用方案库里
                     # 以前验证过的配方；配没配对，最后由「数据检查」五组兜底。
                     json.dumps([], ensure_ascii=False),
                     None])
        except Exception as e:
            self.prog_set(pkey, 100, "导入失败", done=True,
                          error=str(e)[:200], rows=0)
            raise
        el = time.time() - t0
        self.prog_set(pkey, 100, "完成", done=True, rows=int(n), seconds=round(el, 2))
        warnings = []
        if overwrote:
            warnings.append("同文件同工作表之前已导入过，本次已覆盖旧的清洗/映射配置"
                            "（如需保留旧配置，导入时选「作为新表」）。")
        if n and n % 1000000 == 0:
            warnings.append("本表恰好 %d 行，可能是导出工具的分卷或截断，请确认是否还有其他分卷。" % n)
        return {"sid": sid, "columns": names, "n_rows": n,
                "elapsed": round(el, 2), "mapping": mapping,
                "progress_key": pkey, "overwrote": overwrote,
                "warnings": warnings}

    # ------------------------------------------------------------ 标准表数据包
    def load_pack(self, zip_path, tables=None, replace=True):
        """导入标准表数据包（zip 内 tb.parquet / journal.parquet / manifest.json）。

        为什么需要这条路：客户资料在别的工具里已经清洗成标准表（列名与本平台 CANON 一致）
        时，再让人导回 xlsx 是多余的往返——而且 xlsx 转一圈还会丢掉已经归一好的期间和金额。

        按列名映射灌库：包里有的列按 CANON 类型转换，没有的补 NULL，
        所以包里少几列也能用，但会记进 `missing` 提醒。
        """
        import shutil
        import tempfile

        if not os.path.isfile(zip_path):
            raise ValueError("数据包不存在：%s" % zip_path)
        work = tempfile.mkdtemp(prefix="ptwpack_")
        try:
            with zipfile.ZipFile(zip_path) as z:
                names = z.namelist()
                z.extractall(work)
            manifest = {}
            mp = os.path.join(work, "manifest.json")
            if os.path.isfile(mp):
                with io.open(mp, encoding="utf-8") as fh:
                    manifest = json.load(fh)
            fmt = str(manifest.get("format") or "")
            if fmt and not fmt.startswith("kanzhang/"):
                raise ValueError("不认识的数据包格式：%s（本平台认 kanzhang/v1）" % fmt)

            want = [t for t in (tables or list(CANON)) if t in CANON]
            done, missing_all = [], {}
            with self.lock:
                for t in want:
                    pq = os.path.join(work, t + ".parquet")
                    if not os.path.isfile(pq):
                        continue
                    have = [d[0] for d in self.conn.execute(
                        "SELECT * FROM read_parquet(?) LIMIT 0", [pq]).description]
                    sel, missing = [], []
                    for n, tp in CANON[t]:
                        if n in have:
                            sel.append('TRY_CAST("%s" AS %s) AS "%s"' % (n, tp, n))
                        else:
                            sel.append('NULL::%s AS "%s"' % (tp, n))
                            missing.append(n)
                    verb = "CREATE OR REPLACE TABLE" if replace else "CREATE TABLE"
                    self.conn.execute(
                        '%s "%s" AS SELECT %s FROM read_parquet(?)' % (verb, t, ", ".join(sel)),
                        [pq])
                    n_rows = self.conn.execute('SELECT count(*) FROM "%s"' % t).fetchone()[0]
                    done.append({"table": t, "rows": n_rows, "cols": len(have)})
                    if missing:
                        missing_all[t] = missing
                # 记一条来源，让「数据查询」页能看出这批数是哪里来的
                try:
                    nm = os.path.basename(zip_path)
                    fid = "pack_" + (hashlib.md5(
                        ("%s|%s" % (nm, time.time())).encode("utf-8")).hexdigest()[:10])
                    self.conn.execute(
                        "INSERT OR REPLACE INTO _files "
                        "(file_id,file_name,path,size,kind,uploaded_at,md5) "
                        "VALUES (?,?,?,?,?,?,?)",
                        [fid, nm, zip_path, os.path.getsize(zip_path), "pack",
                         datetime.now(), ""])
                except Exception:
                    pass
            return {"format": fmt or "kanzhang/v1", "createdAt": manifest.get("created_at"),
                    "tables": done, "missing": missing_all,
                    "files": len(names)}
        finally:
            shutil.rmtree(work, ignore_errors=True)

    # ------------------------------------------------------------ state
    def state(self):
        with self.lock:
            rows = self.conn.execute(
                "SELECT sid,file_id,file_name,sheet_name,header_row,header_rows,table_name,"
                "col_names,n_rows,target,mapping,clean,stg_table FROM _sheets "
                "ORDER BY file_name,sheet_name"
            ).fetchall()
            sheets = []
            for r in rows:
                stg = r[12]
                n_stg = 0
                cols = json.loads(r[7] or "[]")
                if stg:
                    try:
                        n_stg = self.conn.execute('SELECT count(*) FROM "%s"' % stg).fetchone()[0]
                        # 清洗会新增列（填充固定值、级次推断、辅助核算拆分），
                        # 映射页要用清洗后的列，不然新列选不到
                        cols = [c for c in clean.columns_of(self.conn, stg) if c != "_src_row"]
                    except Exception:
                        stg, n_stg = None, 0
                sheets.append({
                    "sid": r[0], "file_id": r[1], "file_name": r[2], "sheet_name": r[3],
                    "header_row": r[4], "header_rows": r[5], "table_name": r[6],
                    "columns": cols, "n_rows": r[8],
                    "target": r[9], "mapping": json.loads(r[10] or "{}"),
                    "clean": json.loads(r[11] or "[]"), "stg": bool(stg),
                    "n_stg": n_stg})
            stats = {}
            for t in CANON:
                try:
                    n = self.conn.execute('SELECT count(*) FROM "%s"' % t).fetchone()[0]
                except Exception:
                    n = 0
                stats[t] = {"rows": n, "label": TARGETS[t]["label"]}
            # 序时账期间范围：检查页用来预填默认期间
            period_range = None
            try:
                r = self.conn.execute(
                    'SELECT min(_period_norm(period)), max(_period_norm(period)) FROM journal '
                    'WHERE nullif(trim(period), \'\') IS NOT NULL').fetchone()
                if r and r[0]:
                    period_range = {"min": r[0], "max": r[1]}
            except Exception:
                pass
            try:
                dsz = os.path.getsize(self.db_path)
            except OSError:
                dsz = 0
            try:
                n_books = self.conn.execute("SELECT count(*) FROM _books").fetchone()[0]
            except Exception:
                n_books = 0
            return {"sheets": sheets, "stats": stats, "db_size": dsz,
                    "period_range": period_range,
                    "fields": {t: field_dict(t) for t in TARGETS},
                    "targets": {t: TARGETS[t]["label"] for t in TARGETS},
                    "aux_roles": [{"key": k, "label": v} for k, v in AUX_ROLES],
                    "book_modes": [
                        {"key": "col", "label": "从某一列取"},
                        {"key": "file", "label": "用文件名"},
                        {"key": "sheet", "label": "用工作表名"},
                        {"key": "const", "label": "手工填写"},
                    ],
                    "company_modes": [
                        {"key": "books", "label": "按账簿对照表匹配"},
                        {"key": "col", "label": "从某一列取"},
                        {"key": "book", "label": "等于账簿名称"},
                    ],
                    "books_rows": n_books,
                    "excel_ok": bool(self._excel_ok)}

    def _sheet(self, sid):
        r = self.conn.execute(
            "SELECT sid,table_name,col_names,target,mapping,n_rows,clean,stg_table "
            "FROM _sheets WHERE sid=?", [sid]).fetchone()
        if not r:
            raise ValueError("工作表不存在：%s" % sid)
        return {"sid": r[0], "table_name": r[1], "columns": json.loads(r[2] or "[]"),
                "target": r[3], "mapping": json.loads(r[4] or "{}"), "n_rows": r[5],
                "clean": json.loads(r[6] or "[]"), "stg_table": r[7]}

    def _data_table(self, sid):
        """当前有效的数据表：清洗过用 stg，否则用 raw。"""
        m = self._sheet(sid)
        return m["stg_table"] or m["table_name"]

    def raw_preview(self, sid, limit=100, stage="data"):
        """stage=data 看清洗后的样子（没清洗就是原始）；stage=raw 强制看原始。"""
        with self.lock:
            m = self._sheet(sid)
            tbl = m["table_name"] if stage == "raw" else (m["stg_table"] or m["table_name"])
            cur = self.conn.execute('SELECT * FROM "%s" LIMIT %d' % (tbl, min(limit, 5000)))
            cols = [d[0] for d in cur.description]
            # 内部行号列不给前端看：映射下拉里不该出现 _src_row
            keep = [i for i, c in enumerate(cols) if c != "_src_row"]
            return {"columns": [cols[i] for i in keep],
                    "rows": [[_s(x[i]) for i in keep] for x in cur.fetchall()],
                    "table": tbl}

    # ------------------------------------------------------------ 清洗
    def clean_types(self):
        return {"types": clean.types_for_ui(), "presets": clean.PRESETS,
                "groups": ["选行", "改值", "填值", "派生", "校验"]}

    def clean_preview(self, sid, upto=None, sample=30):
        with self.lock:
            m = self._sheet(sid)
            steps = m["clean"]
            res = clean.run(self.conn, sid, m["table_name"], steps, upto=upto,
                            keep=False, sample=sample, data_dir=self.data_dir)
            # 原始前 N 行做对照
            raw = self.conn.execute('SELECT * FROM "%s" LIMIT %d'
                                    % (m["table_name"], int(sample)))
            rcols = [d[0] for d in raw.description]
            rrows = [[_s(v) for v in r] for r in raw.fetchall()]
            res["raw_columns"] = rcols
            res["raw_rows"] = rrows
            res["steps_cfg"] = steps
            return res

    def clean_apply(self, sid):
        with self.lock:
            m = self._sheet(sid)
            res = clean.run(self.conn, sid, m["table_name"], m["clean"],
                            keep=True, sample=0, data_dir=self.data_dir)
            if res.get("error"):
                return res
            self.conn.execute("UPDATE _sheets SET stg_table=? WHERE sid=?",
                              ["stg_" + sid, sid])
            res["stg"] = True
            # run(sample=0) 不取样本行，这里补一份「原始 / 清洗后」对照，
            # 保证前端拿到和 /api/clean/preview 一样的结构（缺字段会直接白屏报错）
            for k, v in (("columns", []), ("rows", []),
                         ("raw_columns", []), ("raw_rows", [])):
                res.setdefault(k, v)
            try:
                c, r = self._sample("stg_" + sid, 20)
                if c:
                    res["columns"], res["rows"] = c, r
                rc, rr = self._sample(m["table_name"], 20)
                if rc:
                    res["raw_columns"], res["raw_rows"] = rc, rr
            except Exception:
                pass
            res["steps_cfg"] = m["clean"]
            # 应用成功 = 这套方案跑通了，自动登记进方案库，右上角下拉框立刻能选
            res["recipe"] = self._recipe_autosave(sid)
            return res

    def clean_apply_all(self, sids=None):
        """批量清洗：把所有配了动作的工作表逐个跑一遍。

        逐表调用 clean_apply（所以自动登记方案库、被剔除行可溯源这些都一致）。
        没配动作的表跳过并说明原因；单张表出错不中断其余表。
        sids 给定时只洗这几张。
        """
        want = set(sids or [])
        with self.lock:
            rows = self.conn.execute(
                "SELECT sid, file_name, sheet_name, clean FROM _sheets "
                "ORDER BY file_name, sheet_name").fetchall()
        items, ok_n, skip_n, err_n = [], 0, 0, 0
        t0 = time.time()
        for sid, fn, sn, cj in rows:
            if want and sid not in want:
                continue
            base = {"sid": sid, "file": fn, "sheet": sn}
            try:
                steps = [s for s in json.loads(cj or "[]") if s.get("enabled", True)]
            except Exception:
                steps = []
            if not steps:
                items.append(dict(base, status="skip", reason="没有配置清洗动作"))
                skip_n += 1
                continue
            try:
                r = self.clean_apply(sid)
            except Exception as e:
                items.append(dict(base, status="error", error=str(e)[:200]))
                err_n += 1
                continue
            if r.get("error"):
                items.append(dict(base, status="error", error=r["error"]))
                err_n += 1
                continue
            items.append(dict(base, status="ok",
                              before=r.get("total_before"), after=r.get("total_after"),
                              total_before=r.get("total_before"),
                              total_after=r.get("total_after"),
                              dropped=r.get("dropped"), n_steps=len(steps),
                              recipe=(r.get("recipe") or {}).get("name") or ""))
            ok_n += 1
        return {"ok": err_n == 0, "total": len(items), "ok_n": ok_n,
                "skipped": skip_n, "failed": err_n,
                "seconds": round(time.time() - t0, 2), "items": items}

    def _sample(self, table, n=20):
        """取前 n 行做预览，跳过内部列 _src_row。"""
        cur = self.conn.execute('SELECT * FROM "%s" LIMIT %d' % (table, int(n)))
        cols = [d[0] for d in cur.description]
        keep = [i for i, c in enumerate(cols) if c != "_src_row"]
        rows = [[_s(v) for v in [r[i] for i in keep]] for r in cur.fetchall()]
        return [cols[i] for i in keep], rows

    def clean_reset(self, sid):
        with self.lock:
            for t in ("stg_" + sid, "_dropped_" + sid):
                self.conn.execute('DROP TABLE IF EXISTS "%s"' % t)
            self.conn.execute("UPDATE _sheets SET stg_table=NULL WHERE sid=?", [sid])
        return {"ok": True}

    def clean_save(self, sid, steps):
        with self.lock:
            self.conn.execute("UPDATE _sheets SET clean=? WHERE sid=?",
                              [json.dumps(steps or [], ensure_ascii=False), sid])
        return {"ok": True, "count": len(steps or [])}

    def clean_preset(self, sid, key):
        with self.lock:
            m = self._sheet(sid)
            cols = clean.columns_of(self.conn, m["table_name"])
            steps = clean.preset(key, cols, m["target"] or "tb")
        return {"steps": steps}

    def clean_dropped(self, sid, limit=200):
        with self.lock:
            m = self._sheet(sid)
            out = clean.dropped_rows(self.conn, sid, m["table_name"], limit)
            # 附带原始行内容（前 6 列），方便看为什么被剔除
            idx = {}
            cols = []
            if out["rows"]:
                try:
                    r0 = self.conn.execute('SELECT * FROM "%s" LIMIT 0' % m["table_name"])
                    cols = [d[0] for d in r0.description][:6]
                    ids = ",".join(str(int(x[0])) for x in out["rows"] if x[0] is not None)
                    if ids:
                        sel = ", ".join('"%s"' % c for c in cols)
                        for rec in self.conn.execute(
                                'SELECT _src_row, %s FROM "%s" WHERE _src_row IN (%s)'
                                % (sel, m["table_name"], ids)).fetchall():
                            idx[rec[0]] = [_s(v) for v in rec[1:]]
                except Exception:
                    pass
            out["cols"] = cols
            out["preview"] = [idx.get(r[0], []) for r in out["rows"]]
            return out

    # ------------------------------------------------------------ 辅助核算角色
    def detect_roles(self, sid, mapping=None, col=None):
        """给 aux1/2/3 猜角色：先看列名，再看样本值里的「客户:xxx」前缀。"""
        with self.lock:
            m = self._sheet(sid)
            tbl = m["stg_table"] or m["table_name"]
            mp = mapping if mapping is not None else m["mapping"]
            out = {}
            pairs = []
            if col:
                pairs = [("", col)]
            else:
                for slot in ("aux1", "aux2", "aux3"):
                    c = (mp or {}).get(slot)
                    if c:
                        pairs.append((slot, c))
                if not pairs:
                    for c in clean.columns_of(self.conn, tbl):
                        if role_from_header(c):
                            pairs.append(("", c))
            for slot, c in pairs:
                role = role_from_header(c)
                conf = 1.0 if role else 0.0
                if not role:
                    try:
                        vals = [r[0] for r in self.conn.execute(
                            'SELECT "%s" FROM "%s" WHERE nullif(trim(coalesce("%s", \'\')), \'\') '
                            'IS NOT NULL LIMIT 400' % (c, tbl, c)).fetchall()]
                    except Exception:
                        vals = []
                    role, conf = role_from_values(vals)
                    if conf < 0.25:
                        role = ""
                    elif conf < 0.7:
                        role = "mixed"      # 一格混了客户/部门/银行，交给前缀识别
                if role:
                    out[slot or c] = {"role": role, "conf": round(conf, 2), "col": c}
            return {"roles": out}

    def set_mapping(self, sid, target, mapping):
        with self.lock:
            self.conn.execute("UPDATE _sheets SET target=?, mapping=? WHERE sid=?",
                              [target, json.dumps(mapping or {}, ensure_ascii=False), sid])
        return {"ok": True}

    def suggest(self, sid, target):
        with self.lock:
            m = self._sheet(sid)
            tbl = m["stg_table"] or m["table_name"]
            try:
                cols = clean.columns_of(self.conn, tbl)
            except Exception:
                cols = m["columns"]
            mp = auto_map(target, cols)
            aux = {}
            if target == "tb":
                try:
                    r = self.detect_roles(sid, mapping=mp)
                    aux = {k: v["role"] for k, v in r["roles"].items()
                           if k in ("aux1", "aux2", "aux3")}
                except Exception:
                    aux = {}
            return {"mapping": mp, "aux": aux}

    def drop_sheet(self, sid):
        """删一张原始表（连带它的清洗结果表与剔除留痕表）。"""
        return self.drop_sheets([sid])

    def drop_sheets(self, sids):
        """批量删原始表：一次锁、一次事务，逐张清表。

        前端「已入库原始表」支持多选删除，逐张删会拿到 N 个中间态、
        也让「删到一半失败」变成半拉子；这里要么全删要么报错。
        """
        sids = [s for s in (sids or []) if s]
        done = []
        with self.lock:
            for sid in sids:
                r = self.conn.execute(
                    "SELECT table_name FROM _sheets WHERE sid=?", [sid]).fetchone()
                if not r:
                    continue
                for t in (r[0], "stg_" + sid, "_dropped_" + sid):
                    self.conn.execute('DROP TABLE IF EXISTS "%s"' % t)
                self.conn.execute("DELETE FROM _sheets WHERE sid=?", [sid])
                done.append(sid)
        return {"ok": True, "deleted": len(done), "sids": done}

    # ------------------------------------------------------------ merge
    def _expr(self, target, key, mapping, sname, cols=None):
        col = (mapping or {}).get(key)
        if not col:
            return None
        # 清洗步骤可能改了列（比如「恢复原始」后 科目级次 就不存在了），
        # 这种情况降级成 NULL，而不是让整次合并失败
        if cols is not None and col not in cols:
            return None
        ty = FTYPE.get((target, key), "text")
        c = qi(col)
        if ty == "date":
            return "_to_date(%s)" % c
        if ty == "amount":
            return "_num(%s)::DECIMAL(20,2)" % c
        if ty == "num":
            return "_num(%s)::DECIMAL(20,4)" % c
        return "_text(%s)" % c

    def _book_expr(self, meta):
        """账簿名称的取值来源。返回 SQL 表达式或 None（回落到普通列映射）。

        配置存在 mapping["_book"]：{mode: col|file|sheet|const, col, value}
        余额表/序时账常常没有「账簿名称」这一列，而是文件名、工作表名说了算，
        所以除了取列，还支持直接拿文件名 / 工作表名 / 手填值当账簿名称。
        """
        bk = (meta.get("mapping") or {}).get("_book") or {}
        mode = str(bk.get("mode") or "")
        if mode == "file":
            fn = os.path.splitext(os.path.basename(meta.get("file_name") or ""))[0]
            fn = _strip_ledger_tail(fn)
            return ql(fn.strip()) if fn.strip() else None
        if mode == "sheet":
            sn = str(meta.get("sheet_name") or "").strip()
            return ql(sn) if sn else None
        if mode == "const":
            v = str(bk.get("value") or "").strip()
            return ql(v) if v else None
        if mode == "col":
            c = bk.get("col") or (meta.get("mapping") or {}).get("book_name")
            cols = meta.get("cols")
            if c and (cols is None or c in cols):
                return "_text(%s)" % qi(c)
        return None

    def _company_expr(self, meta):
        """公司名称的取值来源。返回 SQL 表达式或 None（回落：映射列 / 账簿对照表）。

        配置存在 mapping["_company"]：{mode: books|col|book, col}
          books —— 用「账簿对照表」按账簿名称匹配（默认，不写配置就是这个）
          col   —— 直接取原表的某一列
          book  —— 等于「账簿名称」列（映射后的那份，含文件名/手填等来源）
        """
        cfg = (meta.get("mapping") or {}).get("_company") or {}
        mode = str(cfg.get("mode") or "")
        if mode == "book":
            # 账簿名称本身就是公司名（单体账套最常见），直接复用同一表达式
            return self._book_expr(meta)
        if mode == "col":
            c = cfg.get("col") or (meta.get("mapping") or {}).get("company_name")
            cols = meta.get("cols")
            if c and (cols is None or c in cols):
                return "_text(%s)" % qi(c)
        return None

    def _select_for(self, target, meta):
        mapping = meta["mapping"]
        sname = meta["file_name"] + " / " + meta["sheet_name"]
        table = meta.get("data_table") or meta["table_name"]
        cols = meta.get("cols")
        exprs = []
        for name, tp in CANON[target]:
            if name == "src":
                e = ql(sname)
            elif name == "src_row":
                e = "_src_row::BIGINT"
            elif name == "book_name":
                e = self._book_expr(meta) or self._expr(target, name, mapping, sname, cols)
            elif name == "company_name":
                e = self._company_expr(meta) or self._expr(target, name, mapping, sname, cols)
            else:
                e = self._expr(target, name, mapping, sname, cols)
            exprs.append((name, e))
        d = dict(exprs)

        # ---- 序时账：单列金额 + 方向 自动拆分借贷
        if target == "journal":
            amt = self._expr(target, "amount", mapping, sname, cols)
            if amt:
                dr = (mapping or {}).get("debit")
                cr = (mapping or {}).get("credit")
                dirn = self._expr(target, "direction", mapping, sname, cols)
                if not dr and not cr and dirn:
                    d["debit"] = "CASE WHEN %s LIKE '借%%' THEN %s END" % (dirn, amt)
                    d["credit"] = "CASE WHEN %s LIKE '贷%%' THEN %s END" % (dirn, amt)
                else:
                    if not dr:
                        d["debit"] = "CASE WHEN %s > 0 THEN %s END" % (amt, amt)
                    if not cr:
                        d["credit"] = "CASE WHEN %s < 0 THEN -%s END" % (amt, amt)

            # ---- 原币金额如果是单列（如「原币金额」），按借/贷哪边有值拆成 debit_fx / credit_fx
            amt_fx = self._expr(target, "amount_fx", mapping, sname, cols)
            if amt_fx and not mapping.get("debit_fx") and not mapping.get("credit_fx") \
                    and d.get("debit") and d.get("credit"):
                d["debit_fx"] = "CASE WHEN coalesce(%s, 0) <> 0 THEN %s END" % (d["debit"], amt_fx)
                d["credit_fx"] = "CASE WHEN coalesce(%s, 0) <> 0 THEN %s END" % (d["credit"], amt_fx)
            if not mapping.get("period") and mapping.get("vdate"):
                d["period"] = "strftime(_to_date(%s), '%%Y-%%m')" % qi(mapping["vdate"])
            elif mapping.get("period"):
                # 期间标准化：统一成 YYYY-MM（纯月份补零），从根上消除期间格式分歧
                pe = self._expr(target, "period", mapping, sname, cols)
                if pe:
                    d["period"] = "_period_norm(%s)" % pe
        # ---- 科目余额表：净额自动算
        if target == "tb":
            def _net(_b, _c, _n):
                """净额 = 借 − 贷。若映射指向的列在清洗后已不存在（_expr 返回 None），
                整体置 NULL，绝不拼出 COALESCE(None,0) 让整次合并 500。"""
                if mapping.get(_n):
                    return
                eb = self._expr(target, _b, mapping, sname, cols)
                ec = self._expr(target, _c, mapping, sname, cols)
                if eb and ec:
                    d[_n] = "COALESCE(%s,0) - COALESCE(%s,0)" % (eb, ec)
                else:
                    d[_n] = "NULL::DECIMAL(20,2)"
            _net("open_debit", "open_credit", "open_bal")
            _net("end_debit", "end_credit", "end_bal")
            _net("open_debit_fx", "open_credit_fx", "open_bal_fx")
            _net("end_debit_fx", "end_credit_fx", "end_bal_fx")
        # ---- 辅助核算语义列派生（序时账 / 科目余额表通用）
        if target in ("tb", "journal"):
            # 辅助核算：先按角色从位置槽抽语义列，再拼一个全文检索列
            roles = (mapping or {}).get("_aux") or {}
            slots = []
            for slot in ("aux1", "aux2", "aux3"):
                col = mapping.get(slot)
                if col and (cols is None or col in cols):
                    slots.append((slot, col, roles.get(slot) or ""))
            # 1) 明确指定了角色的槽：允许整列就是一个名字
            for _, col, role in slots:
                if role in ROLE_KEYS and not mapping.get(role):
                    d[role] = "_aux_pick(%s, %s, 1)" % (qi(col), ql(ROLE_KEYS[role]))
            # 2) 未指定角色（混合列）：只认「客户:xxx」这种带前缀的片段，按类型分发
            for _, col, role in slots:
                if role in ROLE_KEYS:
                    continue
                for field in ROLE_KEYS:
                    if mapping.get(field):
                        continue
                    expr = "_aux_pick(%s, %s, 0)" % (qi(col), ql(ROLE_KEYS[field]))
                    cur_e = d.get(field)
                    d[field] = expr if not cur_e else "COALESCE(%s, %s)" % (cur_e, expr)
            d["aux_text"] = "NULLIF(concat_ws('|', %s), '')" % ", ".join(
                "nullif(%s, '')" % (d.get(s) or "NULL::VARCHAR") for s in ("aux1", "aux2", "aux3"))

        # ---- 银行流水：金额双列合并、方向派生、账号归一
        if target == "flow":
            # 表头键值行 / 文件名给出的账户：整张表都是这一个账户，逐行补上
            if not d.get("acct_no") and mapping.get("_flow_acct"):
                d["acct_no"] = "%s::VARCHAR" % ql(str(mapping["_flow_acct"]))
                d["acct_source"] = ql(str(mapping.get("_flow_acct_src") or "表头/文件名"))
            if not d.get("acct_name") and mapping.get("_flow_acct_name"):
                d["acct_name"] = "%s::VARCHAR" % ql(str(mapping["_flow_acct_name"]))
            amt = self._expr(target, "amount", mapping, sname, cols)
            inc = self._expr(target, "income", mapping, sname, cols)
            exp = self._expr(target, "expense", mapping, sname, cols)
            # 单列金额优先；没有单列就用「收入 − 支出」合成（银行导出常见的两列式）
            if not amt and (inc or exp):
                amt = "COALESCE(%s, 0) - COALESCE(%s, 0)" % (inc or "0", exp or "0")
                d["amount"] = amt
            # 方向：配了方向列就把它归一到「收/付」，没配就按金额正负推
            dcol = mapping.get("direction")
            if dcol and (cols is None or dcol in cols):
                de = d.get("direction")
                d["direction"] = (
                    "CASE WHEN %s IS NULL OR trim(%s::VARCHAR) = '' THEN NULL "
                    "WHEN trim(%s::VARCHAR) IN ('收','收入','贷','贷方','进','存入',"
                    "'C','CR','credit','Credit') THEN '收' "
                    "WHEN trim(%s::VARCHAR) IN ('付','支出','借','借方','出','支取',"
                    "'D','DR','debit','Debit') THEN '付' "
                    "ELSE trim(%s::VARCHAR) END" % (de, de, de, de, de))
            elif amt:
                d["direction"] = ("CASE WHEN %s > 0 THEN '收' WHEN %s < 0 THEN '付' END"
                                  % (amt, amt))
            # 账户配对用的归一键（原始账号保留在 acct_no / counterparty_acct 里可追溯）
            for _src, _dst in (("acct_no", "acct_key"), ("counterparty_acct", "cp_key")):
                _e = d.get(_src)
                if _e:
                    d[_dst] = _acct_norm(_e)

        cols = []
        for name, tp in CANON[target]:
            e = d.get(name) or ("NULL::%s" % tp)
            cols.append("%s AS %s" % (e, qi(name)))
        return "SELECT %s FROM %s" % (", ".join(cols), qi(table))

    def merge(self, target, force=False):
        """合并成标准表。

        **清洗达标门禁**：tb / journal 合并后会立刻跑 audit（scope=pre_merge），
        有 fail 级不达标项就回滚到合并前并抛 MergeBlocked——防止带着合计行、父科目、
        空一级科目这样的脏数据往下走（验收方就是这么把「假不符」当成了平台缺陷）。
        force=True 可越过，但不达标项全部写进 warnings 留痕。
        """
        if target not in CANON:
            raise ValueError("未知目标：%s" % target)
        with self.lock:
            rows = self.conn.execute(
                "SELECT table_name,file_name,sheet_name,mapping,n_rows,stg_table FROM _sheets "
                "WHERE target=? ORDER BY file_name,sheet_name", [target]).fetchall()
            metas = [{"table_name": r[0], "file_name": r[1], "sheet_name": r[2],
                      "mapping": json.loads(r[3] or "{}"), "n_rows": r[4],
                      "data_table": r[5] or r[0]} for r in rows]
            if not metas:
                raise ValueError("还没有任何工作表映射到「%s」" % TARGETS[target]["label"])
            warns = []
            for m in metas:
                try:
                    m["cols"] = set(clean.columns_of(self.conn, m["data_table"]))
                except Exception:
                    m["cols"] = set()
                miss = sorted({str(v) for k, v in (m["mapping"] or {}).items()
                               if not str(k).startswith("_") and v
                               and str(v) not in m["cols"]})
                if miss:
                    warns.append("%s / %s：列 %s 不在当前数据表里（清洗后可能没了），已置空"
                                 % (m["file_name"], m["sheet_name"], "、".join(miss)))
            t0 = time.time()
            sql = " UNION ALL ".join(self._select_for(target, m) for m in metas)
            self.conn.execute('CREATE OR REPLACE TABLE "%s_tmp" AS %s' % (target, sql))
            # 旧表先备份成 _prev：门禁不通过要原样回滚，不能把用户上一版标准表弄丢
            had_old = target in [r[0] for r in self.conn.execute("SHOW TABLES").fetchall()]
            if had_old:
                self.conn.execute('DROP TABLE IF EXISTS "%s_prev"' % target)
                self.conn.execute('ALTER TABLE "%s" RENAME TO "%s_prev"' % (target, target))
            self.conn.execute('ALTER TABLE "%s_tmp" RENAME TO "%s"' % (target, target))
            n = self.conn.execute('SELECT count(*) FROM "%s"' % target).fetchone()[0]
            # 先回填公司/账簿名称：门禁里的「必备字段」要按回填后的结果判，
            # 否则账簿对照刚起作用的项目会被误拦
            # 先按名称推断「对照表没覆盖的账簿名」对应哪家公司（写 _book_match 留痕），
            # 再由 _apply_books 回填——顺序不能反，否则推出来的结果用不上。
            try:
                self.infer_book_matches(target)
            except Exception:
                pass
            book = self._apply_books(target)
            # ---- 清洗达标门禁 ----
            gate = self._clean_gate(target)
            if gate and not force:
                self.conn.execute('DROP TABLE IF EXISTS "%s"' % target)
                if had_old:
                    self.conn.execute('ALTER TABLE "%s_prev" RENAME TO "%s"' % (target, target))
                raise MergeBlocked(target, gate)
            forced = bool(gate)
            if forced:
                for i in gate:
                    warns.append("清洗未达标（已强制合并）：%s —— %s"
                                 % (i.get("title") or "", (i.get("detail") or "")[:120]))
            if had_old:
                self.conn.execute('DROP TABLE IF EXISTS "%s_prev"' % target)
            # 必填字段检查：缺了这些，后面的核对会报错或失真，提前提示
            missing_required = []
            try:
                have = set(clean.columns_of(self.conn, target))
            except Exception:
                have = set()
            if target == "journal":
                if "acct_l1" not in have:
                    missing_required.append("一级科目（第三组核对需要）")
                if "debit" not in have and "credit" not in have:
                    missing_required.append("借方/贷方金额（第二组核对需要）")
            elif target == "tb":
                if "acct_l1" not in have:
                    missing_required.append("一级科目（第三组核对需要）")
                if not any(c in have for c in ("ytd_debit", "cur_debit", "open_debit", "end_debit")):
                    missing_required.append("发生额口径（本年累计/本期 借方贷方，第三四组核对需要）")
            elif target == "flow":
                if not any(c in have for c in ("acct_no", "acct_key")):
                    missing_required.append("本方账号（不知道流水属于哪个账户，无法与账面配对）")
                if not any(c in have for c in ("amount", "income", "expense")):
                    missing_required.append("交易金额（收入/支出两列或单列金额至少要有一种）")
                if "txn_date" not in have:
                    missing_required.append("交易日期（截止性测试按日期取区间）")
                if "counterparty_acct" not in have and "counterparty_name" not in have:
                    missing_required.append("对方账号/对方户名（缺了就判断不了是否内部转账）")
                # 长账号被 Excel 读成科学计数法：不拦下来会静默漏配全部账户
                try:
                    bad = self.conn.execute(
                        "SELECT count(*) FROM flow WHERE coalesce(acct_key,'') ~* 'e\\+[0-9]+$'"
                        " OR coalesce(cp_key,'') ~* 'e\\+[0-9]+$'").fetchone()[0]
                    if bad:
                        warns.append("有 %d 行账号看起来是科学计数法（如 7.55933E+14），"
                                     "说明 Excel 把账号存成了数字。请把账号列改成文本格式后重新导入，"
                                     "否则账户配对会大面积失败。" % bad)
                except Exception:
                    pass
            if missing_required:
                warns.append("缺必填字段：%s" % "、".join(missing_required))
            # 银行流水：合并后立刻和账户清单对一次，让"这笔是谁的、对方是不是自己人"当场可见。
            # 清单之后补全了也能单独重跑（/api/flow/link），不用重导流水。
            link_info = None
            if target == "flow":
                from . import accounts as ACC
                try:
                    built = ACC.build(self.conn)
                    if built.get("error"):
                        warns.append(built["error"])
                    else:
                        link_info = ACC.link(self.conn, only_missing=False)
                        if link_info.get("error"):
                            warns.append(link_info["error"])
                            link_info = None
                        elif link_info.get("acctLinked", 0) == 0:
                            # 两种原因的处理动作完全不同，不能给一句含糊的提示
                            if not link_info.get("okAccounts"):
                                warns.append(
                                    "账户清单里 %d 个账户都没有账号，流水无法与账面对应。"
                                    "请到「清洗与映射」把银行账号映射到辅助核算"
                                    "（aux 槽或 bank_account 列），再点「重建账户清单」。"
                                    % link_info.get("accounts", 0))
                            else:
                                warns.append(
                                    "有 %d 行流水的本方账号没配上账户清单"
                                    "（清单里有 %d 个可用账户）。多半是账号格式不一致，"
                                    "或这几户账面上没有账号。"
                                    % (link_info.get("rows", 0),
                                       link_info.get("okAccounts", 0)))
                except Exception as e:
                    warns.append("账户配对失败（流水数据本身已合并成功）：%s" % str(e)[:160])
        return {"target": target, "rows": n, "sheets": len(metas),
                "elapsed": round(time.time() - t0, 2), "warnings": warns,
                "missing_required": missing_required, "books": book,
                "forced": forced,
                "link": link_info, "accounts": built.get("total", 0) if target == "flow"
                and link_info is not None else 0}

    def _clean_gate(self, target):
        """合并结果的清洗达标门禁：返回 fail 级不达标项，空列表=通过。

        只对 tb / journal 生效（flow/fa 的检查在各自流程里）。判的是**刚合并出来的
        这张表**，复用 audit 的同一套口径，不再维护第二份判定逻辑。
        检查本身出错不阻断（audit 内部会把异常记成 fail 项，这里再兜一层，避免
        因为体检代码崩了导致业务卡死）。
        """
        if target not in ("tb", "journal"):
            return []
        try:
            from . import audit as AUD
            rep = AUD.audit(self.conn, scope="pre_merge")
        except Exception:
            return []
        group = "科目余额表" if target == "tb" else "序时账"
        # 只拦 block=True 的 fail（会导致金额翻倍/假不符的硬问题）。
        # 辅助核算缺失这类「客户数据本身不完整」不阻断——它影响底稿明细颗粒度，
        # 不影响两表对账，拦下来只会让人无法推进（真实验收账套就是这种情况）。
        return [i for i in rep.get("items") or []
                if i.get("group") == group and not i.get("ok")
                and i.get("level") == "fail" and i.get("block", True)]

    def infer_book_matches(self, target="tb"):
        """把「表里出现、对照表没有」的账簿名，推断到最可能的公司并留痕。

        对照表（_books）就是完整主体清单。账簿名对不上时不能空着——空着在合并后
        company_name 是 NULL，主体完整性检查会把它判成「超范围」，底稿分主体也会
        把几个主体加在一起比，错得还很安静。

        留痕写 _book_match：一条账簿名一行，记方法 / 分数 / 中文依据 / 是否需确认。
        src='manual' 的行（人工改过的）不会被自动推断覆盖。

        方法由强到弱：norm（归一后完全相同）→ norm_prefix（去掉门店前缀相同）→
        contain（互相包含）→ prefix（共同开头 ≥4 字）→ similar（文字相似度）。
        除 norm 外都要人工确认——「（南海店）」和「（东湖店）」去掉前缀是同一串字，
        但那是两家店，不能自动合并。
        """
        q = qi(target)
        try:
            names = [r[0] for r in self.conn.execute(
                "SELECT DISTINCT trim(book_name) FROM %s "
                "WHERE nullif(trim(book_name), '') IS NOT NULL ORDER BY 1" % q).fetchall()]
        except Exception:
            return None
        if not names:
            return {"names": 0, "matched": 0, "inferred": 0, "need_confirm": 0}
        try:
            books = self.conn.execute(
                "SELECT trim(book_name), trim(company_name) FROM _books "
                "WHERE nullif(trim(coalesce(company_name,'')), '') IS NOT NULL").fetchall()
        except Exception:
            return None
        if not books:
            return None

        # 候选池：对照表的「账簿名」和「公司名」都当比对对象，都指向同一个公司
        cands = []
        for bn, cn in books:
            for src_name in {bn, cn}:
                cands.append((_book_key(src_name), _book_key(src_name, True), cn, bn))


        try:
            manual = {r[0] for r in self.conn.execute(
                "SELECT book_name FROM _book_match WHERE src='manual'").fetchall()}
        except Exception:
            manual = set()

        n_ok = n_inf = n_need = 0
        for name in names:
            if name in manual:
                continue
            sc, method, cn, ev, need = _pick_book(name, cands)
            if method == "norm":
                n_ok += 1
            elif method != "none":
                n_inf += 1
            if need:
                n_need += 1
            self.conn.execute("DELETE FROM _book_match WHERE book_name = ?", [name])
            self.conn.execute(
                "INSERT INTO _book_match(book_name, company_name, method, score, evidence, "
                "need_confirm, src, updated_at) VALUES (?,?,?,?,?,?, 'auto', ?)",
                [name, cn, method, round(float(sc), 3), ev, int(need), datetime.now()])
        return {"names": len(names), "matched": n_ok, "inferred": n_inf,
                "need_confirm": n_need}

    def _apply_books(self, target):
        """按账簿名称回填公司名称 / 简称。

        两张表都调用一次，源数据里已经有公司名称的行不覆盖——人工填的优先。
        返回匹配情况，让前端能提示「有 3 个账簿名称在对照表里找不到」。
        """
        try:
            n_books = self.conn.execute("SELECT count(*) FROM _books").fetchone()[0]
        except Exception:
            return None
        if not n_books:
            return None
        q = qi(target)
        try:
            n_rows = self.conn.execute(
                'SELECT count(*) FROM %s WHERE nullif(trim(book_name), \'\') IS NOT NULL' % q
            ).fetchone()[0]
        except Exception:
            return None
        if not n_rows:
            return {"books": n_books, "rows_with_book": 0, "matched": 0,
                    "unmatched": [], "note": "表里没有账簿名称，未回填公司名称/简称"}
        matched = self.conn.execute(
            'SELECT count(*) FROM %s t JOIN _books b ON _book_norm(t.book_name) = _book_norm(b.book_name)' % q
        ).fetchone()[0]
        if matched:
            try:
                have = set(clean.columns_of(self.conn, target))
            except Exception:
                have = set()
            sets = []
            if "company_name" in have:      # 源数据里已经填了公司名称的不覆盖
                sets.append("company_name = COALESCE(NULLIF(trim(t.company_name), ''), "
                            "b.company_name)")
            if "short_name" in have:
                sets.append("short_name = b.short_name")
            if sets:
                self.conn.execute(
                    'UPDATE %s AS t SET %s FROM _books AS b '
                    'WHERE _book_norm(t.book_name) = _book_norm(b.book_name)' % (q, ", ".join(sets)))
        # 推断兜底：对照表没覆盖的账簿名，用 _book_match 里推断出来的公司名补上。
        # 只补 company_name 为空的行——源数据里已经填的、对照表命中的都不动。
        filled_infer = 0
        try:
            have = set(clean.columns_of(self.conn, target))
        except Exception:
            have = set()
        if "company_name" in have:
            try:
                before = self.conn.execute(
                    "SELECT count(*) FROM %s WHERE nullif(trim(coalesce(company_name,'')),'') IS NULL" % q
                ).fetchone()[0]
                self.conn.execute(
                    'UPDATE %s AS t SET company_name = m.company_name FROM _book_match m '
                    "WHERE nullif(trim(t.company_name), '') IS NULL "
                    "AND nullif(trim(coalesce(m.company_name,'')), '') IS NOT NULL "
                    "AND _book_norm(t.book_name) = _book_norm(m.book_name)" % q)
                after = self.conn.execute(
                    "SELECT count(*) FROM %s WHERE nullif(trim(coalesce(company_name,'')),'') IS NULL" % q
                ).fetchone()[0]
                filled_infer = max(0, before - after)
            except Exception:
                pass
        try:
            un = [r[0] for r in self.conn.execute(
                'SELECT DISTINCT trim(t.book_name) FROM %s t WHERE '
                'nullif(trim(t.book_name), \'\') IS NOT NULL AND NOT EXISTS '
                '(SELECT 1 FROM _books b WHERE _book_norm(b.book_name) = _book_norm(t.book_name)) '
                'ORDER BY 1 LIMIT 20' % q).fetchall()]
        except Exception:
            un = []
        try:
            still_empty = self.conn.execute(
                "SELECT count(*) FROM %s WHERE nullif(trim(coalesce(company_name,'')),'') IS NULL" % q
            ).fetchone()[0]
        except Exception:
            still_empty = 0
        return {"books": n_books, "rows_with_book": n_rows, "matched": matched,
                "unmatched": un, "filled_by_infer": filled_infer,
                "still_empty": still_empty}

    # ------------------------------------------------------------ query
    def schema(self):
        """返回最终合并库的表结构信息，供查账 SQL 界面展示字段。

        只列 journal / tb 两张标准表（_recon 核对结果表核对后才有，单独给）。
        每列带英文名、中文名、类型，方便直接丢给 AI 写 SQL。
        """
        out = {}
        for t in ("journal", "tb"):
            try:
                n = self.conn.execute('SELECT count(*) FROM "%s"' % t).fetchone()[0]
            except Exception:
                n = 0
            cols = [{"key": k, "label": to_cn(k, t), "type": tp} for k, tp in CANON[t]]
            out[t] = {"label": TARGETS[t]["label"], "rows": n, "columns": cols}
        # 核对结果表（存在才给）
        try:
            self.conn.execute("SELECT * FROM _recon LIMIT 0")
            rcols = [d[0] for d in self.conn.execute(
                "SELECT * FROM _recon LIMIT 0").description]
            out["_recon"] = {"label": "核对结果", "rows": 0,
                             "columns": [{"key": c, "label": RECON_LABEL.get(c, c),
                                          "type": "VARCHAR"} for c in rcols]}
        except Exception:
            pass
        return out

    def interrupt(self):
        """中断当前正在执行的查询。

        请求线程正在阻塞执行 SQL 时，其它线程调本方法可取消它。
        DuckDB 的 interrupt() 是线程安全的，专门用于这种场景。
        """
        try:
            self.conn.interrupt()
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)[:200]}

    def query(self, sql, limit=2000, params=None):
        sql = _subst(sql, params)
        sql = sql.strip().rstrip(";").strip()
        if not sql:
            raise ValueError("SQL 为空")
        t0 = time.time()
        try:
            with self.lock:
                cur = self.conn.execute("SELECT * FROM (%s) __q LIMIT %d" % (sql, int(limit) + 1))
                cols = [d[0] for d in cur.description]
                rows = cur.fetchall()
        except Exception as e:
            if "Interrupted" in str(e) or "INTERRUPT" in str(e).upper():
                raise ValueError("查询已被中断") from None
            raise
        ms = int((time.time() - t0) * 1000)
        more = len(rows) > limit
        rows = rows[:limit]
        return {"columns": cols, "rows": [[_num(v) for v in r] for r in rows],
                "truncated": more, "elapsed_ms": ms, "sql": sql}

    def count(self, sql, params=None):
        sql = _subst(sql, params).strip().rstrip(";")
        t0 = time.time()
        try:
            with self.lock:
                n = self.conn.execute("SELECT count(*) FROM (%s) __c" % sql).fetchone()[0]
        except Exception as e:
            if "Interrupted" in str(e) or "INTERRUPT" in str(e).upper():
                raise ValueError("查询已被中断") from None
            raise
        return {"rows": n, "elapsed_ms": int((time.time() - t0) * 1000)}

    def explain(self, sql, params=None):
        sql = _subst(sql, params).strip().rstrip(";")
        with self.lock:
            rows = self.conn.execute("EXPLAIN " + sql).fetchall()
        return "\n".join(str(r[1]) for r in rows)

    # ------------------------------------------------------------ export
    def export(self, sql, fmt="csv", name=None, params=None):
        """导出查询结果。

        库里的列名是英文标准名（acct_code 等），导出文件一律换成中文表头。
        做法是让 DuckDB 只写数据不写表头，我们自己拼第一行——大表也不用把结果
        读进内存。
        """
        sql = _subst(sql, params).strip().rstrip(";")
        stem = (name or ("查询结果_" + datetime.now().strftime("%Y%m%d_%H%M%S")))
        stem = re.sub(r'[\\/:*?"<>|]', "_", stem)
        try:
            with self.lock:
                cur = self.conn.execute("SELECT * FROM (%s) __e LIMIT 0" % sql)
                headers = cn_headers([d[0] for d in cur.description],
                                     _sniff_target(sql))
                if fmt == "xlsx":
                    path = os.path.join(self.out_dir, stem + ".xlsx")
                    self._to_xlsx(sql, path, headers)
                else:
                    path = os.path.join(self.out_dir, stem + ".csv")
                    tmp = path + ".raw"
                    self.conn.execute(
                        "COPY (%s) TO %s (FORMAT CSV, HEADER false, DELIMITER ',')"
                        % (sql, ql(tmp)))
                    buf = io.StringIO()
                    csv.writer(buf, lineterminator="\n").writerow(headers)
                    with open(tmp, "rb") as fi, open(path, "wb") as fo:
                        fo.write(b"\xef\xbb\xbf")
                        fo.write(buf.getvalue().encode("utf-8"))
                        shutil.copyfileobj(fi, fo, 8 << 20)
                    os.remove(tmp)
        except Exception as e:
            if "Interrupted" in str(e) or "INTERRUPT" in str(e).upper():
                raise ValueError("查询已被中断") from None
            raise
        return {"path": path, "file": os.path.basename(path),
                "size": os.path.getsize(path), "columns": headers}

    def _to_xlsx(self, sql, path, headers=None):
        from openpyxl import Workbook
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("结果")
        cur = self.conn.execute(sql)
        ws.append(headers or [d[0] for d in cur.description])
        while True:
            batch = cur.fetchmany(20000)
            if not batch:
                break
            for r in batch:
                ws.append([_num(v) for v in r])
        wb.save(path)

    def share_export(self, tables=("journal", "tb"), note=""):
        st = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = "看账数据包_" + st
        tmp = os.path.join(self.work_dir, "share_" + st)
        os.makedirs(tmp, exist_ok=True)
        manifest = {"format": "kanzhang/v2", "created_at": datetime.now().isoformat(timespec="seconds"),
                    "note": note, "tables": {}}
        with self.lock:
            for t in tables:
                n = self.conn.execute('SELECT count(*) FROM "%s"' % t).fetchone()[0]
                if n == 0:
                    continue
                pq = os.path.join(tmp, t + ".parquet")
                self.conn.execute('COPY (SELECT * FROM "%s") TO %s (FORMAT PARQUET, COMPRESSION zstd)'
                                  % (t, ql(pq)))
                manifest["tables"][t] = {
                    "rows": n, "cols": [c[0] for c in CANON[t]],
                    "label": TARGETS[t]["label"],
                    "size": os.path.getsize(pq)}
            rules = self.get_rules()
            # 数据包不能只装结果：对方拿到一堆洗好的表，既看不懂口径、也没法复现。
            # 把「每张原始表的清洗动作 + 字段映射 + 目标表」和「账簿对照表」一起打进去，
            # 他打开就知道这张表从哪来、按什么规则洗的、每个字段对的是哪一列。
            try:
                srows = self.conn.execute(
                    "SELECT file_name, sheet_name, header_row, header_rows, target, "
                    "mapping, clean, col_names, n_rows, stg_table IS NOT NULL FROM _sheets "
                    "ORDER BY file_name, sheet_name").fetchall()
                sheets = [{
                    "file_name": r[0], "sheet_name": r[1], "header_row": r[2],
                    "header_rows": r[3], "target": r[4],
                    "target_label": (TARGETS.get(r[4]) or {}).get("label", ""),
                    "mapping": json.loads(r[5] or "{}"),
                    "clean": json.loads(r[6] or "[]"),
                    "columns": json.loads(r[7] or "[]"),
                    "n_rows": r[8], "cleaned": bool(r[9]),
                } for r in srows]
            except Exception:
                sheets = []
            try:
                books = self.books_list(limit=1000000)
            except Exception:
                books = {"columns": [], "rows": [], "total": 0}
        manifest["sheetCount"] = len(sheets)
        manifest["bookCount"] = books.get("total", 0)
        with open(os.path.join(tmp, "manifest.json"), "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        with open(os.path.join(tmp, "rules.json"), "w", encoding="utf-8") as f:
            json.dump(rules, f, ensure_ascii=False, indent=2)
        # sheets.json：清洗规则 + 字段映射（每张原始表一条）
        with open(os.path.join(tmp, "sheets.json"), "w", encoding="utf-8") as f:
            json.dump(sheets, f, ensure_ascii=False, indent=1)
        # books.json：账簿 → 公司名称/简称 对照表
        with open(os.path.join(tmp, "books.json"), "w", encoding="utf-8") as f:
            json.dump(books, f, ensure_ascii=False, indent=1)
        zpath = os.path.join(self.out_dir, base + ".zip")
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            for fn in os.listdir(tmp):
                z.write(os.path.join(tmp, fn), fn)
        shutil.rmtree(tmp, ignore_errors=True)
        return {"path": zpath, "file": os.path.basename(zpath),
                "size": os.path.getsize(zpath), "manifest": manifest}

    def share_import(self, fid, mode="replace"):
        f = self._file(fid)
        tmp = os.path.join(self.work_dir, "imp_" + uuid.uuid4().hex[:8])
        os.makedirs(tmp, exist_ok=True)
        skipped = []
        try:
            # 逐成员解包，不用 extractall：包内某个文件 CRC 坏、或被单独加了密码，
            # 不该让整包失败丢掉其余能用的表。
            with zipfile.ZipFile(f["path"]) as z:
                root = os.path.normpath(tmp) + os.sep
                for i in z.infolist():
                    if i.is_dir():
                        continue
                    name = (i.filename or "").replace("\\", "/")
                    dst = os.path.normpath(os.path.join(tmp, name))
                    if name.startswith("/") or not dst.startswith(root):
                        skipped.append({"name": i.filename, "why": "路径非法，已跳过"})
                        continue
                    try:
                        os.makedirs(os.path.dirname(dst), exist_ok=True)
                        with z.open(i) as src, open(dst, "wb") as fo:
                            shutil.copyfileobj(src, fo, 1 << 20)
                    except Exception as e:
                        skipped.append({"name": i.filename, "why": str(e)[:80]})
            mf = os.path.join(tmp, "manifest.json")
            if not os.path.exists(mf):
                raise ValueError("不是有效的数据包（缺少 manifest.json）"
                                 + ("；包内有 %d 个文件读取失败" % len(skipped) if skipped else ""))
            manifest = json.load(open(mf, encoding="utf-8"))
            done = {}
            with self.lock:
                for t, info in manifest.get("tables", {}).items():
                    if t not in CANON:
                        continue
                    pq = os.path.join(tmp, t + ".parquet")
                    if not os.path.exists(pq):
                        continue
                    cols = [c for c, _ in CANON[t]]
                    # 以 parquet 里真实存在的列为准（manifest 只是参考）：
                    # 老版本的数据包没有新加的列，缺的补 NULL，不至于整包导不进来
                    try:
                        actual = {r[0] for r in self.conn.execute(
                            "DESCRIBE SELECT * FROM read_parquet(%s)" % ql(pq)).fetchall()}
                    except Exception:
                        actual = set(info.get("cols") or [])
                    if not actual:
                        actual = set(info.get("cols") or [])
                    sel = ", ".join(('"%s"' % c) if c in actual else
                                    ("NULL::%s AS \"%s\"" % (dict(CANON[t])[c], c))
                                    for c in cols)
                    if mode == "append":
                        self.conn.execute(
                            'INSERT INTO "%s" SELECT %s FROM read_parquet(%s)' % (t, sel, ql(pq)))
                    else:
                        self.conn.execute(
                            'CREATE OR REPLACE TABLE "%s_tmp" AS SELECT %s FROM read_parquet(%s)'
                            % (t, sel, ql(pq)))
                        self.conn.execute('DROP TABLE IF EXISTS "%s"' % t)
                        self.conn.execute('ALTER TABLE "%s_tmp" RENAME TO "%s"' % (t, t))
                    done[t] = self.conn.execute('SELECT count(*) FROM "%s"' % t).fetchone()[0]
            rj = os.path.join(tmp, "rules.json")
            if os.path.exists(rj) and mode == "replace":
                try:
                    self.save_rules(json.load(open(rj, encoding="utf-8")))
                except Exception:
                    pass
            # 账簿对照表：整表替换，这是「公司 mapping」，后面 BKD 的简称全靠它
            extra = {"books": 0, "sheetConfigs": 0}
            bj = os.path.join(tmp, "books.json")
            if os.path.exists(bj):
                try:
                    bk = json.load(open(bj, encoding="utf-8"))
                    rows_ = bk.get("rows") or []
                    with self.lock:
                        if rows_ and mode == "replace":
                            self.conn.execute("DELETE FROM _books")
                        for r in rows_:
                            self.conn.execute(
                                "INSERT OR REPLACE INTO _books "
                                "(book_name, company_name, short_name, src, imported_at) "
                                "VALUES (?,?,?,?, now())",
                                [r[0], r[1] if len(r) > 1 else None,
                                 r[2] if len(r) > 2 else None,
                                 (r[3] if len(r) > 3 else None) or "数据包"])
                            extra["books"] += 1
                except Exception:
                    pass
            # 清洗动作 + 字段映射：按「文件名 + 工作表名」回挂到本地已有的原始表。
            # 本机没有那张表就跳过——它只是配置，没有数据挂上去也没意义。
            sj = os.path.join(tmp, "sheets.json")
            if os.path.exists(sj):
                try:
                    cfg = json.load(open(sj, encoding="utf-8"))
                    with self.lock:
                        for c in cfg or []:
                            r = self.conn.execute(
                                "SELECT sid FROM _sheets WHERE file_name=? AND sheet_name=?",
                                [c.get("file_name"), c.get("sheet_name")]).fetchone()
                            if not r:
                                continue
                            self.conn.execute(
                                "UPDATE _sheets SET mapping=?, clean=?, target=coalesce(target, ?) "
                                "WHERE sid=?",
                                [json.dumps(c.get("mapping") or {}, ensure_ascii=False),
                                 json.dumps(c.get("clean") or [], ensure_ascii=False),
                                 c.get("target"), r[0]])
                            extra["sheetConfigs"] += 1
                except Exception:
                    pass
            return {"tables": done, "manifest": manifest, "config": extra,
                    "skipped": skipped}
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    # ------------------------------------------------------------ 数据检查
    def recon(self, period_from=None, period_to=None, tol=0.01, tb_base="ytd"):
        """四组平衡检查：余额表自平衡 / 序时账自平衡 / 两表一级科目 / 两表科目全称。

        结果同时落成 _recon_* 表，并把参数与清单持久化到 _recon_last，
        刷新页面、重启后都能恢复上一次结果（导出也不依赖进程内存）。
        """
        with self.lock:
            r = recon_mod.run(self.conn, period_from, period_to, tol, tb_base)
            self._recon_tables = r.get("tables") or []
            self._recon_meta = {"summary": r.get("summary") or {},
                                "tol": r.get("tol"), "period": r.get("period"),
                                "tb_label": r.get("tb_label")}
            try:
                self.conn.execute(
                    "CREATE TABLE IF NOT EXISTS _recon_last("
                    "ran_at TIMESTAMP, params VARCHAR, summary VARCHAR, tables VARCHAR)")
                self.conn.execute("DELETE FROM _recon_last")
                self.conn.execute(
                    "INSERT INTO _recon_last VALUES (?,?,?,?)",
                    [datetime.now(),
                     json.dumps({"period": r.get("period"), "tol": r.get("tol"),
                                 "tb_base": r.get("tb_base"), "tb_label": r.get("tb_label")},
                                ensure_ascii=False),
                     json.dumps(r.get("summary") or {}, ensure_ascii=False),
                     json.dumps(r.get("tables") or [], ensure_ascii=False)])
            except Exception:
                pass
            return r

    def recon_last(self):
        """上一次检查的参数与结论（刷新页面 / 重启后恢复用）。"""
        try:
            row = self.conn.execute(
                "SELECT ran_at, params, summary, tables FROM _recon_last LIMIT 1").fetchone()
            if not row:
                return None
            return {"ran_at": str(row[0]), "params": json.loads(row[1] or "{}"),
                    "summary": json.loads(row[2] or "{}"),
                    "tables": json.loads(row[3] or "[]")}
        except Exception:
            return None

    def diagnose(self):
        """数据体检（只读）：对 tb / journal 跑一组常见检查，返回发现 + 建议清洗步骤。

        覆盖：币别长表、重复账簿层、汇总层、小计/合计行、页脚、整百万行、期间错位。
        替代「人工写 SQL 诊断」，让非技术用户也能自己判断口径。
        """
        findings = []

        def q(sql, default=None):
            try:
                r = self.conn.execute(sql).fetchone()
                return r[0] if r else default
            except Exception:
                return default

        def rows(sql):
            try:
                return self.conn.execute(sql).fetchall()
            except Exception:
                return []

        has_tb = recon_mod._has_table(self.conn, "tb")
        has_jr = recon_mod._has_table(self.conn, "journal")

        # 1) 币别长表（正常，但要提醒）
        if has_tb:
            n_mixed = q("""SELECT count(*) FROM (
                SELECT coalesce(nullif(trim(company_name),''), nullif(trim(book_name),'')) e, acct_code
                FROM tb WHERE nullif(trim(currency),'') IS NOT NULL GROUP BY 1,2
                HAVING count(*) FILTER (WHERE currency='综合本位币')>0
                   AND count(*) FILTER (WHERE currency<>'综合本位币')>0)""", 0)
            if n_mixed:
                findings.append({"level": "info", "title": "币别长表（正常）",
                                 "detail": "%d 个「主体×科目」同时有综合本位币行和外币行，"
                                           "核对已按币别分组，不会重复，无需处理。" % n_mixed,
                                 "suggest": None})

        # 2) 重复账簿层
        if has_jr:
            n_dup = q("""SELECT count(*) FROM journal WHERE nullif(trim(book_name),'') IS NOT NULL
                         AND regexp_matches(trim(book_name), '^GB-|副账簿|^合计$')""", 0)
            if n_dup:
                findings.append({"level": "warn", "title": "疑似重复账簿层",
                                 "detail": "序时账里有 %d 行账簿名匹配「GB- / 副账簿 / 合计」，"
                                           "通常与主账簿重叠，全额导入会让发生额翻倍。" % n_dup,
                                 "suggest": ["加一步「按关键词删行」：列=账簿名称，关键词=GB-|副账簿|合计，方式=正则"]})

        # 3) 汇总层：某主体金额 ≈ 其它主体之和
        if has_tb:
            try:
                agg = rows("""SELECT coalesce(nullif(trim(company_name),''), nullif(trim(book_name),'')) e,
                    sum(coalesce(cur_debit,0)) d FROM tb GROUP BY 1 HAVING e IS NOT NULL AND count(*)>0""")
                tot = {r[0]: (r[1] or 0) for r in agg}
                for e, d in tot.items():
                    others = sum(v for k, v in tot.items() if k != e)
                    if others and abs(d - others) < 0.01:
                        findings.append({"level": "warn", "title": "疑似汇总层",
                                         "detail": "主体「%s」的金额 %.2f ≈ 其余主体之和 %.2f，"
                                                   "可能是汇总行，与明细重复。" % (e, d, others),
                                         "suggest": ["加一步「按关键词删行」删掉这个汇总主体行"]})
                        break
            except Exception:
                pass

        # 4) 小计/合计行（只在科目名/辅助核算维度里找整格「合计/小计/总计」，
        #    摘要里带「合计」是正常文字，不算）
        n_sub = 0
        if has_tb:
            n_sub = q("""SELECT count(*) FROM tb WHERE regexp_matches(
                concat_ws('|', nullif(trim(acct_name), ''), nullif(trim(aux1), ''),
                          nullif(trim(aux2), ''), nullif(trim(aux3), '')),
                '(^|\\|)(合计|小计|总计|本页合计)(\\||$)')""", 0)
        if not n_sub and has_jr:
            n_sub = q("""SELECT count(*) FROM journal WHERE regexp_matches(
                concat_ws('|', nullif(trim(acct_name), ''), nullif(trim(aux1), ''),
                          nullif(trim(aux2), ''), nullif(trim(aux3), '')),
                '(^|\\|)(合计|小计|总计)(\\||$)')""", 0)
        if n_sub:
            findings.append({"level": "warn", "title": "疑似小计/合计行",
                             "detail": "表里有 %d 行在科目名/辅助核算维度里整格是「合计/小计/总计」，"
                                       "若与明细重复会翻倍。" % n_sub,
                             "suggest": ["加一步「删除小计行」一键清除，或「按关键词删行」手动指定列"]})

        # 5) 页脚：关键列为空但金额有值
        if has_jr:
            n_foot = q("""SELECT count(*) FROM journal WHERE nullif(trim(acct_code),'') IS NULL
                          AND (abs(coalesce(debit,0))>0.01 OR abs(coalesce(credit,0))>0.01)""", 0)
            if n_foot:
                findings.append({"level": "warn", "title": "疑似页脚行",
                                 "detail": "%d 行科目编码为空但金额有值，可能是页脚/合计行混入。" % n_foot,
                                 "suggest": ["加一步「按关键列判空删行」删除关键列为空的行"]})

        # 6) 整百万行（分卷/截断）
        if has_jr:
            n_jr = q('SELECT count(*) FROM journal', 0)
            if n_jr and n_jr % 1000000 == 0:
                findings.append({"level": "warn", "title": "行数恰为整百万",
                                 "detail": "序时账恰好 %d 行，可能是导出工具的分卷或截断，"
                                           "请确认是否还有其他分卷没导进来。" % n_jr,
                                 "suggest": None})

        # 7) 期间错位：序时账期间超出余额表（余额表无期间列，用检查页期间口径提醒）
        if has_tb and has_jr:
            try:
                jr_mx = q("SELECT max(_period_norm(period)) FROM journal WHERE nullif(trim(period),'') IS NOT NULL")
                if jr_mx:
                    findings.append({"level": "info", "title": "序时账最晚期间",
                                     "detail": "序时账期间最晚到 %s，检查第三/四组时请按余额表口径设期间过滤，"
                                               "否则期间外数据会造成假差异。" % jr_mx,
                                     "suggest": None})
            except Exception:
                pass

        if not findings:
            findings.append({"level": "ok", "title": "未发现明显问题",
                             "detail": "常规检查未命中重复/汇总/小计/页脚等情况。", "suggest": None})
        return {"findings": findings}

    def book_matches(self, only_confirm=False):
        """「账簿名 → 公司名」的完整判定清单（页面展示与导出共用）。

        **全列**，不只列推断出来的：客户要看的是「我这个账簿名最后被认成了谁、
        凭什么」。method=norm 表示归一后完全相同（可以自动认定），
        其余（norm_prefix / contain / prefix / similar）需要人看一眼。
        """
        sql = ("SELECT book_name, company_name, method, score, evidence, need_confirm, src "
               "FROM _book_match")
        if only_confirm:
            sql += " WHERE need_confirm = 1"
        sql += " ORDER BY need_confirm DESC, score ASC, book_name"
        cols = ["账簿名称", "推断的公司名称", "判断方法", "吻合度", "判断依据", "需人工确认", "来源"]
        try:
            rows = self.conn.execute(sql).fetchall()
        except Exception:
            rows = []
        return {"columns": cols, "rows": [list(r) for r in rows], "total": len(rows)}

    def book_matches_export(self, name=None):
        """把推断留痕导成 Excel——给客户复核用，改完可以整表当对照表回传。"""
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        d = self.book_matches()
        stem = re.sub(r'[\\/:*?"<>|]', "_",
                      name or ("账簿主体推断_" + datetime.now().strftime("%Y%m%d_%H%M%S")))
        path = os.path.join(self.out_dir, stem + ".xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "账簿主体推断"
        ws.append(d["columns"])
        for c in ws[1]:
            c.font = Font(bold=True, size=10)
            c.fill = PatternFill("solid", fgColor="DCE6F1")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in d["rows"]:
            ws.append(r)
        for i, w in enumerate((46, 42, 14, 9, 56, 12, 8), 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(d["columns"])), ws.max_row)
        for row in ws.iter_rows(min_row=2):
            if row[5].value:                      # 需人工确认 → 标黄
                for c in row:
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
        ws2 = wb.create_sheet("说明")
        for line in (["账簿主体推断留痕"], [],
                     ["1. 本表只列「表里出现过、但对照表没覆盖」的账簿名。"],
                     ["2. 判断方法：norm=归一后完全相同（可自动认定）；norm_prefix=去掉门店前缀后相同；"],
                     ["   contain=互相包含；prefix=共同开头；similar=文字相似度；none=没有相近主体。"],
                     ["3. 「需人工确认=1」的黄底行必须人看一眼——门店前缀和近似名最容易错配。"],
                     ["4. 确认后可直接把「账簿名称 / 推断的公司名称」两列回传当对照表用。"]):
            ws2.append(line)
        ws2.column_dimensions["A"].width = 90
        wb.save(path)
        return {"file": stem + ".xlsx", "rows": d["total"], "path": path}

    def recon_export(self, name=None):
        """把上一次检查的结果导出成一个多工作表 Excel。

        每个检查分组一张 sheet（2.1/2.2/2.3 各一张），第一张是结论汇总。
        表头用中文（recon 的字段字典），数字保持数值型，方便在 Excel 里继续算。
        """
        from openpyxl import Workbook
        tbls = getattr(self, "_recon_tables", None) or []
        meta = getattr(self, "_recon_meta", None) or {}
        if not tbls:
            # 重启后内存清空，但 _recon_last 还在库里，读回来继续导出
            last = self.recon_last()
            if last:
                tbls = last.get("tables") or []
                p = last.get("params") or {}
                meta = {"summary": last.get("summary") or {},
                        "tol": p.get("tol"), "period": p.get("period"),
                        "tb_label": p.get("tb_label")}
        if not tbls:
            raise ValueError("还没有检查结果，先点「开始检查」再导出")
        stem = name or ("数据检查_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
        stem = re.sub(r'[\\/:*?"<>|]', "_", stem)
        path = os.path.join(self.out_dir, stem + ".xlsx")
        wb = Workbook(write_only=True)
        with self.lock:
            sm = (meta.get("summary") or {})
            ws = wb.create_sheet("结论汇总")
            ws.append(["检查项", "结论", "检查数量", "问题数量"])
            for it in (sm.get("items") or []):
                if it.get("error"):
                    ws.append([it.get("title") or "", "无法检查：" + str(it["error"]), "", ""])
                    continue
                ws.append([it.get("title") or "",
                           "通过" if it.get("ok") else "有问题",
                           it.get("total") or 0, it.get("bad") or 0])
            ws.append([])
            ws.append(["容差", meta.get("tol")])
            ws.append(["余额表口径", meta.get("tb_label") or ""])
            ws.append(["期间从", (meta.get("period") or {}).get("from") or "全部"])
            ws.append(["期间到", (meta.get("period") or {}).get("to") or "全部"])
            ws.append(["导出时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
            ws.append([])
            ws.append(["说明", "「问题数量」= 不符 + 仅余额表有 + 仅序时账有；"
                              "序时账凭证表只列借贷不平的凭证；原币列仅展示、不参与判定"])

            for item in tbls:
                sheet = item.get("sheet") or "结果"
                table = item.get("table") or ""
                try:
                    cur = self.conn.execute('SELECT * FROM "%s" LIMIT 0' % table)
                    cols = [d[0] for d in cur.description]
                except Exception:
                    continue
                headers = list(item.get("headers") or [])
                if len(headers) != len(cols):
                    headers = cols
                ws = wb.create_sheet(_safe_sheet(sheet))
                ws.append(headers)
                # 空表也要说清楚为什么空，否则用户以为没跑成功
                try:
                    n_row = self.conn.execute('SELECT count(*) FROM "%s"' % table).fetchone()[0]
                except Exception:
                    n_row = 0
                if n_row == 0 and item.get("empty_note"):
                    ws.append([item["empty_note"]] + [""] * max(0, len(headers) - 1))
                cur = self.conn.execute('SELECT * FROM "%s"' % table)
                while True:
                    batch = cur.fetchmany(5000)
                    if not batch:
                        break
                    for row in batch:
                        ws.append([_num(v) for v in row])

            # ---- 数据来源 + 清洗留痕（底稿可自证）
            try:
                ws_src = wb.create_sheet("数据来源")
                ws_src.append(["源文件", "工作表", "目标表", "原始行数", "清洗后行数", "剔除行数", "已清洗"])
                for r in self.conn.execute(
                        "SELECT file_name, sheet_name, target, n_rows, stg_table, clean FROM _sheets "
                        "ORDER BY file_name, sheet_name").fetchall():
                    fn, sn, tg, n_raw, stg, cj = r
                    n_stg = 0
                    if stg:
                        try:
                            n_stg = self.conn.execute(
                                'SELECT count(*) FROM "%s"' % stg).fetchone()[0]
                        except Exception:
                            n_stg = 0
                    ws_src.append([fn, sn, tg or "", n_raw, n_stg, (n_raw or 0) - n_stg,
                                   "是" if stg else "否"])
                ws_tr = wb.create_sheet("清洗留痕")
                ws_tr.append(["源文件", "工作表", "步骤", "动作", "剔除原因", "剔除行数"])
                for r in self.conn.execute(
                        "SELECT sid, file_name, sheet_name FROM _sheets ORDER BY file_name, sheet_name"
                ).fetchall():
                    sid, fn, sn = r
                    try:
                        agg = self.conn.execute(
                            'SELECT step, step_name, reason, count(*) FROM "_dropped_%s" '
                            'GROUP BY 1,2,3 ORDER BY 1' % sid).fetchall()
                    except Exception:
                        agg = []
                    if not agg:
                        continue
                    for step, step_name, reason, cnt in agg:
                        ws_tr.append([fn, sn, step, step_name or "", reason or "", cnt])
            except Exception:
                pass
        wb.save(path)
        return {"path": path, "file": os.path.basename(path),
                "size": os.path.getsize(path), "sheets": len(tbls) + 1}

    # 按列值删行支持的条件。空值只能靠「为空 / 不为空」表达
    DEL_MODES = ("等于", "包含", "不等于", "为空", "不为空")

    @staticmethod
    def _del_cond(col, mode, value):
        """一条删除条件 → SQL 片段。col 必须已经过白名单校验。"""
        tgt = '"%s"' % col
        v = "" if value is None else str(value)
        if mode == "为空":
            return "nullif(trim(cast(%s AS VARCHAR)), '') IS NULL" % tgt
        if mode == "不为空":
            return "nullif(trim(cast(%s AS VARCHAR)), '') IS NOT NULL" % tgt
        if mode == "包含":
            # 子串匹配，大小写不敏感；把 % _ 转义掉，否则用户输「1002%」会当通配符用，
            # 一下删掉一大片，而且看不出来为什么。
            esc = v.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
            return ("lower(cast(%s AS VARCHAR)) LIKE lower(%s) ESCAPE '\\'"
                    % (tgt, ql("%" + esc + "%")))
        if mode == "不等于":
            # 空值也算「不等于」：用户选这条就是「留下等于的，其余全删」，
            # SQL 三值逻辑会把空值漏掉，留下一些看不出原因的行。
            return ("(lower(trim(cast(%s AS VARCHAR))) <> lower(trim(%s)) "
                    "OR nullif(trim(cast(%s AS VARCHAR)), '') IS NULL)"
                    % (tgt, ql(v), tgt))
        # 精确等于：忽略首尾空格和大小写
        return ("lower(trim(cast(%s AS VARCHAR))) = lower(trim(%s))"
                % (tgt, ql(v)))

    def _del_sample(self, table, where, n=3):
        """将要被删的前几行（中文列名），给前端二次确认用。"""
        try:
            cur = self.conn.execute(
                'SELECT * FROM "%s" WHERE %s LIMIT %d' % (table, where, n))
            names = [d[0] for d in cur.description]
            rows = cur.fetchall()
            pick = [c for c in ("company_name", "book_name", "period", "vdate",
                                "vtype", "vno", "acct_code", "acct_name",
                                "summary", "currency", "debit", "credit",
                                "end_bal")
                    if c in names]
            idx = [names.index(c) for c in pick]
            out, seen = [], set()
            for r in rows:
                key = tuple(r[i] for i in idx)
                if key in seen:          # 内容一样的样例给三遍没意义
                    continue
                seen.add(key)
                out.append([{"label": to_cn(c, table),
                             "value": "" if r[i] is None else str(r[i])}
                            for c, i in zip(pick, idx)])
            return out
        except Exception:
            return []

    def delete_rows(self, table, col=None, mode="等于", value="",
                    conditions=None, dry_run=False):
        """按列值删行：删除标准表（journal / tb）里满足条件的行。

        conditions = [{"col": "company_name", "mode": "等于", "value": "甲公司"},
                      {"col": "acct_name",   "mode": "等于", "value": "银行存款"}]
        多条之间是 **且（AND）**；老的单条调用（col / mode / value）继续支持。
        mode：等于 / 包含 / 不等于 / 为空 / 不为空
        dry_run=True 只统计会删多少行、给几行样本，不做任何改动。

        删除前把整表快照到 work 目录（_delbak_*.parquet），误删能找回来。
        """
        table = str(table or "")
        if table not in ("journal", "tb"):
            raise ValueError("只能对「序时账 journal」或「科目余额表 tb」操作")

        raw = [c for c in (conditions or []) if isinstance(c, dict)]
        if not raw and col:
            raw = [{"col": col, "mode": mode, "value": value}]
        if not raw:
            raise ValueError("至少要有一个删除条件")

        cols = {c for c, _ in CANON[table]}
        with self.lock:
            real = {r[0] for r in self.conn.execute(
                "SELECT column_name FROM information_schema.columns WHERE table_name = %s"
                % ql(table)).fetchall()}
            norm, sqls = [], []
            for c in raw:
                k = str(c.get("col") or "").strip()
                md = str(c.get("mode") or "等于").strip()
                v = "" if c.get("value") is None else str(c.get("value"))
                if not k:
                    raise ValueError("有一行条件没选「判断列」")
                if k not in cols or k not in real:
                    raise ValueError("列「%s」不在 %s 表里（可能是旧数据包）" % (k, table))
                if md not in self.DEL_MODES:
                    md = "等于"
                # 空串当「等于」会变成「删掉这一列为空的行」，太容易被误点，
                # 必须显式选「为空」才允许不填值。
                if md not in ("为空", "不为空") and not v.strip():
                    raise ValueError("条件「%s」没填值；要按空值删请把条件选成「为空」"
                                     % to_cn(k, table))
                norm.append({"col": k, "mode": md, "value": v})
                sqls.append(self._del_cond(k, md, v))

            where = " AND ".join(sqls)
            before = self.conn.execute('SELECT count(*) FROM "%s"' % table).fetchone()[0]
            matched = self.conn.execute(
                'SELECT count(*) FROM "%s" WHERE %s' % (table, where)).fetchone()[0]

            if dry_run or matched == 0:
                return {"table": table, "conditions": norm, "before": before,
                        "matched": matched, "deleted": 0, "after": before,
                        "dry_run": bool(dry_run),
                        "sample": self._del_sample(table, where) if matched else []}

            # 删除前快照，误删能找回
            try:
                bak = os.path.join(self.work_dir,
                                   "_delbak_%s_%s.parquet" % (table, uuid.uuid4().hex[:8]))
                self.conn.execute('COPY "%s" TO %s (FORMAT PARQUET)' % (table, ql(bak)))
            except Exception:
                bak = None
            sample = self._del_sample(table, where)
            self.conn.execute('DELETE FROM "%s" WHERE %s' % (table, where))
            after = self.conn.execute('SELECT count(*) FROM "%s"' % table).fetchone()[0]
            return {"table": table, "conditions": norm, "before": before,
                    "matched": matched, "deleted": before - after, "after": after,
                    "dry_run": False, "sample": sample,
                    "wiped": bool(before and after == 0),
                    "backup": os.path.basename(bak) if bak else None}

    # ------------------------------------------------------------ 账簿对照表
    def books_preview(self, fid):
        """导入前的预览：复用 inspect（不带 target 就不做字段映射）。"""
        return self.inspect(fid, None)

    def books_import(self, fid, sheet, header_row=1, book_col=None,
                     company_col=None, short_col=None):
        f = self._file(fid)
        hr = max(1, int(header_row or 1))
        if not book_col:
            raise ValueError("必须指定「账簿名称」是哪一列")
        tmp = "_booksrc"
        with self.lock:
            try:
                names = self._read_raw(f, sheet, hr, 1, tmp)
                for label, c in (("账簿名称", book_col), ("公司名称", company_col),
                                 ("公司简称", short_col)):
                    if c and c not in names:
                        raise ValueError("列「%s」不在工作表里（可选列：%s）"
                                         % (c, "、".join(names[:20])))
                def ex(c):
                    return "_text(%s)" % qi(c) if c else "NULL::VARCHAR"
                base = ("SELECT %s AS book_name, %s AS company_name, %s AS short_name "
                        "FROM %s WHERE nullif(trim(%s), '') IS NOT NULL"
                        % (ex(book_col), ex(company_col), ex(short_col), tmp, qi(book_col)))
                conf = self.conn.execute(
                    "SELECT book_name, count(DISTINCT coalesce(company_name, '')) AS c "
                    "FROM (%s) GROUP BY 1 HAVING c > 1 ORDER BY 1 LIMIT 20" % base).fetchall()
                n_rows = self.conn.execute("SELECT count(*) FROM (%s)" % base).fetchone()[0]
                src = f["file_name"] + " / " + sheet
                self.conn.execute(
                    "INSERT OR REPLACE INTO _books "
                    "(book_name, company_name, short_name, src, imported_at) "
                    "SELECT book_name, company_name, short_name, ?, now() "
                    "FROM (SELECT book_name, any_value(company_name) AS company_name, "
                    "             any_value(short_name) AS short_name FROM (%s) GROUP BY 1)"
                    % base, [src])
            finally:
                self.conn.execute("DROP TABLE IF EXISTS %s" % tmp)
            n_all = self.conn.execute("SELECT count(*) FROM _books").fetchone()[0]
        return {"ok": True, "rows": n_rows, "total": n_all, "source": src,
                "conflicts": [{"book_name": c[0], "n": int(c[1])} for c in conf]}

    def books_list(self, limit=500):
        with self.lock:
            try:
                cur = self.conn.execute(
                    "SELECT book_name, company_name, short_name, src FROM _books "
                    "ORDER BY book_name LIMIT %d" % int(limit))
                rows = [list(r) for r in cur.fetchall()]
                n = self.conn.execute("SELECT count(*) FROM _books").fetchone()[0]
            except Exception:
                rows, n = [], 0
        return {"columns": ["账簿名称", "公司名称", "公司简称", "来源"], "rows": rows, "total": n}

    def books_clear(self):
        with self.lock:
            self.conn.execute("DELETE FROM _books")
        return {"ok": True}

    # ------------------------------------------------------------ 清洗方案
    def _recipe_lib(self):
        """方案库：{步骤签名: 方案名}"""
        try:
            rows = self.conn.execute("SELECT name, steps FROM _recipes").fetchall()
        except Exception:
            return {}
        out = {}
        for name, sj in rows:
            try:
                out[recipe.steps_key(json.loads(sj or "[]"))] = name
            except Exception:
                continue
        return out

    def recipes_list(self):
        with self.lock:
            try:
                rows = self.conn.execute(
                    "SELECT name, steps, sig, uses, updated_at FROM _recipes "
                    "ORDER BY updated_at DESC NULLS LAST, name").fetchall()
            except Exception:
                rows = []
        out = []
        for name, sj, sig, uses, ts in rows:
            try:
                steps = json.loads(sj or "[]")
            except Exception:
                steps = []
            for s in steps:
                spec = clean.BY_TYPE.get(s.get("type"))
                s["name"] = (spec or {}).get("name") or s.get("type")
            out.append({"name": name, "steps": steps, "n": len(steps),
                        "sig": sig or "", "uses": int(uses or 0),
                        "updated_at": ts.strftime("%Y-%m-%d %H:%M") if ts else ""})
        return {"recipes": out, "total": len(out)}

    def _recipe_free_name(self, base, keep=None):
        base = re.sub(r"\.(xlsx|xlsm|csv|txt)$", "", str(base or ""), flags=re.I).strip()
        base = re.sub(r'[\\/:*?"<>|]', " ", base).strip() or "清洗方案"
        base = base[:32]
        try:
            have = {r[0] for r in self.conn.execute("SELECT name FROM _recipes").fetchall()}
        except Exception:
            have = set()
        if base not in have or base == keep:
            return base
        n = 1
        while True:
            n += 1
            cand = "%s (%d)" % (base[:26], n)
            if cand not in have:
                return cand

    def recipes_save(self, name, steps, sid=None, note=""):
        """把一套动作存进方案库。name 为空时按工作表名自动起名。"""
        steps = steps or []
        if not steps:
            raise ValueError("这套方案里没有动作，存不了")
        with self.lock:
            if not str(name or "").strip():
                base = ""
                if sid:
                    try:
                        base = self._sheet(sid).get("sheet_name") or ""
                    except Exception:
                        base = ""
                name = self._recipe_free_name(base or "清洗方案")
            name = str(name).strip()[:40]
            key = recipe.steps_key(steps)
            self.conn.execute(
                "INSERT OR REPLACE INTO _recipes (name, steps, sig, note, uses, updated_at) "
                "VALUES (?,?,?,?,COALESCE((SELECT uses FROM _recipes WHERE name=?),0)+1, now())",
                [name, json.dumps(steps, ensure_ascii=False), key, note or "", name])
            n_all = self.conn.execute("SELECT count(*) FROM _recipes").fetchone()[0]
        return {"ok": True, "name": name, "total": n_all}

    def recipes_save_many(self, plans):
        """把导入文件里的多个方案一次性存进方案库。

        导入清洗方案 xlsx 时调用：存一次，之后所有工作表都能从右上角下拉框直接选，
        不用为每张表重新导一遍文件。同名方案按最新内容覆盖。
        """
        saved, skipped = [], []
        for p in (plans or []):
            name = str((p or {}).get("name") or "").strip()
            steps = (p or {}).get("steps") or []
            if not steps:
                skipped.append({"name": name, "reason": "没有动作"})
                continue
            try:
                r = self.recipes_save(name, steps, note="从清洗方案 xlsx 导入")
                saved.append(r["name"])
            except Exception as e:
                skipped.append({"name": name, "reason": str(e)[:120]})
        return {"ok": True, "saved": saved, "n": len(saved),
                "skipped": skipped, "total": self.recipes_list()["total"]}

    def recipes_delete(self, name):
        with self.lock:
            self.conn.execute("DELETE FROM _recipes WHERE name=?", [str(name)])
        return {"ok": True}

    def recipes_rename(self, name, new_name):
        new_name = str(new_name or "").strip()[:40]
        if not new_name:
            raise ValueError("新名字不能为空")
        with self.lock:
            hit = self.conn.execute("SELECT count(*) FROM _recipes WHERE name=?", [new_name]).fetchone()[0]
            if hit:
                raise ValueError("已经有叫「%s」的方案了" % new_name)
            self.conn.execute("UPDATE _recipes SET name=? WHERE name=?", [new_name, str(name)])
        return {"ok": True, "name": new_name}

    def recipes_apply(self, name, sid):
        """把方案库里的方案套到某张工作表上（沿用导入时的降级规则）。"""
        with self.lock:
            r = self.conn.execute("SELECT steps FROM _recipes WHERE name=?", [str(name)]).fetchone()
            if not r:
                raise ValueError("方案「%s」不存在" % name)
            steps = json.loads(r[0] or "[]")
        return self.clean_recipe_apply(sid, steps)

    def _recipe_autosave(self, sid):
        """应用清洗成功后自动登记方案：同一套动作只登记一次。"""
        try:
            m = self._sheet(sid)
            steps = m.get("clean") or []
            if not steps:
                return None
            key = recipe.steps_key(steps)
            lib = self._recipe_lib()
            if key in lib:
                self.conn.execute("UPDATE _recipes SET uses=COALESCE(uses,0)+1, "
                                  "updated_at=now() WHERE name=?", [lib[key]])
                return {"name": lib[key], "exists": True}
            row = self.conn.execute("SELECT sheet_name, file_name FROM _sheets WHERE sid=?",
                                    [sid]).fetchone()
            base = (row[0] or row[1]) if row else ""
            name = self._recipe_free_name(base or "清洗方案")
            self.conn.execute(
                "INSERT OR REPLACE INTO _recipes (name, steps, sig, note, uses, updated_at) "
                "VALUES (?,?,?,?,1,now())",
                [name, json.dumps(steps, ensure_ascii=False), key,
                 "应用清洗时自动登记"])
            return {"name": name, "exists": False}
        except Exception:
            return None

    def clean_recipe_export(self):
        """把所有工作表的清洗动作 + 字段映射导出成 xlsx（一种模式一个 sheet）。"""
        with self.lock:
            rows = self.conn.execute(
                "SELECT file_name, sheet_name, target, clean, mapping, col_names "
                "FROM _sheets ORDER BY file_name, sheet_name").fetchall()
        if not rows:
            raise ValueError("还没有导入任何工作表，没有可导出的清洗方案")
        sheets, maps = [], {}
        for fn, sn, tgt, cj, mj, cols in rows:
            steps = json.loads(cj or "[]")
            for s in steps:
                spec = clean.BY_TYPE.get(s.get("type"))
                if spec:
                    s["_name"] = spec["name"]
                    s["_desc"] = spec.get("desc") or ""
            try:
                col_list = json.loads(cols or "[]")
            except Exception:
                col_list = []
            sig = recipe.sig_of(col_list)
            sheets.append({"file_name": fn, "sheet_name": sn, "target": tgt or "",
                           "steps": steps, "columns": col_list, "sig": sig})
            mp = json.loads(mj or "{}")
            if mp:
                # 同一套表头只留一份映射；前面已经存过且更全的就不覆盖
                old = maps.get(sig)
                if not old or len(mp) > len(old["mapping"]):
                    maps[sig] = {"target": tgt or "", "mapping": mp, "cols": col_list}
        plans = recipe.group(sheets)
        lib = self._recipe_lib()
        for p in plans:
            # 方案库里存过的，用你起过的名字；没存过的用自动名
            p["name"] = lib.get(p["key"]) or p.get("auto_name") or p["name"]
        stem = "清洗方案_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(self.out_dir, stem + ".xlsx")
        recipe.export(plans, maps, path,
                      meta={"app": "kanzhang", "sheets": len(sheets),
                            "plans": len(plans), "maps": len(maps),
                            "created_at": datetime.now().isoformat(timespec="seconds")})
        return {"path": path, "file": os.path.basename(path),
                "size": os.path.getsize(path), "plans": len(plans),
                "sheets": len(sheets), "maps": len(maps)}

    def clean_recipe_parse(self, fid):
        f = self._file(fid)
        low = (f["file_name"] or "").lower()
        if not low.endswith((".xlsx", ".xlsm")):
            raise ValueError("清洗方案只支持 .xlsx（用「导出清洗方案」得到的文件）")
        if not self._load_excel():
            raise RuntimeError("Excel 扩展不可用，读不了 xlsx")
        res = recipe.parse(f["path"])
        out = []
        for p in res["plans"]:
            out.append({"name": p["name"],
                        "steps": [{k: v for k, v in s.items() if k != "_no"}
                                  for s in p["steps"]],
                        "uses": p.get("uses") or [],
                        "n": len(p["steps"])})
        return {"plans": out, "maps": res.get("maps") or {},
                "maps_hk": self._maps_by_head_key(res.get("maps") or {}),
                "warnings": res.get("warnings") or []}

    @staticmethod
    def _maps_by_head_key(maps):
        """把「按签名索引的映射」再按表头原文索引一份，前端没有 md5，用原串直接对。

        老文件（v1.1.0）只有工作表名，原串就是工作表名，前端对不上时会自然跳过。
        """
        out = {}
        for sig, m in (maps or {}).items():
            k = str((m or {}).get("hkey") or "").strip() or sig
            out[k] = {"target": (m or {}).get("target") or "",
                      "mapping": (m or {}).get("mapping") or {},
                      "legacy": bool((m or {}).get("legacy"))}
        return out

    def clean_recipe_apply(self, sid, steps, mapping=None, target=None):
        with self.lock:
            m = self._sheet(sid)
            tbl = m["stg_table"] or m["table_name"]
            try:
                cols = clean.columns_of(self.conn, tbl)
            except Exception:
                cols = m["columns"]
            fixed, warns = recipe.normalize_steps(steps or [], clean.BY_TYPE, set(cols))
            self.conn.execute("UPDATE _sheets SET clean=? WHERE sid=?",
                              [json.dumps(fixed, ensure_ascii=False), sid])
            applied_map = 0
            if mapping:
                cur = dict(m["mapping"] or {})
                keep = set(cols)
                for k, v in mapping.items():
                    if str(k).startswith("_") or not v:
                        continue
                    if str(v) in keep:
                        cur[k] = v
                        applied_map += 1
                self.conn.execute(
                    "UPDATE _sheets SET mapping=?, target=coalesce(?, target) WHERE sid=?",
                    [json.dumps(cur, ensure_ascii=False), target, sid])
        return {"ok": True, "steps": fixed, "n": len(fixed),
                "mapping": applied_map, "warnings": warns}

    # ------------------------------------------------------------ rules
    def rules_path(self):
        return os.path.join(self.data_dir, "rules.json")

    def get_rules(self):
        d = os.path.join(self.app_dir, "rules", "default_rules.json")
        try:
            default = json.load(open(d, encoding="utf-8"))
        except Exception:
            default = {"rules": []}
        p = self.rules_path()
        if not os.path.exists(p):
            return default
        try:
            cur = json.load(open(p, encoding="utf-8"))
        except Exception:
            return default
        # 版本升级带进来的新内置规则：只补用户库里没有的 id，用户改过的同名规则不动
        try:
            have = {r.get("id") for r in (cur.get("rules") or [])}
            add = [r for r in (default.get("rules") or []) if r.get("id") not in have]
            if add:
                cur.setdefault("rules", []).extend(add)
        except Exception:
            pass
        return cur

    def save_rules(self, data):
        with open(self.rules_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return {"ok": True}

    def export_rules_xlsx(self):
        """把当前规则导出成 xlsx（分类/名称/描述/参数/SQL），便于编辑与分发给 AI 或同事。"""
        from openpyxl import Workbook
        rules = (self.get_rules() or {}).get("rules") or []
        stem = "查询规则_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(self.out_dir, stem + ".xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "查询规则"
        ws.append(["分类", "名称", "描述", "参数(JSON)", "SQL"])
        for r in rules:
            params = json.dumps(r.get("params") or [], ensure_ascii=False)
            ws.append([r.get("cat") or "", r.get("name") or "",
                       r.get("desc") or "", params, r.get("sql") or ""])
        # 列宽：SQL 列加宽，方便看
        for col, w in (("A", 18), ("B", 22), ("C", 40), ("D", 30), ("E", 90)):
            ws.column_dimensions[col].width = w
        wb.save(path)
        return {"path": path, "file": os.path.basename(path),
                "size": os.path.getsize(path), "count": len(rules)}

    def import_rules_xlsx(self, fid, mode="append"):
        """从 xlsx 导入查询规则。

        mode=append：追加（id 冲突时用新 id 覆盖同名）；mode=replace：整库替换。
        xlsx 结构：分类/名称/描述/参数(JSON)/SQL（与导出格式一致）。
        """
        from openpyxl import load_workbook
        f = self._file(fid)
        wb = load_workbook(f["path"], data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            raise ValueError("规则文件是空的")
        hdr = [str(x).strip() if x is not None else "" for x in rows[0]]
        # 容错：不强制表头顺序，按名字定位
        def col_of(name):
            for i, h in enumerate(hdr):
                if name in h:
                    return i
            return -1
        i_cat, i_name, i_desc, i_params, i_sql = (
            col_of("分类"), col_of("名称"), col_of("描述"), col_of("参数"), col_of("SQL"))
        if i_name < 0 or i_sql < 0:
            raise ValueError("规则文件缺少「名称」或「SQL」列（请用导出的模板格式）")
        imported = []
        for row in rows[1:]:
            if row is None:
                continue
            name = str(row[i_name]).strip() if i_name < len(row) and row[i_name] is not None else ""
            sql = str(row[i_sql]).strip() if i_sql < len(row) and row[i_sql] is not None else ""
            if not name or not sql:
                continue
            cat = str(row[i_cat]).strip() if i_cat >= 0 and i_cat < len(row) and row[i_cat] is not None else "导入"
            desc = str(row[i_desc]).strip() if i_desc >= 0 and i_desc < len(row) and row[i_desc] is not None else ""
            params = []
            if i_params >= 0 and i_params < len(row) and row[i_params]:
                try:
                    params = json.loads(str(row[i_params]))
                except Exception:
                    params = []
            imported.append({"id": "x_" + uuid.uuid4().hex[:10], "cat": cat, "name": name,
                             "desc": desc, "params": params, "sql": sql})
        if not imported:
            raise ValueError("没有解析到任何有效规则（每行至少要有「名称」和「SQL」）")
        if mode == "replace":
            cur = {"version": 1, "rules": imported}
        else:
            cur = self.get_rules() or {"version": 1, "rules": []}
            # 追加：同名覆盖（保留原 id），其余新增
            by_name = {r.get("name"): r for r in (cur.get("rules") or [])}
            for r in imported:
                if r["name"] in by_name:
                    old = by_name[r["name"]]
                    old.update({"cat": r["cat"], "desc": r["desc"],
                                "params": r["params"], "sql": r["sql"]})
                else:
                    cur.setdefault("rules", []).append(r)
        self.save_rules(cur)
        return {"ok": True, "count": len(imported), "mode": mode}

    def backup(self):
        st = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = os.path.join(self.out_dir, "ledger_backup_" + st + ".duckdb")
        with self.lock:
            self.conn.execute("CHECKPOINT")
        shutil.copy2(self.db_path, dst)
        return {"path": dst, "file": os.path.basename(dst), "size": os.path.getsize(dst)}


def _s(v):
    if v is None:
        return ""
    if isinstance(v, (date, datetime)):
        return v.isoformat()
    return str(v)


def _subst(sql, params):
    """参数按原文替换：文本请自行在 SQL 里加引号（LIKE '%{{kw}}%'），
    这样会写 SQL 的人可以直接照着改，不会出现双重引号。"""
    if not params:
        return sql
    for k, v in params.items():
        if v is None:
            lit = "NULL"
        elif isinstance(v, bool):
            lit = "TRUE" if v else "FALSE"
        else:
            lit = str(v)
        sql = sql.replace("{{%s}}" % k, lit)
    return sql
