# -*- coding: utf-8 -*-
"""判断留痕（_judgments / _todos）· 纯数据层

为什么单独一个文件：这是横切三个板块的能力（货币资金、销管费用、以后 G3），
塞进 modules/bkd/core.py（已 2500+ 行）会变成第二个单体。本模块只碰数据，
不碰 xlsx、不碰 HTTP，谁都能调。

核心概念（三者要分清，否则会和 gaps 混在一起）：

  gap      症状：本次生成发现的一个空格子。实时算，不落表。
  cause    成因：为什么空。复用 recon._aux_origin 的五态。
  judgment 处方：值 + 依据 + 状态。落表，可跨期继承、可复核。

写死的边界：
  · 只有「人工」和「AI」的判断入库。规则命中、上年底稿回捞每次实时算——
    它们可复现，入库只会制造「改了规则被旧缓存压住」的不一致。
  · AI 最高到 draft，本模块不提供写 confirmed 的 AI 入口（接口层就不给）。
"""
import hashlib
import re
from datetime import datetime

# ---------------------------------------------------------------- 常量

SOURCES = ("rule", "prior_bkd", "prior_judgment", "ai", "human", "calc")
STATUSES = ("draft", "confirmed", "rejected", "expired")

#: 需要职业判断的字段。这类空着不能去问客户，只能 AI 起草 + 人确认（C 类）。
JUDGE_FIELDS = {
    "nature", "confirm", "conclusion", "sample_pick",
    "restricted", "reclass",
}

#: 底稿表头中文 -> 字段 key。gap 里存的是中文标签，判定要 key。
FIELD_KEYS = {
    "性质/用途": "nature", "性质": "nature", "用途": "nature",
    "是否函证": "confirm", "函证": "confirm",
    "银行/存款机构": "bankName", "开户行": "bankName", "银行名称": "bankName",
    "账号": "acctNo", "银行账号": "acctNo",
    "币种": "currency", "币别": "currency",
    "Notes": "notes", "备注": "notes",
    "费用类别": "expCat", "受限": "restricted",
}

#: 缺料时要客户补什么。直接可复制进邮件（B 类的 ask 文案）。
ASK_TEXT = {
    "bankName": "该账户的开户银行全称（含支行）",
    "acctNo": "该账户的银行账号",
    "currency": "该账户的币种",
    "nature": "该账户是否为受限资金（保证金 / 冻结 / 质押），受限金额多少",
    "confirm": "该账户是否为银行账户、是否需要函证",
    "notes": "该账户的补充说明",
}

#: 缺料成因 -> 要客户补的东西（用于 _todos.why）
CAUSE_TEXT = {
    "never_present": "原始数据没有这一列（客户导出时就没带）",
    "never_filled": "原始表有这一列，客户没填",
    "not_mapped": "字段映射漏配（平台可立即修复，不用重导）",
    "lost": "映射配了但合并结果为空，疑清洗误删",
    "no_prior": "上年底稿没有该账户，无历史可回捞",
    "prior_blank": "上年底稿该账户这一项是空的",
}

#: 缺料成因 -> gap 分类（A 能定 / B 缺料 / C 待判 / D 配错）
CAUSE_CLASS = {
    "present": "A",
    "never_present": "B",
    "never_filled": "B",
    "not_mapped": "D",
    "lost": "D",
    "no_prior": "C",
    "prior_blank": "C",
}

_DDL = [
    """CREATE TABLE IF NOT EXISTS _judgments(
        jid VARCHAR PRIMARY KEY,
        scope VARCHAR, entity VARCHAR, subject VARCHAR,
        dim_type VARCHAR, dim_key VARCHAR, period VARCHAR, field VARCHAR,
        value VARCHAR, source VARCHAR, basis VARCHAR, confidence DOUBLE,
        status VARCHAR, run_id VARCHAR,
        created_by VARCHAR, created_at TIMESTAMP,
        reviewed_by VARCHAR, reviewed_at TIMESTAMP)""",
    """CREATE TABLE IF NOT EXISTS _todos(
        tid VARCHAR PRIMARY KEY,
        scope VARCHAR, entity VARCHAR, dim_type VARCHAR, dim_key VARCHAR, field VARCHAR,
        title VARCHAR, why VARCHAR, ask VARCHAR, affected INTEGER,
        status VARCHAR, note VARCHAR, updated_at TIMESTAMP)""",
]


# ---------------------------------------------------------------- 基础

def ensure(conn):
    """建表，幂等。db._init() 也会调，这里再兜一次（换账套时）。"""
    if conn is None:
        return False
    for ddl in _DDL:
        try:
            conn.execute(ddl)
        except Exception:
            return False
    return True


def _norm_dim(dim_type, dim_key):
    """维度键归一。账户走账号归一（全角→半角、去空格连字符、去前导零），
    其余维度只去空白——供应商名不能被「去前导零」糟蹋。"""
    if dim_key is None:
        return ""
    t = str(dim_key).strip()
    if not t:
        return ""
    if (dim_type or "") == "acct":
        from .accounts import _acct_norm_py
        return _acct_norm_py(t)
    return re.sub(r"\s+", "", t)


def jid(scope, entity, subject, dim_type, dim_key, period, field):
    """业务键哈希做主键：同一格重生成十次落在同一条上（天然 upsert）。
    自增 ID 会在跑三个月后堆出 800 条重复判断，没人敢用。"""
    raw = "|".join([
        str(scope or ""), str(entity or "").strip(), str(subject or "").strip(),
        str(dim_type or ""), _norm_dim(dim_type, dim_key), str(period or ""),
        str(field or ""),
    ])
    return hashlib.md5(raw.encode("utf-8")).hexdigest()[:16]


def _row_to_dict(row):
    keys = ("jid", "scope", "entity", "subject", "dimType", "dimKey", "period",
            "field", "value", "source", "basis", "confidence", "status", "runId",
            "createdBy", "createdAt", "reviewedBy", "reviewedAt")
    d = dict(zip(keys, row))
    for k in ("createdAt", "reviewedAt"):
        if d.get(k) is not None:
            d[k] = str(d[k])[:19]
    return d


_SELECT = ("SELECT jid, scope, entity, subject, dim_type, dim_key, period, field, "
           "value, source, basis, confidence, status, run_id, "
           "created_by, created_at, reviewed_by, reviewed_at FROM _judgments")


def get(conn, **kw):
    """按业务键取一条（不继承）。找不到返回 None。"""
    if conn is None:
        return None
    j = jid(kw.get("scope"), kw.get("entity"), kw.get("subject"),
            kw.get("dim_type"), kw.get("dim_key"), kw.get("period"), kw.get("field"))
    try:
        row = conn.execute(_SELECT + " WHERE jid=?", [j]).fetchone()
    except Exception:
        return None
    return _row_to_dict(row) if row else None


# ---------------------------------------------------------------- 写

def upsert(conn, scope, entity, subject, dim_type, dim_key, period, field,
           value, source="human", basis="", confidence=None,
           status="confirmed", run_id="", created_by=""):
    """写一条判断。同键覆盖（业务键哈希保证幂等）。

    返回 (jid, status, error)。error 非空表示没写进去。
    """
    if conn is None:
        return None, None, "没有账套连接"
    if source not in SOURCES:
        return None, None, "未知来源：%s" % source
    if status not in STATUSES:
        return None, None, "未知状态：%s" % status
    if source == "ai" and status == "confirmed":
        return None, None, "AI 起草不能直接置为已确认（须人工确认）"
    v = "" if value is None else str(value).strip()
    if not v and status == "confirmed":
        return None, None, "确认的值为空"
    ensure(conn)
    j = jid(scope, entity, subject, dim_type, dim_key, period, field)
    # 人工已确认的点位，AI 再起草一次不能把它冲掉——起草是草案，不能覆盖定论。
    # 要改结论只能由人再 save 一次，或先 reject。
    if status == "draft":
        try:
            old = conn.execute("SELECT status FROM _judgments WHERE jid=?", [j]).fetchone()
            if old and old[0] == "confirmed":
                return None, None, "该点位已有确认结论，AI 起草不能覆盖（如需修改请人工确认）"
        except Exception:
            pass
    now = datetime.now()
    try:
        conn.execute(
            "INSERT OR REPLACE INTO _judgments VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            [j, scope or "", str(entity or "").strip(), str(subject or "").strip(),
             dim_type or "", _norm_dim(dim_type, dim_key), period or "", field or "",
             v, source, basis or "", confidence, status, run_id or "",
             created_by or "", now, None, None])
    except Exception as e:
        return None, None, "写入失败：%s" % e
    return j, status, None


def confirm(conn, j, by="", value=None, basis=None):
    """人工确认（可顺带改值、改依据）。只有这一步能把状态推到 confirmed。"""
    if conn is None:
        return False, "没有账套连接"
    row = conn.execute(_SELECT + " WHERE jid=?", [j]).fetchone()
    if not row:
        return False, "没有这条判断"
    d = _row_to_dict(row)
    nv = d["value"] if value is None else str(value).strip()
    if not nv:
        return False, "确认的值不能为空"
    nb = d["basis"] if basis is None else basis
    conn.execute(
        "UPDATE _judgments SET value=?, basis=?, status='confirmed', "
        "reviewed_by=?, reviewed_at=? WHERE jid=?", [nv, nb, by or "", datetime.now(), j])
    return True, None


def reject(conn, j, by="", why=""):
    """驳回。保留原值但状态置 rejected，下次生成不再命中。"""
    if conn is None:
        return False, "没有账套连接"
    try:
        conn.execute(
            "UPDATE _judgments SET status='rejected', basis=?, reviewed_by=?, "
            "reviewed_at=? WHERE jid=?", [why or "", by or "", datetime.now(), j])
        return True, None
    except Exception as e:
        return False, str(e)


def list_judgments(conn, scope=None, status=None, entity=None, limit=500):
    if conn is None:
        return []
    ensure(conn)
    sql, args = _SELECT + " WHERE 1=1", []
    if scope:
        sql += " AND scope=?"; args.append(scope)
    if status:
        sql += " AND status=?"; args.append(status)
    if entity:
        sql += " AND entity=?"; args.append(entity)
    sql += " ORDER BY created_at DESC LIMIT ?"
    args.append(int(limit))
    try:
        return [_row_to_dict(r) for r in conn.execute(sql, args).fetchall()]
    except Exception:
        return []


def stats(conn):
    """面板用：各状态计数 + 按 scope 分布。"""
    out = {"total": 0, "draft": 0, "confirmed": 0, "rejected": 0,
           "byScope": {}, "todos": {}}
    if conn is None:
        return out
    ensure(conn)
    try:
        for st, n in conn.execute(
                "SELECT status, count(*) FROM _judgments GROUP BY status").fetchall():
            out[st] = int(n or 0)
            out["total"] += int(n or 0)
        for sc, n in conn.execute(
                "SELECT scope, count(*) FROM _judgments GROUP BY scope").fetchall():
            out["byScope"][sc] = int(n or 0)
        for st, n in conn.execute(
                "SELECT status, count(*) FROM _todos GROUP BY status").fetchall():
            out["todos"][st] = int(n or 0)
    except Exception:
        pass
    return out


# ---------------------------------------------------------------- 取值（核心）

def _lookup(conn, scope, entity, subject, dim_type, dim_key, field, period, want):
    """按周期优先级取一条：本期精确 > 往期精确（最近）> 跨期通用。
    返回 dict 或 None。"""
    if conn is None:
        return None
    ensure(conn)
    base = (_SELECT + " WHERE scope=? AND entity=? AND subject=? AND dim_type=? "
            "AND dim_key=? AND field=? AND status=?")
    args = [scope or "", str(entity or "").strip(), str(subject or "").strip(),
            dim_type or "", _norm_dim(dim_type, dim_key), field, want]
    try:
        if period:
            row = conn.execute(base + " AND period=?", args + [period]).fetchone()
            if row:
                return _row_to_dict(row)
            row = conn.execute(
                base + " AND period<>'' AND period<? ORDER BY period DESC LIMIT 1",
                args + [period]).fetchone()
            if row:
                return _row_to_dict(row)
        row = conn.execute(base + " AND period='' LIMIT 1", args).fetchone()
        return _row_to_dict(row) if row else None
    except Exception:
        return None


def resolve(conn, scope, entity, subject, dim_type, dim_key, period, field, *,
            rules=None, prior_row=None):
    """六级链条取一个值，返回 (value, meta)。value 为 None 表示六条都没命中。

    1 本期人工确认 → 2 往期确认继承 → 3 上年底稿回捞 → 4 平台规则
    → 5 AI 起草（**仍然标黄**） → 6 空

    规则必须排在 AI 之前：准则明确的事（银行存款一律函证）不该让 AI 再说一遍，
    说了只增加复核量。
    """
    meta = {"source": None, "basis": "", "status": None, "jid": None,
            "confidence": None, "label": ""}

    # 1 / 2：人工已确认（含跨期继承）
    hit = _lookup(conn, scope, entity, subject, dim_type, dim_key, field, period, "confirmed")
    if hit and hit.get("value"):
        inherited = bool(period) and hit.get("period") != period
        meta.update(source="prior_judgment" if inherited else "human",
                    basis=hit.get("basis") or "", status="confirmed",
                    jid=hit.get("jid"),
                    label=("沿用 %s 判断" % hit.get("period")) if inherited else "人工判断")
        return hit["value"], meta

    # 3：上年底稿回捞
    if prior_row:
        pv = prior_row.get(field)
        if pv is not None and str(pv).strip():
            meta.update(source="prior_bkd", status="confirmed",
                        basis="上年底稿回捞", label="上年底稿回捞")
            return str(pv).strip(), meta

    # 4：平台写死规则
    if rules:
        rv = rules.get(field) if hasattr(rules, "get") else None
        if rv is not None and str(rv).strip():
            meta.update(source="rule", status="confirmed",
                        basis=str((rules.get("_basis") or {}).get(field, "平台规则")
                                  if isinstance(rules.get("_basis"), dict) else "平台规则"),
                        label="规则")
            return str(rv).strip(), meta

    # 5：AI 起草 —— 命中但仍是 draft，调用方必须继续标黄
    draft = _lookup(conn, scope, entity, subject, dim_type, dim_key, field, period, "draft")
    if draft and draft.get("value"):
        meta.update(source="ai", basis=draft.get("basis") or "",
                    status="draft", jid=draft.get("jid"),
                    confidence=draft.get("confidence"),
                    label="AI 起草（%s），待复核" %
                          ("%.2f" % draft["confidence"] if draft.get("confidence") is not None else "-"))
        return draft["value"], meta

    return None, meta


def note_of(value, meta, prefix=""):
    """给底稿 Notes / 单元格批注用的一句话。依据必须跟着值走。"""
    if value is None:
        return ""
    lab = meta.get("label") or ""
    bas = (meta.get("basis") or "").strip()
    txt = lab
    if bas and bas != lab:
        txt = ("%s：%s" % (lab, bas)) if lab else bas
    return (prefix + txt) if prefix else txt


def is_draft(meta):
    """AI 起草的值必须继续标黄——写进格子是为了让人看着改，不是让它变结论。"""
    return bool(meta) and meta.get("status") == "draft"


# ---------------------------------------------------------------- 分诊

def field_key(label):
    if not label:
        return ""
    if label in FIELD_KEYS:
        return FIELD_KEYS[label]
    return FIELD_KEYS.get(str(label).strip(), "")


def classify_gap(gap, cause=None, has_col=True, all_blank=False, has_judgment=False,
                 source=None):
    """一个黄格子该归谁处理。返回 (class, hint)。

    判定顺序不能换：先看有没有结论，再看是不是自己配错，最后才分「待判」和「缺料」。
    D 类（配错）必须排在 B 类之前——映射漏配被当成缺资料发给客户，是纯粹的内耗。
    """
    fk = gap.get("field_key") or field_key(gap.get("field"))
    if has_judgment or gap.get("resolved"):
        return "A", "已有结论，直接填（来源：%s）" % (
            source or gap.get("source") or "判断留痕")
    if cause in CAUSE_CLASS:
        cls = CAUSE_CLASS[cause]
    elif cause:
        cls = "B"
    elif not has_col:
        cls = "C" if fk in JUDGE_FIELDS else "B"
    elif all_blank:
        cls = "B"
    else:
        cls = "C"
    if cls == "C" and fk not in JUDGE_FIELDS:
        cls = "B"
    return cls, hint_of(cls, fk, cause, gap)


def hint_of(cls, fk, cause=None, gap=None):
    """每个黄格子后面直接给「怎么办」，不要只写「需人工判断」。"""
    g = gap or {}
    if cls == "A":
        return "已有结论，直接填"
    if cls == "D":
        stage = {"not_mapped": "字段映射漏配", "lost": "清洗误删"}.get(cause or "", "配置问题")
        return "%s：改配置即可，不用重导（按主体 %s / 账户 %s 去查）" % (
            stage, g.get("company") or "-", g.get("acct") or "-")
    if cls == "C":
        return "需职业判断：可由 AI 起草依据，人工确认后生效"
    ask = ASK_TEXT.get(fk) or ("该字段（%s）" % (g.get("field") or fk or "?"))
    why = CAUSE_TEXT.get(cause or "", "源表未提供")
    return "缺料：请客户补充%s（%s）" % (ask, why)


def classify_all(gaps, causes=None, cols=None):
    """批量分诊。causes: {(entity, dim_key, field): cause}；cols: {field: has_col}。
    返回 (gaps 带 class/hint, 四类计数)。"""
    causes = causes or {}
    cols = cols or {}
    out, cnt = [], {"A": 0, "B": 0, "C": 0, "D": 0}
    for g in gaps or []:
        fk = g.get("field_key") or field_key(g.get("field"))
        cause = causes.get((g.get("company"), _norm_dim("acct", g.get("acct")), fk)) \
            or g.get("cause")
        cls, hint = classify_gap(g, cause=cause,
                                 has_col=cols.get(fk, True),
                                 all_blank=bool(g.get("allBlank")),
                                 has_judgment=bool(g.get("resolved")))
        gg = dict(g)
        gg["field_key"] = fk
        gg["class"] = cls
        gg["hint"] = hint
        out.append(gg)
        cnt[cls] = cnt.get(cls, 0) + 1
    return out, cnt


# ---------------------------------------------------------------- 待办

def todo_sync(conn, items, scope=""):
    """B 类缺料合并成待办。同一 (主体, 维度, 字段) 只留一条，affected 累加——
    客户回一封邮件能补 8 个账户，清单应该是 1 条而不是 8 条。

    已 received / waived 的不回退到 open（否则每次重生成都会「复活」已办的事）。
    """
    if conn is None:
        return {"open": 0, "asked": 0, "received": 0, "waived": 0, "added": 0}
    ensure(conn)
    added = touched = 0
    errs = []
    for it in items or []:
        dt = it.get("dim_type") or "acct"
        dk = _norm_dim(dt, it.get("dim_key"))
        fk = it.get("field_key") or it.get("field") or ""
        tid = hashlib.md5(("|".join([scope or "", str(it.get("entity") or ""),
                                     dk, fk])).encode("utf-8")).hexdigest()[:16]
        try:
            row = conn.execute("SELECT status, affected FROM _todos WHERE tid=?",
                               [tid]).fetchone()
            if row:
                st = row[0] or "open"
                n = int(row[1] or 0) + int(it.get("affected") or 1)
                if st in ("received", "waived"):
                    conn.execute("UPDATE _todos SET affected=? WHERE tid=?", [n, tid])
                else:
                    conn.execute(
                        "UPDATE _todos SET affected=?, why=?, ask=?, title=?, updated_at=? "
                        "WHERE tid=?", [n, it.get("why") or "", it.get("ask") or "",
                                        it.get("title") or "", datetime.now(), tid])
                touched += 1
            else:
                conn.execute(
                    "INSERT INTO _todos VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    [tid, scope or "", str(it.get("entity") or ""), dt, dk, fk,
                     it.get("title") or "", it.get("why") or "", it.get("ask") or "",
                     int(it.get("affected") or 1), "open", "", datetime.now()])
                added += 1
        except Exception as e:
            # 不静默吞：一条写失败要让调用方知道，否则清单少了人还以为齐了
            errs.append(str(e))
            continue
    out = dict(todo_stats(conn), added=added, touched=touched)
    if errs:
        out["errors"] = errs[:5]
    return out


def todo_stats(conn):
    out = {"open": 0, "asked": 0, "received": 0, "waived": 0}
    if conn is None:
        return out
    try:
        for st, n in conn.execute(
                "SELECT status, count(*) FROM _todos GROUP BY status").fetchall():
            if st in out:
                out[st] = int(n or 0)
    except Exception:
        pass
    return out


def list_todos(conn, status=None, limit=500):
    if conn is None:
        return []
    ensure(conn)
    sql = ("SELECT tid, scope, entity, dim_type, dim_key, field, title, why, ask, "
           "affected, status, note, updated_at FROM _todos")
    args = []
    if status:
        sql += " WHERE status=?"
        args.append(status)
    sql += " ORDER BY updated_at DESC LIMIT ?"
    args.append(int(limit))
    try:
        rows = conn.execute(sql, args).fetchall()
    except Exception:
        return []
    keys = ("tid", "scope", "entity", "dimType", "dimKey", "field", "title", "why",
            "ask", "affected", "status", "note", "updatedAt")
    out = []
    for r in rows:
        d = dict(zip(keys, r))
        if d.get("updatedAt") is not None:
            d["updatedAt"] = str(d["updatedAt"])[:19]
        out.append(d)
    return out


def todo_set(conn, tid, status, note=""):
    """推进待办状态。waived（客户明确不提供）必须留痕——
    底稿要能回答「这一格为什么空」，答案不能是「忘了」。"""
    if conn is None:
        return False, "没有账套连接"
    if status not in ("open", "asked", "received", "waived"):
        return False, "未知状态：%s" % status
    try:
        conn.execute("UPDATE _todos SET status=?, note=?, updated_at=? WHERE tid=?",
                     [status, note or "", datetime.now(), tid])
        return True, None
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------------- 导出

def export_rows(conn):
    """判断留痕导出（Excel 检查表 / CSV 共用）。"""
    rows = list_judgments(conn, limit=5000)
    out = [["主体", "科目", "账户/维度", "期间", "字段", "值", "来源", "依据",
            "置信度", "状态", "复核人", "时间"]]
    src_cn = {"rule": "规则", "prior_bkd": "上年底稿", "prior_judgment": "往期继承",
              "ai": "AI 起草", "human": "人工", "calc": "计算"}
    st_cn = {"draft": "待确认", "confirmed": "已确认",
             "rejected": "已驳回", "expired": "已失效"}
    for r in rows:
        out.append([r.get("entity"), r.get("subject"), r.get("dimKey"),
                    r.get("period"), r.get("field"), r.get("value"),
                    src_cn.get(r.get("source"), r.get("source")), r.get("basis"),
                    r.get("confidence"), st_cn.get(r.get("status"), r.get("status")),
                    r.get("reviewedBy") or "", r.get("createdAt") or ""])
    return out
