# -*- coding: utf-8 -*-
"""费用 OAR（分析性复核）—— 现场生成，结论交给 AI 写。

三条分工，是踩过一轮之后定下来的：

  1. **工具负责「算」**：本期/上期审定数、变动金额、变动比例、占全部费用比重，
     全部是公式，链接 BKD 的 J/K 列 —— BKD 一改，OAR 立刻跟着变。
  2. **工具负责「摆证据」**：把每个科目的序时账按「对方单位 / 业务类型」分组，
     算出每组金额、笔数、占该科目比重，按金额排序。这是变动归因的**候选集**，
     不是结论 —— 谁推动了变动，要结合业务判断。
  3. **AI 负责「下结论」**：一段话说明变动原因。工具只写占位符并标黄，
     AI 通过 apply_conclusions() 写进来之后变白。

为什么分析对象按**变动金额/变动比例**挑，不按本期金额：
  OAR 看的是「变动」。本期花 1000 万、上期也花 1000 万，说明不了任何问题；
  本期只花 100 万、上期是 20 万，才是要问清楚的。按本期金额排会把这些全漏掉。

为什么不是「摘金额最大的两笔凭证」：
  最大的那笔往往只是月末结转或一次性调整（真实数据里就有一条 1,148 万的中介费
  结转分录、一条 −299 万的审计调整），摘出来对解释变动毫无用处。
  必须先**按业务分组汇总**，再看哪一组构成了变动的大头。
"""

import json
import os
import re

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

#: 旧版的单页名（V1.2.78 及以前）。只用于「重跑时清残页」，不再生成。
OAR_SHEET = "VC&VD OAR"
#: V1.2.79 起销售费用、管理费用各占一张页。两类费用挤在一页上，
#: 复核的人要在一页里来回找两套「合计」，还容易把销售费用的口径当成管理费用的。
OAR_SHEET_VC = "VC&VD OAR 销售费用"
OAR_SHEET_VD = "VC&VD OAR 管理费用"
OAR_SHEET_OF = {"销售费用": OAR_SHEET_VC, "管理费用": OAR_SHEET_VD}
OAR_SHEETS = (OAR_SHEET_VC, OAR_SHEET_VD)

GAP_COLOR = "FFFF00"
HEAD_COLOR = "4F6228"
ACC_FMT = r'_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '
MONEY_FMT = ACC_FMT
PCT_FMT = "0.00%"

_THIN = Side(style="thin", color="808080")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
GAP_FILL = PatternFill("solid", fgColor=GAP_COLOR)

#: AI 未写结论时的占位符。apply_conclusions() 靠它定位要替换的那段文字。
AI_TOKEN = "【待AI撰写：变动原因需结合业务判断】"

#: 摘要里表示「结转损益」的，不是真实业务。
#: **只用来决定「贷方要不要抵」，不能拿来整行剔除** ——
#: 真实数据里「结转----确认上市费用」是一笔 1,148 万的**借方**真实费用，
#: 整行剔掉会让分组金额少一半（实测 8.69M vs 17.31M）。
SKIP_SUMMARY = ("结转", "转出", "结平")

# ---------------------------------------------------------------- 摘要解析
_CORP_STRONG = (
    "股份有限公司", "有限责任公司", "有限公司", "会计师事务所", "律师事务所",
    "税务师事务所", "工作室", "事务所", "研究院", "设计院", "学院", "大学",
    "医院", "分公司", "子公司", "集团", "公司", "分行", "支行", "营业部",
    "旅行社", "酒店", "宾馆", "超市", "工厂", "门市部", "经营部",
)
_CORP_WEAK = ("银行", "中心", "商店")
_SPLIT_RE = re.compile(r"[-—－–~～:：/、,，;；\s]+")
_LEAD_NOISE = re.compile(r"^(?:支付给|付给|支付|预付|代付|与|向|从|对|给|和|及|跟|把|将)+[「『\"'《]{0,2}")
_TAIL_NOISE = re.compile(r"^[」』\"'》]?(发生的|发生|签订|支付|收取)?")
_BRACKET_TAIL = re.compile(r"[（(][^（()）]{0,20}[)）]$")

#: 只放多字词 + 「付」。单字的「转」「冲」不能放：会把「转账手续费」削成「账手续费」
_NOISE = ("审计调整", "调整", "计提", "付款", "支付", "预付", "暂估", "红冲",
          "冲销", "结转", "转入", "报销", "付")

#: 分组时要跳过的「不像业务名」的片段：月份、凭证号、纯数字
_NOT_BIZ = re.compile(
    r"^(\d{1,2}月|\d{1,2}|\d{1,3}#?\d*#?|\d{4}年|本期|上期|本年|上年|\d+-\d+月?)$")


def _clean_biz(s, limit=26):
    """清掉动作词、分隔符，截到能读的长度。"""
    t = str(s or "").strip()
    for k in _NOISE:
        if t.startswith(k):
            t = t[len(k):]
            break
    t = _SPLIT_RE.sub("-", t).strip("-—.:： ")
    t = re.sub(r"-{2,}", "-", t)
    return t[:limit].rstrip("-—.:： ")


def _corp_in(part):
    """从一个片段里切出单位名，返回 (单位名, 该片段剩余部分)。

    不能只判「片段以公司结尾」：`与「浙江阿里商旅旅行社有限公司」发生的1月-差旅费`
    里单位名夹在中间。也不能只取第一个命中：`中国银行股份有限公司广州分行` 里
    「股份有限公司」在前，会把「广州分行」丢掉。所以：
      1) 强后缀优先（公司/分行/事务所…），弱后缀（银行/中心）兜底；
      2) 每档取**最靠右**的命中。
    """
    core = _BRACKET_TAIL.sub("", str(part or "")).strip()
    core = _LEAD_NOISE.sub("", core).strip()
    if not core:
        return "", ""
    for pool in (_CORP_STRONG, _CORP_WEAK):
        best_end, best_len = -1, 0
        for suf in pool:
            i = core.rfind(suf)
            if i < 2:
                continue
            end = i + len(suf)
            if end > best_end or (end == best_end and len(suf) > best_len):
                best_end, best_len = end, len(suf)
        if best_end >= 0:
            name = core[:best_end]
            tail = _TAIL_NOISE.sub("", core[best_end:]).lstrip("的").strip("-—.:： ")
            return name, tail
    return "", core


def extract_counterparty(summary):
    """从摘要里抽「对方单位」和「业务描述」。

    为什么必须从摘要抽：真实账套里 counterparty / counterparty_name 两列
    **填充率是 0**（实测 10,727 条费用分录全空）。摘要才是唯一有信息的地方。
    """
    txt = str(summary or "").strip()
    if not txt:
        return "", ""
    parts = [p for p in _SPLIT_RE.split(txt) if p]
    cp, cp_idx, cp_tail = "", -1, ""
    for i, p in enumerate(parts):
        name, tail = _corp_in(p)
        if name and len(name) > len(cp):
            cp, cp_idx, cp_tail = name, i, tail
    if not cp:
        return "", _clean_biz(txt)
    biz = _clean_biz("-".join([cp_tail] + parts[cp_idx + 1:]) if cp_tail
                     else "-".join(parts[cp_idx + 1:]))
    if not biz:
        biz = _clean_biz("-".join(parts[:cp_idx]))
    return cp, biz


def _is_close(summary):
    s = str(summary or "")
    return any(k in s for k in SKIP_SUMMARY)


def group_key(summary):
    """把一条摘要归到「业务组」：优先对方单位，其次业务类型。

    为什么必须有这一步：直接按凭证取前两笔，捞出来的常是「结转」或一次性调整
    （真实数据里一笔 1,148 万的中介费结转、一笔 −299 万的审计调整），
    对解释变动没有意义。按业务组汇总之后，才看得出是「哪一类业务」在推动。
    """
    cp, biz = extract_counterparty(summary)
    if cp:
        return cp, biz
    segs = [s for s in _SPLIT_RE.split(biz or "") if s]
    for s in segs:
        if _NOT_BIZ.match(s) or len(s) < 2:
            continue
        return s[:16], biz
    return "其他业务", biz or str(summary or "")[:20]


# ---------------------------------------------------------------- 取数
def group_account(conn, period_from="1", period_to="6", per_acct=6, max_rows=60000):
    """按科目把序时账分组：{科目名: {"total": 金额, "groups": [...]}}。

    groups 按 |金额| 降序，是 AI 写变动原因的**证据**，不是结论。
    """
    from core import recon
    try:
        cond = recon._period_cond(conn, period_from, period_to, ["period"])
    except Exception:
        cond = ""
    where_period = (" AND " + cond) if cond else ""
    # 与 read_expense 同一口径：借 −（贷 − 结转损益贷方）。
    # 不能用「整行剔除」，那会把摘要写着结转、但实际是借方费用确认的分录一起丢掉。
    sql = """
        SELECT acct_name, summary,
               COALESCE(debit, 0)
               - CASE WHEN COALESCE(summary,'') LIKE '%%结转%%' THEN 0
                      ELSE COALESCE(credit, 0) END AS amt,
               period
        FROM journal
        WHERE (COALESCE(acct_l1,'') LIKE '%%销售费用%%'
            OR COALESCE(acct_l1,'') LIKE '%%管理费用%%'
            OR acct_name LIKE '销售费用%%' OR acct_name LIKE '管理费用%%'
            OR acct_name LIKE '%%_销售费用%%' OR acct_name LIKE '%%_管理费用%%')
          %s
    """ % where_period
    try:
        raw = conn.execute(sql).fetchall()
    except Exception:
        return {}

    acc = {}
    for name, summary, amt, period in raw[:max_rows]:
        name = str(name or "").strip()
        if not name:
            continue
        try:
            a = float(amt or 0)
        except (TypeError, ValueError):
            continue
        # 结转损益的贷方已经被上面的 CASE 抵掉了，整笔剩 0 就跳过。
        # 注意这里**不按摘要关键字跳整行**：借方是真实费用时它必须留着。
        if abs(a) < 0.004:
            continue
        d = acc.setdefault(name, {"total": 0.0, "groups": {}})
        d["total"] += a
        key, _biz = group_key(summary)
        g = d["groups"].setdefault(key, {"key": key, "amt": 0.0, "n": 0,
                                         "periods": set(), "sample": ""})
        g["amt"] += a
        g["n"] += 1
        g["periods"].add(str(period or ""))
        if not g["sample"]:
            g["sample"] = str(summary or "").strip()[:50]

    out = {}
    for name, d in acc.items():
        gs = sorted(d["groups"].values(), key=lambda x: -abs(x["amt"]))[:per_acct]
        tot = abs(d["total"]) or 1.0
        out[name] = {
            "total": round(d["total"], 2),
            "groups": [{"key": g["key"], "amt": round(g["amt"], 2),
                        "rows": g["n"], "months": len(g["periods"]),
                        "share": round(abs(g["amt"]) / tot, 4),
                        "sample": g["sample"]} for g in gs],
        }
    return out


# ---------------------------------------------------------------- 版式
def _paint_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD_COLOR)
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _fallback_reason(groups):
    """AI 没参与时只陈述事实，不下结论。"""
    if not groups:
        return "序时账中未找到该科目的明细摘要，无法说明构成"
    top = groups[0]
    return ("本期该科目主要由「%s」构成，金额 %s 元，占该科目 %.2f%%（明细见 evidence 文件）"
            % (top["key"], format(abs(top["amt"]), ",.2f"), top["share"] * 100))


# ---------------------------------------------------------------- 主流程
def build_oar(wb, stats, meta=None, groups=None, te=None, conclusions=None,
              sad=None, chg_amount_min=None, chg_pct_min=0.10, top_n=20,
              prior_bkd=None, cm=None):
    """生成 OAR 页。

    与旧版的区别（V1.2.16）：
      · 分析对象从「科目」变成「OAR 口径」（薪酬/折旧摊销/广告费/其他大额单列/其他）。
        78 个科目逐行排，复核人看不出"薪酬一共多少、比去年多多少"。
      · 变动分析改成叙述式（一句话一段），不再用五列表格——金额和比例写在句子里。

    V1.2.79：**销售费用、管理费用各一张页**（以前两块表并在一页上）。
    本函数只做调度：算一次阈值、清旧页、按类别调 _draw_sheet，再把两页的结果并起来返回。

    stats：_fill_sheet 返回的 [{sheet, rows, sum_row, data_start, last_data}, ...]
    groups：group_account() 的结果（证据，按科目名）
    conclusions：{口径名: AI 写的结论}
    """
    meta = meta or {}
    groups = groups or {}
    # AI 写过的结论兜底：没显式传就用上次落盘的。
    # 不落盘的话，只要重新生成一次（改 TE、补数据、换参数），AI 写的整段分析
    # 就被默认描述顶掉，项目组还以为是 AI 没写。
    _concl = load_conclusions()
    _concl.update(conclusions or {})
    conclusions = _concl

    from . import expense as EXP

    spans = []
    for st in stats:
        if st["sheet"] not in wb.sheetnames:
            continue
        sh = wb[st["sheet"]]
        col = EXP._locate(sh)
        if not col:
            continue
        amt = col.get("amount") or 5
        oar_c = amt + 15 + EXP.FX_SHIFT       # OAR 口径列（金额列 +17）
        _label = "销售费用" if "销售" in st["sheet"] else "管理费用"
        _bk = "VC" if _label == "销售费用" else "VD"
        spans.append({
            "sheet": st["sheet"],
            "label": _label,
            "ws": sh, "col": col,
            "name": get_column_letter(col["name"]),
            "oar": get_column_letter(oar_c),
            "aud": get_column_letter(amt + 5 + EXP.FX_SHIFT),    # 本期审定数（+7）
            "prior": get_column_letter(amt + 8 + EXP.FX_SHIFT),  # 上期审定数（+10）
            "d1": st["data_start"], "d2": st["data_start"] + st["rows"] - 1,
            "sum": st["sum_row"],
            # 上期实体值：BKD 的上期列是公式，读不到；上期页是搬进来的真数
            "priorIdx": _prior_index(wb, (prior_bkd or {}).get(_bk)),
            # BKD 的 OAR 口径列现在是 XLOOKUP 公式，openpyxl 读不到值 ——
            # 生成时算好的口径随 stats 传进来，_buckets_of 优先用它。
            "bucketByRow": st.get("bucketByRow") or {},
            # 客户公司mapping（含余额/损益汇率）：BKD 汇率列是 XLOOKUP 公式，
            # 汇总本期金额时读不到值，按同一份 mapping 再取一次
            "cm": cm,
        })
    if not spans:
        return {"status": "skip", "reason": "BKD 列定位失败，OAR 无法链接"}

    # ---- 阈值：波动幅度取 BKD 里 C22 的口径（TE × 各项认定占比的最小值）----
    wave = _wave_threshold(spans[0]["ws"], te)
    pct_min = chg_pct_min if isinstance(chg_pct_min, (int, float)) else 0.10
    if chg_amount_min is not None and chg_amount_min > 0:
        wave_use = float(chg_amount_min)
    else:
        wave_use = wave

    # 旧页先清：含单页时代的「VC&VD OAR」，重跑不留残页
    for _n in (OAR_SHEET,) + OAR_SHEETS:
        if _n in wb.sheetnames:
            del wb[_n]

    infos = []
    for sp in spans:
        _i = _draw_sheet(wb, sp, wave_use, pct_min, groups, conclusions, top_n)
        if _i:
            infos.append(_i)
    if not infos:
        return {"status": "skip", "reason": "BKD 里没有能归到 OAR 口径的科目"}

    # 两页的结果合并（evidence 导出、AI 回填都按合并后的 candidates 走）。
    # 用各页已拼好的 "candidates"（name 带「类别·口径」、带 subjects），
    # 不要用「_cand」：那是画页时的中间结构，name 只有口径名，
    # 落进 JSON 后 AI 分不清是销售费用的薪酬还是管理费用的薪酬。
    for _i in infos:
        _i.pop("_cand", None)          # 中间结构，不往外传
    cand_all = [c for _i in infos for c in (_i.pop("candidates") or [])]
    pend_all = [p for _i in infos for p in _i.pop("_pending")]
    blocks_all = [b for _i in infos for b in _i.pop("blocks")]
    first = infos[0]
    tot_items = sum(_i["items"] for _i in infos)
    return {"status": "ok",
            "sheet": first["sheet"],            # 兼容旧调用方：默认看第一页（销售费用）
            "sheets": [{"sheet": _i["sheet"], "label": _i["label"],
                        "items": _i["items"], "analyzed": _i["analyzed"]}
                       for _i in infos],
            "items": tot_items, "analyzed": sum(_i["analyzed"] for _i in infos),
            "detailRow": first["detailRow"], "totalRow": first["totalRow"],
            "sectionRow": first["sectionRow"], "firstReasonRow": first["firstReasonRow"],
            "blocks": blocks_all,
            "wave": round(wave_use, 2), "te": te, "sad": sad,
            "threshold": {"sad": sad, "chgMin": wave_use, "pctMin": pct_min},
            "pending": pend_all,
            "candidates": cand_all,
            "_groups": groups,
            "withEvidence": sum(_i["withEvidence"] for _i in infos)}


def _draw_sheet(wb, sp, wave_use, pct_min, groups, conclusions, top_n):
    """画一张 OAR 页：一张页只放一个费用类别（销售费用 或 管理费用）。

    为什么拆：一页里两套「本期/上期/变动/合计」，复核人往下读时很容易把
    销售费用的口径当成管理费用的；而且两张 BKD 的合计要各自核对，
    分页之后「这张页对不对」一眼就能看出来。
    """
    ws = wb.create_sheet(OAR_SHEET_OF[sp["label"]])
    label = sp["label"]

    # ---- 抬头（全部引用 Lead，不复制数字）----
    ws["B1"] = "VC&VD OAR 分析性复核 · %s" % label
    ws["B1"].font = Font(name="Arial", size=14, bold=True)
    for r, _lb, ref in ((2, "客户名称", "C2"), (3, "资产负债表日", "C3"),
                        (4, "可容忍误差(TE)", "C4"), (5, "名义金额(SAD)", "C5")):
        ws.cell(r, 2, _lb).font = Font(name="Arial", size=10, bold=True)
        c = ws.cell(r, 4, "='Uexp.00 Lead'!%s" % ref)
        c.font = Font(name="Arial", size=10)
        c.number_format = "yyyy-mm-dd" if ref == "C3" else MONEY_FMT
    ws.cell(6, 2, "波动幅度阈值（需调查起点）").font = Font(name="Arial", size=10, bold=True)
    c = ws.cell(6, 4, round(wave_use, 2))
    c.font = Font(name="Arial", size=10)
    c.number_format = MONEY_FMT
    ws.cell(6, 5, "= TE × 50%（各项认定均为 Minimal 时的最小值，取自 C22）").font = \
        Font(name="Arial", size=9, italic=True)

    # ---- 一、按口径汇总（本页只放一个类别）----
    heads = ("序号", "OAR 口径", "本期审定数", "上期审定数", "变动金额", "变动率",
             "占本类比重", "是否需调查")
    buckets = _buckets_of(sp)
    if not buckets:
        # 这个类别在 BKD 里没有能归类的科目：不画空页，交给调度方决定
        del wb[ws.title]
        return None
    sec_row = 7
    title = "一、%s · 按 OAR 口径汇总（只取审定数，全部链接 BKD）" % label
    ws.cell(sec_row, 2, title).font = Font(name="Arial", size=11, bold=True)
    hdr = sec_row + 1
    for i, h in enumerate(heads):
        ws.cell(hdr, 2 + i, h)
    _paint_header(ws, hdr, 2, 9)

    row0 = hdr + 1
    for i, b in enumerate(buckets):
        r = row0 + i
        ws.cell(r, 2, i + 1)
        ws.cell(r, 3, b["name"])
        ws.cell(r, 4, "=" + _sumif(sp, "aud", 3, r))
        ws.cell(r, 5, '=IF(COUNTA(\'%s\'!$%s$%d:$%s$%d)=0,"",%s)' % (
            sp["sheet"], sp["prior"], sp["d1"], sp["prior"], sp["d2"],
            _sumif(sp, "prior", 3, r)))
        ws.cell(r, 6, '=IF(E%d="","",D%d-E%d)' % (r, r, r))
        ws.cell(r, 7, '=IF(OR(E%d="",E%d=0),"",F%d/E%d)' % (r, r, r, r))
        ws.cell(r, 8, '=IF($D$%d=0,"",D%d/$D$%d)' % (row0 + len(buckets), r, row0 + len(buckets)))
        ws.cell(r, 9, _survey_formula(r, wave_use, pct_min))

    tot = row0 + len(buckets)
    ws.cell(tot, 2, "合计")
    ws.cell(tot, 3, "全部%s" % label)
    for c in (4, 5, 6):
        L = get_column_letter(c)
        ws.cell(tot, c, '=IF(COUNT(%s%d:%s%d)=0,"",SUM(%s%d:%s%d))' % (
            L, row0, L, tot - 1, L, row0, L, tot - 1))
    ws.cell(tot, 7, '=IF(OR(E%d="",E%d=0),"",F%d/E%d)' % (tot, tot, tot, tot))
    ws.cell(tot, 8, '=IF(D%d=0,"",1)' % tot)

    for r in range(row0, tot + 1):
        for c in range(2, 10):
            cell = ws.cell(r, c)
            cell.font = Font(name="Arial", size=10, bold=(r == tot))
            cell.border = BORDER
            if c in (4, 5, 6):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in (7, 8):
                cell.number_format = PCT_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c == 2:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 15

    sec_row = tot + 2

    # ---- 二、变动分析（叙述式，不用表格）----
    sec = sec_row
    ws.cell(sec, 2, "二、变动分析（按变动金额排序；金额与比例写在句中，原因由 AI 依据序时账撰写）")
    ws.cell(sec, 2).font = Font(name="Arial", size=11, bold=True)

    cand = []
    for i, b in enumerate(buckets):
        src = row0 + i
        chg = b["chg"]
        abschg = abs(chg) if chg is not None else abs(b["now"])
        pct = b["pct"]
        if chg is None:
            need = abs(b["now"]) >= wave_use
        else:
            need = abs(chg) >= wave_use and (pct is None or abs(pct) >= pct_min)
        if not need:
            continue
        cand.append({"label": label, "name": b["name"], "src": src,
                     "now": b["now"], "prior": b["prior"], "chg": chg,
                     "pct": pct, "absChg": abschg,
                     "subjects": b["subjects"]})
    cand.sort(key=lambda x: -x["absChg"])
    cand = cand[:top_n]

    first_an = sec + 1
    if not cand:
        c = ws.cell(first_an, 2, "本期没有达到调查阈值的口径变动，无需逐项分析"
                                 "（阈值 %s 元，见 B6）。" % format(wave_use, ",.2f"))
        c.font = Font(name="Arial", size=10)
        ws.merge_cells(start_row=first_an, start_column=2, end_row=first_an, end_column=9)
        last = first_an
    else:
        for i, c in enumerate(cand):
            r = first_an + i
            key = "%s·%s" % (c["label"], c["name"])
            ai_text = (conclusions.get(key) or conclusions.get(c["name"]))
            ws.cell(r, 2, key).font = Font(name="Arial", size=10, bold=True)
            ws.cell(r, 2).alignment = Alignment(horizontal="left", vertical="top")
            f, _filled = _reason_text(r, c, groups, ai_text, wave_use)
            cell = ws.cell(r, 3, f)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            # 只有真没写出来（无证据）才标黄；工具写出了业务描述就不该留黄底，
            # 否则复核的人会以为这一段还是空白。
            if not ai_text:
                # 还是工具拼的描述：把任务卡挂在这一格（V1.2.85）
                # 打开底稿就能看见该写哪个文件、什么格式、写完怎么生效 ——
                # 换一台电脑、没有任何上下文的 AI 也能照着做完。
                try:
                    from openpyxl.comments import Comment
                    cell.comment = Comment(_reason_task_note(key, len(cand)), "PTW")
                    try:
                        cell.comment.width, cell.comment.height = 460, 320
                    except Exception:
                        pass
                except Exception:
                    pass
            if not _filled:
                cell.fill = GAP_FILL
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
            for cc in range(2, 10):
                ws.cell(r, cc).border = BORDER
            # 行高按**显示出来的字数**估，不能按公式长度：公式里有 TEXT()、
            # 引号、连接符，长度约为显示文字的两倍，按它算出来每一行都空一大截。
            # 前缀（本期 X 千元，较上期增加…）约 45 字，合并 C:I 每行约 53 字。
            if ai_text:
                _tlen = len(str(ai_text))
            else:
                _nt, _ok = _narrate(c, groups)
                _tlen = len(_nt) if _ok else len(AI_TOKEN) + 22
            need = -(-(45 + _tlen) // 53) + 1             # 向上取整 + 1 行余量
            ws.row_dimensions[r].height = max(48, min(165, 15 * need))
        last = first_an + len(cand) - 1

    # ---- 三、说明 ----
    note = last + 2
    ws.cell(note, 2, "三、说明")
    ws.cell(note, 2).font = Font(name="Arial", size=11, bold=True)
    lines = [
        "1. 汇总口径见「%s」BKD 最右侧的「OAR 口径」列：薪酬 / 折旧摊销 / 广告费 / 单列大额 / 其他。"
        "本页用 SUMIF 按这一列汇总，BKD 一改本页自动更新，不要手工改数字。" % sp["sheet"],
        "2. 需调查起点 = 波动幅度阈值 %s 元（= TE × 50%%，取自对应 BKD 的 C22），"
        "变动率另需达到 %d%%。两个条件同时满足才列入分析；"
        "上期无数据时，只看本期金额是否达到该阈值。" % (format(wave_use, ",.2f"),
                                                round(pct_min * 100)),
        "3. 金额与比例是公式引用，BKD 一改本页自动变；结论描述由 AI 依据序时账证据"
        "（对手方、业务性质、发生月份）撰写，写进 bkd_work/oar_concl_result.json 后"
        "重新生成即自动带入，不要手改单元格。",
        "4. 黄底表示该口径既无 AI 结论、序时账里也找不到可归类证据，需人工补写。",
        "5. 描述只讲驱动因素与是否一次性，不重复金额与比例（那是公式），也不逐笔罗列"
        "对方单位——明细可查 BKD 的 OAR 口径列与同目录的 evidence JSON。",
        "6. 本页属分析性复核，**不构成审计结论**。异常波动是否需追加程序，由审计师判断。",
        "7. 销售费用与管理费用各成一页（%s / %s），两页互不引用，交底时按科目分开看。"
        % (OAR_SHEET_VC, OAR_SHEET_VD),
    ]
    for i, t in enumerate(lines):
        c = ws.cell(note + 1 + i, 2, t)
        c.font = Font(name="Arial", size=9)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.merge_cells(start_row=note + 1 + i, start_column=2,
                       end_row=note + 1 + i, end_column=9)

    for k, w in {"A": 3, "B": 5, "C": 24, "D": 16, "E": 16, "F": 16,
                 "G": 11, "H": 12, "I": 18}.items():
        ws.column_dimensions[k].width = w
    ws.sheet_view.showGridLines = False

    return {"status": "ok", "sheet": OAR_SHEET_OF[label], "label": label,
            "items": len(buckets), "analyzed": len(cand),
            "detailRow": row0, "totalRow": tot,
            "sectionRow": sec, "firstReasonRow": first_an,
            "blocks": [{"label": label, "rows": tot - row0,
                        "buckets": [x["name"] for x in buckets]}],
            "_cand": cand, "_pending": [
                ("%s·%s" % (c["label"], c["name"]))
                for c in cand
                if not (conclusions.get("%s·%s" % (c["label"], c["name"]))
                        or conclusions.get(c["name"]))
                and not _narrate(c, groups)[1]],
            "candidates": [{"name": "%s·%s" % (c["label"], c["name"]),
                            "label": c["label"], "bucket": c["name"],
                            "now": c["now"], "prior": c["prior"], "chg": c["chg"],
                            "pct": c["pct"], "isNew": c["chg"] is None,
                            "absChg": c["absChg"],
                            "subjects": list(c.get("subjects") or [])} for c in cand],
            "withEvidence": sum(1 for c in cand if _evidence_of(c, groups))}


def _wave_threshold(bkd_ws, te):
    """BKD C22 的口径：先按各项认定的风险等级算 D 列，再取最小值。

    C22 在 openpyxl 里读到的是公式字符串，不能用——按模板同一套规则重算。
    第 8 行（完整性）用 50/25/15/10%，其余三行用 100/75/50/25%。
    """
    if not isinstance(te, (int, float)) or te <= 0:
        return 0.0
    risk = ("Minimal", "Low", "Moderate", "High")
    vals = []
    for r in (8, 9, 10, 11):
        lv = str(bkd_ws.cell(r, 3).value or "Minimal").strip()
        if lv not in risk:
            continue
        i = risk.index(lv)
        pcts = (0.50, 0.25, 0.15, 0.10) if r == 8 else (1.00, 0.75, 0.50, 0.25)
        vals.append(te * pcts[i])
    return min(vals) if vals else te * 0.5


def _cell_txt(v):
    """单元格文字归一：去首尾空白、全角空格、不可见字符。

    匹配键必须归一——上期页搬进来时科目编码可能带空格，`6601.025 ` 与
    `6601.025` 是同一个科目，直接比会静默匹配不上（上期数全变 0）。
    """
    s = str(v or "").replace("\u3000", " ").replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", s)


def _prior_index(wb, pb):
    """上期 BKD 页 → {(账套名称, 科目名称): 本期未审数}，按行读**实体值**。

    为什么不读 BKD 的「上期未审数」列：那是 SUMIFS 公式，openpyxl 读回来是
    字符串不是数字，结果「有没有上期」永远判成没有 —— 表上明明有上期数、
    变动金额也都算得出来，分析句子里却写着「上期无数据，无法计算变动」，
    同一张纸上两处对着干（这是复核人第一眼就会挑的毛病）。
    上期页是整页搬进来的实体数据，读得到真数，按「账套名称 + 科目名称」一一对上。
    """
    if not isinstance(pb, dict) or pb.get("sheet") is None:
        return {}
    # 键统一用「账套名称 + 科目名称」：与 BKD 的 K 列那个 SUMIFS 完全一致。
    # 不用 pb["values"]（那是按科目编码建的），编码跨年不保证一致，
    # 两张表用不同键会出现"K 列有数、OAR 却判成上期无数据"。
    ns = pb["sheet"]
    if ns not in wb.sheetnames:
        return {}
    ws = wb[ns]
    try:
        last = int(pb.get("last") or 0)
    except (TypeError, ValueError):
        last = 0
    if last <= 0:
        last = ws.max_row
    out = {}
    for r in range(30, last + 1):
        nm = _cell_txt(ws.cell(r, 4).value)      # D 科目名称
        if not nm:
            continue
        v = ws.cell(r, 7).value                  # G 本期未审数（搬入时已是数值）
        if isinstance(v, (int, float)):
            k = (_cell_txt(ws.cell(r, 2).value), nm)
            out[k] = out.get(k, 0.0) + v          # 同名多行时相加，跟 SUMIFS 行为一致
    return out


def _buckets_of(sp):
    """从一张 BKD 里读出口径清单，并汇总本期/上期金额（供挑选分析对象用）。

    金额用 E + F + I 现算，不读 J：J 是公式，openpyxl 拿到的是字符串，
    直接判断会全部当成 0，排序就变成「按科目名字母序」，完全不是要的效果。

    上期同理，但更隐蔽：K/M 列是 SUMIFS + IF 公式，读回来是字符串，
    `hasPrior` 恒为假 → 变动恒为空 → 句子写「上期无数据」而表格里 E/F/G
    三列都有数。所以上期优先按「账套名称 + 科目编码」去上期 BKD 页取实体值，
    取不到才退回读 K 列（只有当 K 被写成静态数时才有用）。
    """
    from . import expense as EXP
    agg = {}
    ws, col = sp["ws"], sp["col"]
    oar_c = col["amount"] + 15 + EXP.FX_SHIFT     # OAR 口径列（金额列 +17）
    amt_c = col["amount"]              # E 本期账面数
    pidx = sp.get("priorIdx") or {}
    for r in range(sp["d1"], sp["d2"] + 1):
        nm = str(ws.cell(r, col["name"]).value or "").strip()
        if not nm:
            continue
        by_row = sp.get("bucketByRow") or {}
        if by_row:
            bucket = str(by_row.get(r) or EXP.OAR_BUCKET_OTHER).strip()
        else:
            bucket = str(ws.cell(r, oar_c).value or "").strip() or EXP.OAR_BUCKET_OTHER
        d = agg.setdefault(bucket, {"name": bucket, "now": 0.0, "prior": 0.0,
                                    "hasPrior": False, "subjects": []})
        # 本期审定数走共用的实体值算法：人民币/未审/审定都是公式，
        # 直接读单元格会拿到字符串，汇总出来全是 0。
        d["now"] += EXP.aud_amount(ws, r, col, sp.get("cm"))
        pv = None
        if pidx:
            pv = pidx.get((_cell_txt(ws.cell(r, 2).value), _cell_txt(nm)))
        if pv is None:
            kv = ws.cell(r, amt_c + EXP.FX_SHIFT + 6).value
            pv = kv if isinstance(kv, (int, float)) else None
        if isinstance(pv, (int, float)):
            d["prior"] += pv
            d["hasPrior"] = True
        if nm not in d["subjects"]:
            d["subjects"].append(nm)
    out = []
    for b in agg.values():
        b["chg"] = (b["now"] - b["prior"]) if b["hasPrior"] else None
        b["pct"] = (b["chg"] / b["prior"]) if (b["hasPrior"] and b["prior"]) else None
        # 两期都是 0 的桶不列：表上多一行「0」只会让人怀疑是不是漏了数，
        # 变动分析还会写出一句「本期 0 千元…已超阈值」这种自相矛盾的话。
        if abs(b["now"]) < 0.005 and abs(b["prior"]) < 0.005:
            continue
        out.append(b)
    out.sort(key=lambda x: -abs(x["now"]))
    return out


def _sumif(sp, col_key, crit_col, r):
    """SUMIF 汇总：按本页 C 列的口径，汇总 BKD 某一列。"""
    return "SUMIF('%s'!$%s$%d:$%s$%d,$%s%d,'%s'!$%s$%d:$%s$%d)" % (
        sp["sheet"], sp["oar"], sp["d1"], sp["oar"], sp["d2"],
        get_column_letter(crit_col), r,
        sp["sheet"], sp[col_key], sp["d1"], sp[col_key], sp["d2"])


def _survey_formula(r, wave, pct_min):
    """是否需调查：金额达阈值且变动率达门槛；上期无数时只看本期金额。

    G（变动率）为空有两种情况：上期整列没数（F 也空）、或上期是 0（F 有数但除不了）。
    这里必须**用嵌套 IF 逐层判空**：Excel 的 OR/AND 不短路，`OR(G="",ABS(G)>=10%)`
    里那个 ABS("") 照样会被算，整格变 #VALUE!（表上出现错误值，复核的人第一反应
    是整张表不可信）。
    """
    return ('=IF(F{r}="",IF(ABS(D{r})>={w},"是（本期达阈值，无上期可比）","否"),'
            'IF(ABS(F{r})<{w},"否（未达阈值）",'
            'IF(G{r}="","是",IF(ABS(G{r})>={p},"是","否（变动率小）"))))'
            ).format(r=r, w=round(wave, 2), p=round(pct_min, 4))


def _squeeze_cn(s):
    """去掉中文之间的空白：摘要里「北 京弘讯和」这类断字很常见，直接显示很怪。"""
    s = str(s or "").replace("\u3000", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return re.sub(r"(?<=[\u4e00-\u9fa5])\s+(?=[\u4e00-\u9fa5])", "", s)


def _evidence_of(c, groups, limit=3):
    """把一个口径下各科目的证据合并，取金额最大的 limit 组。"""
    out = []
    for nm in c.get("subjects") or []:
        g = groups.get(nm) or {}
        for x in (g.get("groups") or []):
            out.append(x)
    out.sort(key=lambda x: -abs(float(x.get("amt") or 0)))
    seen, uniq = set(), []
    for x in out:
        k = str(x.get("key") or "")
        if k in seen:
            continue
        seen.add(k)
        uniq.append(x)
    return uniq[:limit]


#: 业务性质词典（顺序即优先级）。写叙述时用「性质」而不是「公司名 + 金额」——
#: 底稿读者要知道的是「钱花在哪类事上」，不是一串公司和数字。
#: ⚠️ 顺序有关系：「股权激励」必须先于「人力成本」，否则会被「员工」这个词抢走，
#: 把股份支付说成人头工资，那是性质完全不同的两件事。
_NATURE = [
    ("市场投放", ("广告", "推广", "宣传", "投放", "媒体", "展会", "展览", "参展",
                 "google", "谷歌", "百度", "抖音", "腾讯广告", "字节", "小红书",
                 "aigc", "渠道", "品牌")),
    ("股份支付", ("股权激励", "股份支付", "期权", "限制性股票")),
    ("人力成本", ("工资", "薪酬", "提成", "奖金", "社保", "社会保险", "公积金",
                 "福利", "离职", "劳务", "人力", "年终", "津贴", "员工", "招聘")),
    ("中介与专业服务", ("中介", "审计", "律师", "法律", "评估", "咨询", "顾问",
                      "保荐", "承销", "上市", "ipo", "尽调", "税务", "财务顾问")),
    ("折旧摊销", ("折旧", "摊销")),
    ("技术服务", ("技术", "云服务", "软件", "开发", "运维", "系统", "saas",
                 "许可", "license", "信息服务", "数据服务")),
    ("差旅交通", ("差旅", "交通", "住宿", "机票", "车费", "油费", "市内交通")),
    ("办公及租赁", ("办公", "租金", "水电", "物业", "房租", "物管", "耗材")),
    ("业务招待", ("招待", "餐费", "礼品", "业务活动")),
    ("会议培训", ("会议", "培训", "会务", "研讨")),
]

#: 一次性事项。只收**强信号**：
#: · 不带「调整」两字——「调整投放策略」这类正常摘要会被误判；
#: · 不收「暂估」——它多是月度计提的入账标记（真实数据里「计提-某某-5月-…-暂估」），
#:   不代表事项性质，写进变动原因里是噪音。
_ONCE_WORDS = ("上市", "ipo", "股权激励", "股份支付", "减值", "重组", "清算",
               "诉讼", "赔偿", "罚款", "审计调整", "以前年度", "追溯调整")


def _nature_of(x):
    """证据 → 业务性质。认不出就退回「其他」。"""
    t = ("%s %s" % (x.get("key") or "", x.get("sample") or "")).lower()
    for nat, words in _NATURE:
        for w in words:
            if w in t:
                return nat
    return "其他"


def _q(v):
    """金额 → 千元整数串（叙述里统一按千元说，跟表格口径一致）。"""
    try:
        return "{:,.0f}".format(float(v or 0) / 1000.0)
    except (TypeError, ValueError):
        return "0"


def _share_word(s):
    """占比 → 区间词。不用精确百分比：这句是生成时的快照，写死了数据一变就错。

    ⚠️ 只在**没有具体金额**可用时兜底。能给金额就别给"几成"——
    底稿读者要的是"1,870 千元"，不是"六成左右"。
    """
    if s >= 0.95:
        return "近乎全部"
    if s >= 0.85:
        return "八成以上"
    if s >= 0.70:
        return "七成以上"
    if s >= 0.55:
        return "六成左右"
    if s >= 0.40:
        return "四成左右"
    if s >= 0.25:
        return "约四分之一"
    return ""


def _once_kind(word):
    """一次性事项的说法，按命中词分。不要笼统写「一次性」——「上市费用」和
    「股权激励」对复核人的含义完全不同。"""
    if word in ("上市", "ipo", "保荐", "承销"):
        return "上市相关一次性支出"
    if word in ("股权激励", "股份支付"):
        return "股份支付确认"
    if word in ("减值",):
        return "减值计提"
    if word in ("暂估", "调整"):
        return "暂估或调整性质"
    if word in ("重组", "清算"):
        return "重组相关支出"
    if word in ("诉讼", "赔偿", "罚款"):
        return "争议事项支出"
    return "一次性事项"


def _narrate(c, groups):
    """按证据写一段业务描述，替代「公司名 + 金额」的罗列。

    能写的只有事实：钱花在哪类业务、集不集中、含不含一次性事项、规模往哪走。
    「为什么」是审计判断，工具不编——留给人写，但不留占位符占着位置。

    返回 (文字, 是否写出来了)。
    """
    ev = _evidence_of(c, groups, limit=20)
    if not ev:
        return "", False

    total = sum(abs(float(x.get("amt") or 0)) for x in ev) or 1.0
    # 占比的分母必须是「本口径本期审定数」：证据只取了各科目金额最大的几组，
    # 拿证据之和当分母会把 30% 算成 80%，写进底稿就是错的。
    base = abs(float(c.get("now") or 0)) or total
    by = {}
    for x in ev:
        nat = _nature_of(x)
        by[nat] = by.get(nat, 0.0) + abs(float(x.get("amt") or 0))
    ranked = sorted(by.items(), key=lambda kv: -kv[1])
    top, top_amt = ranked[0]
    s1 = top_amt / base
    second = ranked[1][0] if len(ranked) > 1 else ""
    third = ranked[2][0] if len(ranked) > 2 else ""
    name1 = _squeeze_cn(ev[0].get("key") or "")
    top1 = abs(float(ev[0].get("amt") or 0)) / base
    n_ev = len(ev)
    # 证据只取了各科目金额最大的几组，覆盖不全时说「全部」就是假的
    cover = total / base

    # ① 性质：钱花在哪类业务上
    skip_name = False
    if top == "其他":
        # 认不出业务类别时不要硬编一个「其他」，如实说认不出
        nat_txt = "主要由「%s」等 %d 项构成，业务类别需人工归类" % (name1, n_ev)
        skip_name = True
    elif not second or second == "其他":
        if top1 >= 0.95:
            # 一项占满整个口径：性质和名字合成一句，别再补「其中某某占近乎全部」
            nat_txt = ("全部为%s（即「%s」）" % (top, name1)) if name1 and name1 != top \
                else "全部为%s" % top
            skip_name = True
        else:
            nat_txt = ("全部为%s" if cover >= 0.85 else "主要集中于%s") % top
    elif s1 >= 0.60:
        nat_txt = "以%s为主，另有%s" % (top, second)
    elif s1 >= 0.40:
        nat_txt = "以%s为主，兼有%s%s" % (top, second, ("、" + third) if third else "")
    else:
        nat_txt = "由%s、%s等 %d 类业务构成，集中度不高" % (top, second, len(ranked))

    # ② 单家集中度：明显集中就点名，并给出**千元金额**（不再用「六成左右」这种模糊说法）
    one_txt = ""
    if not skip_name:
        if top1 >= 0.45:
            one_txt = "，其中「%s」%s 千元（占 %.1f%%）" % (
                name1, _q(abs(float(ev[0].get("amt") or 0))), top1 * 100)
        elif n_ev >= 3:
            one_txt = "，涉及 %d 家（项）主要对手，无单家过半" % n_ev
    head = "本期" + nat_txt + one_txt

    # ③ 一次性事项：对变动解释力最强，但也最容易误报。
    # 两个门槛：占比要够大（否则是噪音）、性质不能和①重复（否则「股份支付…属股份支付确认」）
    once, once_kind, once_amt = [], "", 0.0
    for x in ev:
        t = "%s %s" % (x.get("key") or "", x.get("sample") or "")
        hit = [w for w in _ONCE_WORDS if w in t.lower()]
        if hit:
            once.append(_squeeze_cn(x.get("key") or ""))
            once_amt += abs(float(x.get("amt") or 0))
            if not once_kind:
                once_kind = _once_kind(hit[0])
    if once and once_amt / base >= 0.20 and once_kind and top not in once_kind:
        uniq_once = []
        for o in once:
            if o and o not in uniq_once:
                uniq_once.append(o)
        if uniq_once and uniq_once[0] == name1:
            head += "；该项属%s，宜确认是否已完整入账、后期是否还会发生" % once_kind
        else:
            head += "；其中「%s」属%s，宜确认是否已完整入账" % ("」「".join(uniq_once[:2]), once_kind)

    # ④ 规模方向
    pct = c.get("pct")
    if pct is not None:
        a = abs(pct)
        word = ("显著" if a >= 0.5 else ("明显" if a >= 0.3 else ("有所" if a >= 0.1 else "")))
        move = "收缩" if pct < 0 else "扩张"
        if not word:
            head += "；整体规模与上期基本持平"
        else:
            head += "；整体规模较上期%s%s" % (word, move)
    return head + "。", True


def ev_txt_safe(c, groups):
    """证据文字（用于估行高，不参与公式）。"""
    return "、".join(str(x.get("key") or "") for x in _evidence_of(c, groups))


def _reason_text(r, c, groups, ai_text, wave):
    """拼「金额/比例（公式）+ 业务描述（静态）」那一句话。

    ⚠️ 列引用必须指对自己的列：C=口径、D=本期、E=上期、F=变动、G=变动率、H=占比。
    旧版把「科目名称」当成金额套进 TEXT()，生成的句子是
    「金额 销售费用_宣传推广费 元」——用户看到的就是这堆"乱码"。

    描述部分不再罗列「公司名 + 金额」（那是把明细表抄了一遍，不是分析），
    改按业务性质归类后写构成特点；写得出就写，写不出才留占位。
    返回 (公式, 是否已写出内容)。
    """
    s = c["src"]
    if ai_text:
        tail, filled = str(ai_text), True
    else:
        note, ok = _narrate(c, groups)
        tail = note if ok else (AI_TOKEN + " 请结合上方构成与序时账摘要撰写。")
        filled = ok

    # 句式必须**由公式自己判**，不能用生成时的 chg 定死结构：
    # 生成时上期列还是公式（读不到数），写死就会产出「上期无数据，无法计算变动」，
    # 而打开底稿重算后 E/F/G 三列全有数 —— 同一张纸上两处互相打脸。
    # 这里两种情形都写进 IF，重算后自动选对的那一句。
    head = ('="「{n}」本期 "&TEXT($D{s}/1000,"#,##0")&" 千元"'
            '&IF($E{s}="","；上期无数据，无法计算变动，本期按金额达阈值入选。",'
            'IF(ABS($F{s})<0.5,"，与上期基本持平。",'
            '"，较上期"&IF($F{s}>0,"增加 ","减少 ")&TEXT(ABS($F{s})/1000,"#,##0")'
            '&" 千元（"&IF($G{s}="","—",TEXT(ABS($G{s}),"0.0%"))&"）。"))'
            '&"占本类 "&TEXT($H{s},"0.0%")&"。"'
            ).format(n=c["name"], s=s)
    return head + '&"%s"' % str(tail).replace('"', "'"), filled


def export_evidence(info, meta, path):
    """把「变动 + 分组证据」导出成 JSON，交给 AI 写结论。

    工具算得快、AI 判得准，两边各干各的：工具不算「为什么」，AI 不重算金额。
    """
    payload = {
        "kind": "expense-oar-evidence",
        "version": 1,
        "meta": meta or {},
        "threshold": info.get("threshold") or {},
        "instruction": (
            "请为每个科目写一句中文结论说明变动原因，要求：\n"
            "1) 指出主要由哪一类业务 / 哪家公司推动（用 groups 的 key、amt、share）；\n"
            "2) 只陈述证据支持的事实，不要编造未出现过的业务；\n"
            "3) 一句话，40~80 字，不要以科目名开头（表格里已有科目列）；\n"
            "4) 输出 JSON：{\"科目名称\": \"结论文字\"}"
        ),
        "accounts": [],
    }
    gp = info.get("_groups") or {}
    for c in info.get("candidates") or []:
        payload["accounts"].append({
            "name": c["name"],
            "now": round(c["now"], 2),
            "prior": None if c["prior"] is None else round(c["prior"], 2),
            "change": None if c["chg"] is None else round(c["chg"], 2),
            "changePct": None if c["pct"] is None else round(c["pct"], 4),
            "isNew": c["chg"] is None,
            "shareOfSelectedChange": c.get("shareOfChg"),
            "subjects": list(c.get("subjects") or []),
            "groups": _cand_evidence(c, gp, limit=15),
        })
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
    return path


def _concl_path():
    """AI 写过的 OAR 分析文字落在这里，下次重新生成自动带上。

    与 BKD 表2（table2_concl.json）同一套做法：不落盘的话，只要重新生成一次
    （改 TE、补数据、换参数），AI 写的分析就被「本期以人力成本为主…」这类
    默认描述顶掉，项目组还以为是 AI 没写。
    """
    base = os.environ.get("KZ_DATA_DIR")
    if base:
        return os.path.join(base, "bkd_work", "oar_concl.json")
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(os.path.normpath(os.path.join(here, "..", "..", "..",
                                                      "data", "bkd_work")), "oar_concl.json")


def _label_of(sheet):
    return "销售费用" if "销售" in str(sheet or "") else "管理费用"


def _reason_task_note(key, n_items):
    """OAR 分析段还没 AI 文字时，挂在该格批注上的任务卡。"""
    try:
        from .tasks import result_path
        p = _concl_path()
        task_json = os.path.join(os.path.dirname(p), "ai_tasks", "oar_reason_pending.json")
        res = result_path(p)
    except Exception:
        task_json, res = "", ""
    label, _, bucket = str(key).partition("·")
    return "\n".join([
        "【OAR 变动分析 · 待补任务】",
        "口径：%s" % (key,),
        "",
        "一、读什么",
        "  %s" % task_json,
        "  本次共 %d 个口径待写，每个 item 带本期/上期金额、变动额、明细构成与对手方。" % n_items,
        "",
        "二、写什么（路径与格式都不能改）",
        "  %s" % res,
        '  {"items":[{"sheet":"VC&VD OAR %s","bucket":"%s","text":"结论文字"}]}'
        % (label or "销售费用", bucket or key),
        "",
        "三、怎么写",
        "  40~80 字，只写原因与性质：由哪类业务、哪家公司推动，是否一次性；",
        "  金额、变动率、是否需调查都由公式给，不要写进文字；",
        "  evidence 里没有的事实不要补；写不出就写「需向客户了解…」并说明缺什么。",
        "",
        "四、收尾",
        "  重新生成一次费用底稿，本段会自动换成你写的文字；不要手改 Excel。",
    ])


def load_conclusions():
    """读回 AI 写过的分析文字：{"销售费用·广告费": 文字, "广告费": 文字}。

    双键是为了 callers 少踩坑：分析段按「类别·口径」找，别的场景可能只给口径名。

    同时合并 `oar_concl_result.json`：任务卡让外部的 AI 只写结果侧文件，
    不动正本（正本里可能有上一年人工改过的段落）。
    """
    p = _concl_path()
    docs = []
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as fh:
                docs.append(json.load(fh))
        except Exception:
            pass
    # AI 按任务卡写的侧文件（正本不动，合并时侧文件优先——它是最新一次判断）
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
            txt = str(it.get("text") or "").strip()
            bucket = str(it.get("bucket") or "").strip()
            if not txt or not bucket:
                continue
            label = str(it.get("label") or "").strip() or _label_of(it.get("sheet"))
            out["%s·%s" % (label, bucket)] = txt
            out.setdefault(bucket, txt)
    return out


def save_conclusions(raw_items):
    """把 AI 写的分析合并落盘（同键覆盖）。

    raw_items 两种写法都收：
        [{"sheet": "VC&VD OAR 销售费用", "bucket": "广告费", "text": "..."}]
        {"销售费用·广告费": "文字", "广告费": "文字"}
    """
    p = _concl_path()
    old = load_conclusions()
    bucket_of = {}
    for k, v in old.items():
        if "·" in k:
            label, b = k.split("·", 1)
            bucket_of[(label, b)] = v
    if isinstance(raw_items, dict):
        raw_items = [{"bucket": k.split("·")[-1],
                      "label": (k.split("·")[0] if "·" in k else ""),
                      "text": v} for k, v in raw_items.items()]
    for it in raw_items or []:
        if not isinstance(it, dict):
            continue
        b = str(it.get("bucket") or it.get("name") or "").strip()
        txt = str(it.get("text") or it.get("reason") or it.get("结论") or "").strip()
        if not b or not txt:
            continue
        label = str(it.get("label") or "").strip()
        if not label and "·" in b:
            label, b = b.split("·", 1)
        label = label or _label_of(it.get("sheet"))
        bucket_of[(label, b)] = txt
    os.makedirs(os.path.dirname(p), exist_ok=True)
    items = [{"label": k[0], "bucket": k[1],
              "sheet": OAR_SHEET_OF.get(k[0], ""), "text": v}
             for k, v in bucket_of.items()]
    with open(p, "w", encoding="utf-8") as fh:
        json.dump({"items": items}, fh, ensure_ascii=False, indent=1)
    return {"total": len(items), "path": p}


def _cand_evidence(c, gp, limit=12):
    """按口径下的科目集合合并证据，取金额最大的 limit 组（去重）。

    ⚠️ 不能用 c["name"] 去 gp 里查：gp 的键是**科目名**
    （销售费用_宣传推广费），c["name"] 是**口径名**（销售费用·广告费），
    永远查不到 —— 以前导出的证据就一直是个空数组，AI 只能写空话。
    """
    ev = []
    for sub in (c.get("subjects") or []):
        ev.extend(((gp.get(sub) or {}).get("groups") or []))
    ev.sort(key=lambda x: -abs(float(x.get("amt") or 0)))
    seen, picked = set(), []
    for e in ev:
        k = str(e.get("key") or "")
        if not k or k in seen:
            continue
        seen.add(k)
        picked.append(e)
        if len(picked) >= limit:
            break
    return picked


def ai_reason_items(info, limit=12):
    """把 OAR 的候选口径 + 序时账证据摊平成 AI 可直接写的清单。

    为什么要摊平：BKD 表2 的分析之所以写得专业，是因为丢给 AI 的不只是
    「本期多少、上期多少」，还有**明细构成与对手方**（谁、什么性质、多少钱）。
    OAR 这边以前只给了金额，AI 只能写空话，所以默认描述才一直是
    「本期以人力成本为主」这种模板句。
    """
    gp = info.get("_groups") or {}
    out = []
    for c in info.get("candidates") or []:
        nm = str(c.get("name") or "")
        label, _, bucket = nm.partition("·")
        if not bucket:
            label, bucket = "", nm
        label = str(c.get("label") or label)
        bucket = str(c.get("bucket") or bucket)
        picked = [{kk: e.get(kk) for kk in
                   ("key", "nature", "sample", "amt", "share", "count", "n")
                   if e.get(kk) is not None}
                  for e in _cand_evidence(c, gp, limit)]
        out.append({
            "label": label, "bucket": bucket,
            "now": None if c.get("now") is None else round(float(c["now"]), 2),
            "prior": None if c.get("prior") is None else round(float(c["prior"]), 2),
            "change": None if c.get("chg") is None else round(float(c["chg"]), 2),
            "changePct": None if c.get("pct") is None else round(float(c["pct"]), 4),
            "isNew": bool(c.get("isNew")),
            "subjects": list(c.get("subjects") or []),
            "evidence": picked,
        })
    return out


def apply_conclusions(data, conclusions, mark_white=True):
    """把 AI 写的结论注入已生成的底稿，返回 (新的 xlsx bytes, info)。

    只替换占位符那一段文字，**公式其余部分原样保留** ——
    这样「本期较上期增加 X 元（Y%）」仍然是活的引用，不会因为补了文字就变死数。
    """
    import io as _io
    from openpyxl import load_workbook

    wb = load_workbook(_io.BytesIO(data))
    targets = [n for n in OAR_SHEETS if n in wb.sheetnames]
    if not targets and OAR_SHEET in wb.sheetnames:      # 旧版单页底稿也认
        targets = [OAR_SHEET]
    if not targets:
        return data, {"applied": 0, "reason": "底稿里没有 OAR 页"}

    applied, missing = [], []
    for _name in targets:
        ws = wb[_name]
        for r in range(1, ws.max_row + 1):
            # 分析段：B 列「费用类别·口径」，C 列整句公式（V1.2.16 起不是五列表格了）
            nm = ws.cell(r, 2).value
            f = ws.cell(r, 3).value
            if not isinstance(nm, str) or "·" not in nm or "按 OAR 口径汇总" in nm:
                continue
            # 公式一定以 = 开头且带 TEXT()；不能只认 IF()——上期无数时整句里没有 IF
            if not isinstance(f, str) or not f.startswith("=") or "TEXT(" not in f:
                continue
            txt = conclusions.get(nm)
            if not txt:
                txt = conclusions.get(nm.split("·")[-1])      # 也接受只写口径名
            if not isinstance(txt, str) or not txt.strip():
                continue
            safe = txt.strip().replace('"', "'")
            if AI_TOKEN in f:
                new = re.sub(re.escape(AI_TOKEN) + r'[^"]*', safe, f, count=1)
            else:
                # 已有旧结论：把最后那段文字换掉
                new = re.sub(r'&"[^"]*"$', '&"%s"' % safe, f)
            ws.cell(r, 3).value = new
            if mark_white:
                for cc in range(3, 10):
                    ws.cell(r, cc).fill = PatternFill(fill_type=None)
                ws.row_dimensions[r].height = max(48, min(150, 15 * (len(safe) // 52 + 3)))
            applied.append(nm)

    for nm in conclusions:
        # 只报「类别·口径」形式的键：load_conclusions 会额外给一份只含口径名的
        # 双键，那份天然匹配不上（同名口径两张页都有），报出来全是噪音。
        if "·" in str(nm) and nm not in applied:
            missing.append(nm)

    bio = _io.BytesIO()
    wb.save(bio)
    return bio.getvalue(), {"applied": len(applied), "accounts": applied,
                            "unmatched": missing}
