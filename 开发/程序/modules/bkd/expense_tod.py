# -*- coding: utf-8 -*-
"""费用 TOD 年审（VC&VD.01.2 TOD 年审）+ TOD 样本池。

为什么单独一份池子
------------------
截止性测试（01.4）只看资产负债表日前后 30 天，TOD 年审要看**全年**。两者筛选
条件不同但底料完全一样（同一份序时账、同一套排除口径），所以 TOD 池就是截止性
池复制一份再补四列（TOD 分类 / 是否关键项 / 抽样ID / 说明）—— 数据只有一份，
改口径两边同时生效，不会出现"截止性说 2658 行、TOD 说 2700 行"。

TOD 分类与模板「1.2 剔除」的对应
--------------------------------
模板把剔除分两段：a 段（19-26 行）是在其他底稿已执行程序的项目，b 段（27 行）
是负值。**一笔只能进一个桶**，否则 G29 合计会重复计算。所以：

* 结转损益 → 不进任何桶（它不是费用发生，是期末结账）
* 负数 → 一律进「负值」（包括预提冲销里的负数）
* 正数再按性质分：折旧/摊销/薪酬/租赁/预提

这样 G27（负值）= 表2 的 D82（冲销预提中的负值）+ D83（其他负值），
模板自带的核对行 `D85 = D84 - G27` 才能归零。这是这份表最容易做错的地方。
"""

import glob
import io
import os
from collections import defaultdict

TOD_SHEET = "VC&VD.01.2 TOD 年审"
PRE_SHEET = "VC&VD.01.2 TOD 预审+剩余期间"
TOD_POOL = "样本池-TOD"
SRC_GLOB = "TOD SWP*.xlsx"

# ---- 模板固定行号（改模板要同步这里）----
R_TOT0, R_TOT_N = 12, 3          # 1.1 总体：12-14，F15=SUM
# 1.2 剔除 9 项（19-27）：26 是「股权激励（股份支付）」—— 生成时从模板 26 行
# 插入一行得到，27 才是「其他剔除」，28 是负值。插入后其下所有行号整体 +1。
R_EXCL0, R_EXCL_N = 19, 9
R_EXCL_NEG = 28
R_KEY0 = 36                      # 2 关键项：36-38，40 合计
R_T1_HDR = 56                    # 表1 明细表头
R_T1_0 = 57                      # 表1 明细首行
R_T1_CAP = 14                    # 表1 明细可用行数（57-70）
R_T1_END = 70
R_SEC2 = 78                      # 表2 起始
R_NEG_ACCR, R_NEG_OTHER = 83, 84
R_NEG_SUM, R_NEG_CHK = 85, 86
R_NEG_HDR = 100                  # 负值明细表头（100-101 两行合并）
R_NEG_0 = 102

# ---- 「预审+剩余期间」页：同一套模板的另一版，金额按 预审段/剩余段 拆两列 ----
P_TOT0 = 11                      # 1.1 总体 11-14（F 预审 / G 剩余 / H 合计）
P_EXCL0, P_EXCL_NEG = 19, 28     # 1.2 剔除（同样 9 项，28 是负值）
P_KEY0 = 37                      # 2 关键项（E/F 数量金额预审，G/H 剩余）
P_T1_HDR, P_T1_0, P_T1_CAP, P_T1_END = 60, 61, 13, 73       # 表1：预审阶段样本
P_T12_HDR, P_T12_0, P_T12_CAP, P_T12_END = 83, 84, 13, 96   # 表1.2：年审阶段样本
P_NEG_ACCR, P_NEG_OTHER, P_NEG_SUM, P_NEG_CHK = 110, 111, 112, 113
P_NEG_HDR, P_NEG_0, P_NEG_CAP = 127, 129, 3                 # b.2 负值明细（129-131）

# 1.2 剔除的 8 个桶：行号 → 桶名（顺序与模板一致，别按金额排）
EXCL_ROWS = [
    (19, "固定资产折旧费"),
    (20, "使用权资产折旧费"),
    (21, "无形资产摊销费"),
    (22, "长期待摊费用摊销费"),
    (23, "工资薪酬"),
    (24, "租赁费"),
    (25, "预提费用（含预提和冲销）"),
    (26, "股权激励（股份支付）"),      # 生成时插在「其他剔除」上面（见 _ensure_equity_row）
    (27, "其他剔除"),
]
# F 列「执行审计工作底稿 X-ref」：按桶给对应底稿名（日期在运行时拼）
# 预提费用给「—」：它靠同主体同科目同金额一正一负配对识别，没有哪张底稿能
# 承接它（其他应付款对不上），填个横杠占位，别让项目组按名字去找底稿。
EXCL_XREF = {
    "固定资产折旧费": "固定资产",
    "使用权资产折旧费": "使用权资产",
    "无形资产摊销费": "无形资产",
    "长期待摊费用摊销费": "长期待摊费用",
    "工资薪酬": "应付职工薪酬",
    "租赁费": "—",          # V1.2.80 起不排除，没有可承接的底稿
    "预提费用（含预提和冲销）": "—",
    "股权激励（股份支付）": "股份支付",
    "其他剔除": "—",
}
# 分类从 cutoff 引：一套「是否排除」两处共用，不再维护第二套 TOD 分类
from .expense_cutoff import (KIND_XFER, KIND_NEG, KIND_NONE, KIND_ACCRUAL,
                             KIND_PAY, ALL_KINDS, DEP_KINDS, is_pay_bucket)


def _d8(balance_date):
    return str(balance_date or "").replace("-", "").replace("/", "").strip()


def _base_amt(r):
    """判定金额 = 池子「折合人民币」（原币 × 折算率），**不是原币、也不是本位币**。

    外币主体的记账本位币是美元：一笔 5 万美元、汇率 6.8454，台账「借方金额」是
    5 万（美元），只有折人民币 34.2 万才够 TE×50% 的关键项阈值。
    V1.2.112 之前这里取的是 amt（本位币），香港/新加坡两个主体整片判不出关键项；
    现在与池子 M 列同源：amt_fx × rate。取不到原币/汇率时退回 amt（人民币主体等价）。
    """
    try:
        amt = float(r.get("amt") or 0)
    except (TypeError, ValueError):
        amt = 0.0
    try:
        fx = float(r.get("amt_fx"))
        rt = float(r.get("rate") or 1.0)
        return fx * rt
    except (TypeError, ValueError):
        return amt


def _fmt_amt(x):
    try:
        return "{:,.2f}".format(float(x))
    except (TypeError, ValueError):
        return str(x)


def mark_rows(rows, threshold=275000.0, balance_date=""):
    """给每行打「是否关键项 / 抽样ID / TOD 说明」。分类直接读 cutoff 已算好的 excl。

    关键项只在**本期**里挑（vdate ≤ 资产负债表日）。池子是全量的，含截止性
    测试要看的日后凭证；那些不属于本期总体，放进表1 会让 E35 与 G31 打架。

    **抽样ID 覆盖「剔除项之外的全部本期样本」，不是只给关键项编号**：
    关键项只是其中达阈值的一小部分，代表性样本要从同一批编号里挑，
    只给关键项编号等于抽样总体没有名册。编号按台账顺序（日期→主体→科目→凭证），
    所以池子从上往下读编号是连续的，抽到 TOD-0137 直接能在池子里找到第 137 个样本。
    """
    import datetime as _dt

    def _in_period(v):
        if not balance_date:
            return True
        try:
            bal = _dt.date.fromisoformat(str(balance_date)[:10])
        except ValueError:
            return True
        if isinstance(v, _dt.datetime):
            v = v.date()
        if not isinstance(v, _dt.date):
            return True
        return v <= bal

    keys = []
    pool = []
    for r in rows:
        k = r.get("excl") or KIND_NONE
        r["todKey"] = ""
        r["sid"] = ""
        r["todNote"] = ""
        r["rmb"] = _base_amt(r)          # 折合人民币：明细表金额列与阈值判定同一个数
        r["inPeriod"] = _in_period(r.get("vdate"))
        if k == KIND_NONE:
            if r["inPeriod"]:
                pool.append(r)          # 本期剩余样本：稍后统一编号
                _a = abs(_base_amt(r))
                if _a >= threshold:
                    r["todKey"] = "是"
                    keys.append(r)
                    r["todNote"] = "关键项：折合人民币 %s ≥ 阈值 %s" % (
                        _fmt_amt(_a), _fmt_amt(threshold))
                else:
                    r["todNote"] = "未达关键项阈值（折合人民币 %s < %s），进代表性样本池" % (
                        _fmt_amt(_a), _fmt_amt(threshold))
            else:
                r["todNote"] = "资产负债表日后凭证，属截止性测试范围"
        elif k == KIND_XFER:
            r["todNote"] = "期末结转，不构成费用交易"
        elif k == KIND_NEG:
            r["todNote"] = "无对应科目性质的负值，按模板 1.2 b 段单独分析"
        else:
            r["todNote"] = "%s：已在其他底稿执行程序" % k
    # 编号按台账顺序（rows 已按 日期→主体→科目→凭证 排好），不按金额：
    # 金额序会随数据变化整体重排，抽样底稿里的编号就对不上上一版了。
    # **关键项与剩余样本分两套前缀**：关键项全查、代表性样本才做抽样，
    # 混用一串号会让复核人以为 KEY 也是抽出来的，也分不清哪个是必查项。
    #   KEY-001…  达阈值的关键项（全部执行程序，不抽样）
    #   TOD-0001… 关键项之外的本期总体（代表性样本从这里抽）
    keys.sort(key=lambda r: (str(r.get("vdate") or ""), str(r.get("vkey") or "")))
    for i, r in enumerate(keys, start=1):
        r["sid"] = "KEY-%03d" % i
    _n = 0
    for r in pool:
        if r.get("todKey") == "是":
            continue                    # 关键项已用 KEY- 前缀，不占剩余样本号
        _n += 1
        r["sid"] = "TOD-%04d" % _n
    return {
        "keys": len(keys),
        "keyAmount": round(sum(float(r["amt"]) for r in keys), 2),
        # 代表性样本名册的条数（不含关键项）：关键项已用 KEY- 前缀单列
        "poolIds": max(0, len(pool) - len(keys)),
        "kindCount": _cnt(rows),
    }


def _cnt(rows):
    """按「是否排除」分类计数（不再有独立的 TOD 分类列）。"""
    c = defaultdict(int)
    for r in rows:
        c[r.get("excl") or KIND_NONE] += 1
    return dict(c)


# ---------------------------------------------------------------- 复制模板 sheet
def copy_sheet(tpl_dir, dst_wb, title=TOD_SHEET):
    """把 TOD SWP 模板里的 TOD 年审页整页复制进费用底稿。

    openpyxl 没有跨工作簿 copy_worksheet，只能逐格搬 value + style。
    模板页 24 列 × 103 行，搬一遍不到 0.1 秒，不值得为此引入别的库。
    """
    from copy import copy
    from openpyxl import load_workbook

    path = ""
    for p in sorted(glob.glob(os.path.join(tpl_dir, SRC_GLOB))):
        try:
            wb0 = load_workbook(p, read_only=False)
            if title in wb0.sheetnames:
                path = p
                break
            wb0.close()
        except Exception:
            continue
    if not path:
        return None, "找不到 TOD 模板（%s，需含「%s」页）" % (SRC_GLOB, title)

    src = load_workbook(path)
    s = src[title]
    if title in dst_wb.sheetnames:
        del dst_wb[title]
    d = dst_wb.create_sheet(title)

    for row in s.iter_rows():
        for c in row:
            nc = d.cell(c.row, c.column)
            nc.value = c.value
            if c.has_style:
                # 必须复制样式的**值**，不能复制 `_style`（那是 StyleArray，
                # 里面全是 fontId/fillId/borderId 这些索引）。索引只在源工作簿
                # 的样式表里有意义，搬到目标工作簿就指向了别的样式 ——
                # 表现就是「Arial 10 变成微软雅黑 14、细边框变粗边框」，
                # 而且不报错。这是整页格式乱掉的真正原因。
                try:
                    nc.font = copy(c.font)
                    nc.fill = copy(c.fill)
                    nc.border = copy(c.border)
                    nc.alignment = copy(c.alignment)
                    nc.protection = copy(c.protection)
                    nc.number_format = c.number_format
                except Exception:
                    pass
    for mc in s.merged_cells.ranges:
        try:
            d.merge_cells(str(mc))
        except Exception:
            pass
    for k, dim in s.column_dimensions.items():
        try:
            if dim.width:
                d.column_dimensions[k].width = dim.width
        except Exception:
            pass
    for k, dim in s.row_dimensions.items():
        try:
            if dim.height:
                d.row_dimensions[k].height = dim.height
        except Exception:
            pass
    try:
        d.sheet_view.showGridLines = s.sheet_view.showGridLines
    except Exception:
        pass
    try:
        src.close()
    except Exception:
        pass
    return d, ""


#: 上期 TOD → 本期「预审+剩余期间」**预审段** 的取数映射。
#: (本期行, 本期列, 上期行, 上期列)
#: 预审段不能按本期阈值从序时账重算：两期重要性水平不同（TE/CRA 百分比不同），
#: 重算出来的是「用新尺子量旧期间」，与上期底稿上的数字对不上。一律引用上期原值。
#:
#: 两套是因为上期可能用两种模板，列位置不同：
#:   annual（年审模板）：单列，1.1 在 F12/F13、1.2 在 G、2 在 E/F、3 在 E44
#:   split（预审+剩余模板）：三列（预审/剩余/合计），**必须取合计列** ——
#:          上期 630 若拆成 1-3/4-6，它的 F 列只有 1-3，本期预审段要的是 1-6 全量。
PRIOR_LINKS_ANNUAL = [
    # 1.1 总体：上期 F12/F13（销售/管理）→ 本期 F11/F12；F15 合计是 SUM，不写
    (11, 6, 12, 6),
    (12, 6, 13, 6),
    # 1.2 剔除：八个桶 + 负值，上期 G 列 → 本期 G 列（本期 H 列仍是剩余期间，自己算）
    (19, 7, 19, 7), (20, 7, 20, 7), (21, 7, 21, 7), (22, 7, 22, 7),
    (23, 7, 23, 7), (24, 7, 24, 7), (25, 7, 25, 7), (26, 7, 26, 7),
    (27, 7, 27, 7),
    # 2 关键项：上期 E/F（数量/金额）→ 本期 E/F；本期 G/H 是剩余期间，自己算
    (36, 5, 35, 5), (36, 6, 35, 6),
    (37, 5, 36, 5), (37, 6, 36, 6),
    (38, 5, 37, 5), (38, 6, 37, 6),
    (40, 5, 39, 5), (40, 6, 39, 6),
    # 3 的剩余样本总体（E45=G31-F40）是模板公式，不写死链接 ——
    # 它的输入（F15/G29/F40）都已来自上期，自己算出来必然等于上期结果
]

#: split 模板：上期 1.1 在 11/12 行（H=合计），1.2 在 I 列，2 在 36~40 行（I/J=合计）
PRIOR_LINKS_SPLIT = [
    (11, 6, 11, 8),          # 1.1 销售 → H11
    (12, 6, 12, 8),          # 1.1 管理 → H12
    (19, 7, 19, 9), (20, 7, 20, 9), (21, 7, 21, 9), (22, 7, 22, 9),
    (23, 7, 23, 9), (24, 7, 24, 9), (25, 7, 25, 9), (26, 7, 26, 9),
    (27, 7, 27, 9),          # 1.2 八桶 + 负值 → I 列（合计）
    (36, 5, 36, 9), (36, 6, 36, 10),    # 2 关键项 → I/J
    (37, 5, 37, 9), (37, 6, 37, 10),
    (38, 5, 38, 9), (38, 6, 38, 10),
    (40, 5, 40, 9), (40, 6, 40, 10),
]


def prior_links(kind):
    """按上期模板类型取映射表；未知类型退回 annual。"""
    return PRIOR_LINKS_SPLIT if kind == "split" else PRIOR_LINKS_ANNUAL


def _is_tod_sheet(ws):
    """这张表是不是 TOD 页：D 列要有 1.1 的「本期全部费用交易」标签。

    只靠表名会抓到「VC&VD.01.2 详细测试 TOD」那种只有标题的说明页。
    """
    for r in range(1, min(ws.max_row or 0, 45) + 1):
        if "本期全部费用交易" in str(ws.cell(r, 4).value or ""):
            return True
    return False


def tod_kind(ws):
    """年审模板 = annual，预审+剩余模板 = split。认不出返回 ""。

    判定看内容不只看名字：1.1 区（第 8~12 行）里出现「剩余期间」就是分段页，
    它比年审页多一列合计、且 1.1 行号是 11/12 而不是 12/13。
    """
    if not _is_tod_sheet(ws):
        return ""
    for r in range(8, 14):
        for c in (4, 5, 6, 7, 8):
            t = str(ws.cell(r, c).value or "")
            if "剩余" in t:
                return "split"
    return "annual"


def _pick_prior_tod(wb, want_sheet=None):
    """选上期 TOD 页，返回 (worksheet, kind)；找不到返回 (None, None)。

    优先级：前端指定的 sheet > 精确名「TOD 年审」> 精确名「预审+剩余期间」> 内容探测。
    两页都在时优先年审页（单列、覆盖面最全）；但年审页没数（空壳）时改选分段页。
    """
    if want_sheet:
        if want_sheet in wb.sheetnames:
            k = tod_kind(wb[want_sheet])
            if k:
                return wb[want_sheet], k
        return None, None
    if TOD_SHEET in wb.sheetnames and _is_tod_sheet(wb[TOD_SHEET]):
        return wb[TOD_SHEET], "annual"
    if PRE_SHEET in wb.sheetnames and _is_tod_sheet(wb[PRE_SHEET]):
        return wb[PRE_SHEET], "split"
    for ws in wb.worksheets:
        k = tod_kind(ws)
        if k:
            return ws, k
    return None, None
    return None


def copy_prior_tod_as_sheet(wb, prior_bytes, title=None, want_sheet=None):
    """把上期 TOD 整页搬进本年底稿：**只搬值，格式与颜色全部保留**。

    支持两种模板：年审（单列）与预审+剩余（三列，取合计列时在 fill_pre 里处理）。
    want_sheet 给定时按名字取，供项目组在界面上手动指定。

    为什么只要值：上期的重要性水平与本期不同，公式搬过来会在本年的数据上重算，
    得到「用新阈值重跑旧期间」的假数。项目组要的是上期当时怎么判的原样证据。

    为什么不用 copy_worksheet：那个只能同工作簿内复制，且会把公式一起带过来。
    这里开两份（data_only=True 取值、False 取样式）逐格合并。

    返回 {"sheet": 名称, "priorDate": 上期资产负债表日, "rows": 内容行数}
    """
    from copy import copy as _cp
    from openpyxl import load_workbook

    wv = load_workbook(io.BytesIO(prior_bytes), data_only=True)
    wsv, kind = _pick_prior_tod(wv, want_sheet)
    if wsv is None:
        try:
            wv.close()
        except Exception:
            pass
        if want_sheet:
            return {"error": "指定的工作表「%s」不是可识别的 TOD 页" % want_sheet}
        return {"error": "上期 TOD 底稿里找不到可识别的 TOD 页（年审或预审+剩余期间）"}
    wf = load_workbook(io.BytesIO(prior_bytes), data_only=False)
    wsf = wf[wsv.title]
    src_title = wsv.title

    # 上期资产负债表日：TOD 页本身通常没有这个标签，退到整份文件找（Lead 页一般有）
    prior_date = ""

    def _find_date(ws_v, ws_f):
        for r in range(1, min(ws_f.max_row or 0, 40) + 1):
            for c in range(1, 8):
                if "资产负债表日" in str(ws_f.cell(r, c).value or ""):
                    for cc in range(c + 1, c + 6):
                        v = ws_v.cell(r, cc).value
                        if v not in (None, ""):
                            return str(v)[:10].replace(" 00:00:00", "")
        return ""

    prior_date = _find_date(wsv, wsf)
    if not prior_date:
        for _ws in wv.worksheets:
            prior_date = _find_date(_ws, wf[_ws.title])
            if prior_date:
                break

    title = title or ("上期TOD_%s" % (prior_date.replace("-", "") or "上一期"))
    if title in wb.sheetnames:
        del wb[title]
    out = wb.create_sheet(title)

    maxr = wsf.max_row or 1
    maxc = wsf.max_column or 1
    n_val = 0
    for r in range(1, maxr + 1):
        for c in range(1, maxc + 1):
            src_f = wsf.cell(r, c)
            dst = out.cell(r, c)
            # 值取 data_only 那份：公式格拿到的是计算结果；没有缓存值就是空
            v = wsv.cell(r, c).value
            dst.value = v
            if isinstance(v, (int, float)) and v:
                n_val += 1
            if src_f.has_style:
                dst.font = _cp(src_f.font)
                dst.fill = _cp(src_f.fill)
                dst.border = _cp(src_f.border)
                dst.alignment = _cp(src_f.alignment)
                dst.number_format = src_f.number_format
                dst.protection = _cp(src_f.protection)

    # 一个数都没搬过来 = 源文件是「程序生成、公式无缓存值」的状态，
    # 搬进来会是一张空白页，链接过去也全是 0 —— 必须当场说清楚，不能静默。
    if n_val < 5:
        del wb[title]
        try:
            wv.close()
            wf.close()
        except Exception:
            pass
        return {"error": "上一期 TOD 里没有可用数值（公式没有缓存值）。"
                         "请先用 Excel 打开该文件、保存一次，再重新生成。"}

    # 列宽 / 行高 / 合并 / 网格线 —— 少了这些，整页会挤成一团，颜色也对不上
    for k, d in wsf.column_dimensions.items():
        if d.width:
            out.column_dimensions[k].width = d.width
    for k, d in wsf.row_dimensions.items():
        if d.height:
            out.row_dimensions[k].height = d.height
    for mc in list(wsf.merged_cells.ranges):
        try:
            out.merge_cells(str(mc))
        except Exception:
            pass
    try:
        out.sheet_view.showGridLines = wsf.sheet_view.showGridLines
        out.freeze_panes = wsf.freeze_panes
    except Exception:
        pass

    # 说明不插行（行号必须与上期一致，PRIOR_LINKS 才是对的），写到页签注释与返回信息里
    try:
        from openpyxl.worksheet.properties import WorksheetProperties  # noqa: F401
    except Exception:
        pass

    try:
        wv.close()
        wf.close()
    except Exception:
        pass
    return {"sheet": title, "priorDate": prior_date, "kind": kind,
            "sourceSheet": src_title, "rows": maxr, "cols": maxc}


#: 上一期 BKD 的两张明细表（名字与我们生成的一致）
BKD_SHEETS = (("VC", "VC.00 销售费用BKD"), ("VD", "VD.00 管理费用BKD"))


def copy_prior_bkd_as_sheet(wb, prior_bytes, resolve=None):
    """把上一期 BKD 的两张明细表整页搬进本年底稿（只搬值，格式原样）。

    本期 BKD 的「上期未审数」改成链接这两张表——项目组点过去能看到上期原始底稿，
    而不是一串不知道从哪来的数字。与 TOD 页同一个道理：只搬值，公式搬过来会重算。

    resolve(账簿名) -> 公司全称：上期底稿的「账套名称」是账簿名（【瑞芸上海（成都）】(009)），
    本年 B 列写的是公司全称。不折一次，按「主体 + 科目名称」的 SUMIFS 永远匹配不上，
    上期数会整列为 0。折算只动搬运后的 B 列，原始文件不碰。
    """
    from copy import copy as _cp
    from openpyxl import load_workbook

    wv = load_workbook(io.BytesIO(prior_bytes), data_only=True)
    wf = load_workbook(io.BytesIO(prior_bytes), data_only=False)
    out = {}
    try:
        date = ""
        for ws_f in wf.worksheets:
            for r in range(1, 30):
                for c in range(1, 8):
                    if "资产负债表日" in str(ws_f.cell(r, c).value or ""):
                        ws_v = wv[ws_f.title]
                        for cc in range(c + 1, c + 6):
                            v = ws_v.cell(r, cc).value
                            if v not in (None, ""):
                                date = str(v)[:10].replace(" 00:00:00", "")
                                break
                        if date:
                            break
                if date:
                    break
            if date:
                break

        for key, want in BKD_SHEETS:
            if want not in wv.sheetnames:
                continue
            wsv, wsf = wv[want], wf[want]
            title = "上期BKD_%s_%s" % ((date.replace("-", "") or "上一期"), key)
            if title in wb.sheetnames:
                del wb[title]
            dst = wb.create_sheet(title)
            maxr, maxc = wsf.max_row or 1, wsf.max_column or 1
            for r in range(1, maxr + 1):
                for c in range(1, maxc + 1):
                    s, d = wsf.cell(r, c), dst.cell(r, c)
                    val = wsv.cell(r, c).value
                    if resolve is not None and c == 2 and r >= 30 and val not in (None, ""):
                        try:
                            comp = resolve(str(val).strip())
                        except Exception:
                            comp = ""
                        if comp:
                            val = comp
                    d.value = val
                    if s.has_style:
                        d.font = _cp(s.font)
                        d.fill = _cp(s.fill)
                        d.border = _cp(s.border)
                        d.alignment = _cp(s.alignment)
                        d.number_format = s.number_format
                        d.protection = _cp(s.protection)
            for k, dim in wsf.column_dimensions.items():
                if dim.width:
                    dst.column_dimensions[k].width = dim.width
            for k, dim in wsf.row_dimensions.items():
                if dim.height:
                    dst.row_dimensions[k].height = dim.height
            for mc in list(wsf.merged_cells.ranges):
                try:
                    dst.merge_cells(str(mc))
                except Exception:
                    pass
            # 数据区上界：D 列「合计」行往上，链接 SUMIFS 只查这一段。
            # 客户底稿有的页（如 VD）压根没有「合计」行，表1 下面直接接审计调整表；
            # 那就用「连续 5 行空」当结束，否则调整表会被 SUMIFS 一并求和。
            last = max(30, maxr - 1)
            _hit = False
            for r in range(20, min(maxr, 300) + 1):
                if str(wsf.cell(r, 4).value or "").replace(" ", "").strip() == "合计":
                    last = r - 1
                    _hit = True
                    break
            if not _hit:
                _blank = 0
                for r in range(30, maxr + 1):
                    _n = wsv.cell(r, 4).value
                    _b = wsv.cell(r, 2).value
                    if (_n in (None, "")) and (_b in (None, "")):
                        _blank += 1
                        if _blank >= 5:
                            last = r - 5
                            break
                    else:
                        _blank = 0
            # 顺手把上期各科目的「本期未审数」（G 列）读成数值字典。
            # 键是 (主体, 科目名称) —— 与本期 K 列那个 SUMIFS 的匹配键完全一致。
            # V1.2.88 起统一按科目名称：科目编码跨年不保证一致（客户换系统/重编码），
            # 按编码建键会让「表上有上期数、导出给 AI 的却是空」。
            vals = {}
            for r in range(30, min(last, (wsv.max_row or 0)) + 1):
                comp = str(wsv.cell(r, 2).value or "").strip()
                # 键里的主体必须和搬进新页的 B 列一致（已折成公司全称），
                # 否则本年按公司全称去查，永远查不到，上期数会成为一片空白/0。
                if resolve is not None and comp:
                    try:
                        comp = resolve(comp) or comp
                    except Exception:
                        pass
                nm = str(wsv.cell(r, 4).value or "").strip()
                v = wsv.cell(r, 7).value
                if nm and isinstance(v, (int, float)):
                    vals[(comp, nm)] = vals.get((comp, nm), 0.0) + float(v)
            out[key] = {"sheet": title, "last": last, "values": vals}
        out["priorDate"] = date
        return out
    except Exception as e:
        out["error"] = "%s: %s" % (type(e).__name__, str(e)[:140])
        return out
    finally:
        try:
            wv.close()
            wf.close()
        except Exception:
            pass


def probe_prior_tod_sheets(data):
    """列出上期 TOD 文件里所有工作表，标出哪些是可用的 TOD 页、属于哪种模板。

    为什么要探测：页签常被改名（「TOD」「630底稿」「Sheet3」），按名字猜不可靠。
    让项目组在界面上指哪张，比猜稳；同时把类型回传，生成时知道取哪一列。
    """
    from openpyxl import load_workbook
    if not data:
        return {"sheets": [], "recommended": -1, "error": "未提供文件"}
    try:
        wv = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        return {"sheets": [], "recommended": -1, "error": "打不开：%s" % e}
    try:
        return probe_tod_sheets_wb(wv)
    finally:
        try:
            wv.close()
        except Exception:
            pass


def probe_tod_sheets_wb(wv):
    """probe_prior_tod_sheets 的实体：在已打开的 workbook 上做。

    上一期 TOD 与上年费用 BKD 常在同一册底稿里，界面只让传一个文件，
    后端就得一次解析同时认两类页，所以把这段抽出来单独可调。
    """
    out = []
    for ws in wv.worksheets:
        k = tod_kind(ws)
        item = {"sheet": ws.title, "kind": k or "", "usable": bool(k),
                "rows": 0, "hasValues": 0}
        if k:
            n = 0
            for r in range(1, min(ws.max_row or 0, 120) + 1):
                for c in range(5, 11):
                    v = ws.cell(r, c).value
                    if isinstance(v, (int, float)) and v:
                        n += 1
            item["rows"] = ws.max_row or 0
            item["hasValues"] = n
        out.append(item)

    # 推荐：先年审（单列最省事），再分段，最后任何可用页；都要有数
    rec = -1
    for want in ("annual", "split"):
        for i, s in enumerate(out):
            if s["kind"] == want and s["hasValues"] > 3:
                rec = i
                break
        if rec >= 0:
            break
    if rec < 0:
        for i, s in enumerate(out):
            if s["usable"]:
                rec = i
                break
    return {"sheets": out, "recommended": rec}


def _insert_rows(ws, at, n):
    """插行并搬走下方合并区 / 行高（openpyxl 的 insert_rows 两样都不管）。"""
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
        try:
            ws.merge_cells(start_row=r1 + n, start_column=c1,
                           end_row=r2 + n, end_column=c2)
        except Exception:
            pass
    for (rr, h) in heights:
        try:
            ws.row_dimensions[rr + n].height = h
        except Exception:
            pass
    return len(merges)


#: 「股权激励（股份支付）」在模板里的插入位置：原「其他剔除」那一行
EQUITY_ROW_AT = 26


def _shift_formula_rows(ws, at, n=1):
    """插行后，把公式里**本表**引用的行号 >= at 的整体 +n。

    openpyxl 的 insert_rows 只搬单元格，公式文本原样保留 —— 模板里
    「剩余样本总体 =G31-F39」「=G36+E36」这类引用不会自己跟，
    插一行后全指错行（实测 6 个 #VALUE!，而且更多是不报错、只是数值错的）。
    只改本表引用：带 sheet 前缀的（`'样本池-TOD'!$M$4:$M$6717`）一律不动，
    字符串字面量也不动。
    """
    import re as _re
    if n <= 0:
        return 0
    # 引用：可选的 sheet 前缀 + 单元格或区域
    REF = _re.compile(r"(?:(?:'[^']*'|[A-Za-z_][\w.]*)!)?"
                      r"\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?")
    CELL = _re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")
    STR = _re.compile(r'("(?:[^"]|"")*")')

    def _fix_ref(m):
        s = m.group(0)
        if "!" in s:                 # 跨表引用，行号是别的表的，别动
            return s
        def _one(x):
            row = int(x.group(4))
            return "%s%s%s%d" % (x.group(1), x.group(2), x.group(3),
                                 row + n if row >= at else row)
        return CELL.sub(_one, s)

    def _fix_formula(f):
        parts = STR.split(f)         # 偶数下标是公式本体，奇数是字符串字面量
        for i in range(0, len(parts), 2):
            parts[i] = REF.sub(_fix_ref, parts[i])
        return "".join(parts)

    cnt = 0
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                nv = _fix_formula(c.value)
                if nv != c.value:
                    c.value = nv
                    cnt += 1
    return cnt


def _clone_row_formulas(ws, at):
    """插进来的新行是空行：把下一行（原行，已下移）的公式复制过来，行号改成 at。

    典型就是分段页 1.2 的 I 列 `=G25+H25`（预审+剩余=合计）。不补的话
    I 列合计少这一项，剩余总体跟着差一整块（实测差 385 万）。
    """
    import re as _re
    pat = _re.compile(r"(\$?[A-Z]{1,3}\$?)%d(?![\d])" % (at + 1))
    n = 0
    for c in range(2, (ws.max_column or 0) + 1):
        src = ws.cell(at + 1, c).value
        if isinstance(src, str) and src.startswith("="):
            ws.cell(at, c).value = pat.sub(lambda m: "%s%d" % (m.group(1), at), src)
            n += 1
    return n


def _ensure_equity_row(ws, at=EQUITY_ROW_AT):
    """在「其他剔除」上面插一行「股权激励（股份支付）」。返回插入行数（幂等）。

    为什么插行而不是复用空行：模板 1.2 的 8 个桶是固定行位，股权激励原先并进
    「工资薪酬」，复核时看不出这块（无现金流出的股份支付）被排除了多少、对应哪张
    底稿。插一行后 1.2 各项之和仍等于总体与剩余之差的勾稽关系，只是分项更细。
    插行后其下所有固定行号整体 +1，常量里已经按新布局写好。
    """
    if "股权激励" in str(ws.cell(at, 4).value or ""):
        _fix_excl_total(ws)
        return 0
    _insert_rows(ws, at, 1)
    _shift_formula_rows(ws, at, 1)     # 模板公式不会自己跟，必须重写本表引用的行号
    _clone_row_formulas(ws, at)        # 新行本身是空行，把上一行的「=G+H」这类公式补回来
    _copy_style(ws, at + 1, at, cols=range(2, 20))
    # 保留相邻行的缩进（分段页是 6 个空格），否则新行看着比别的项低一级
    _ref = str(ws.cell(at + 1, 4).value or "")
    _pad = _ref[:len(_ref) - len(_ref.lstrip(" "))]
    ws.cell(at, 4, _pad + "股权激励（股份支付）")
    _fix_excl_total(ws)
    return 1


def _fix_excl_total(ws):
    """「剔除样本合计」的 SUM 范围补到新的负值行。

    模板原来写死 =SUM(G19:G27)（8 项 + 负值）；插了一行之后负值挪到 28，
    公式不会自己扩，合计就少一块负值，1.2 的行内勾稽立刻对不上。
    """
    from openpyxl.utils import get_column_letter as _L
    for r in range(EQUITY_ROW_AT, min(ws.max_row or 0, EQUITY_ROW_AT + 30) + 1):
        if "剔除样本合计" not in str(ws.cell(r, 4).value or ""):
            continue
        for c in (7, 8):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.upper().startswith("=SUM("):
                L = _L(c)
                # 两页 1.2 的起始行都是 19、负值行都是 28（常量已按新布局）
                ws.cell(r, c).value = "=SUM(%s%d:%s%d)" % (
                    L, R_EXCL0, L, R_EXCL_NEG)
        return


def _copy_style(ws, src_row, dst_row, cols=range(2, 20)):
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


# ---------------------------------------------------------------- 填表
def _fmt(cell, fmt):
    """只在模板没设格式（General/空）时补一个；模板已有的数字格式一律沿用。

    模板是甲方给的，会计格式 `_(* #,##0.00_);...` 看着差不多其实不是同一种，
    覆盖它就等于改了模板格式 —— 用户要求「插行、填值，格式不要动」。
    """
    cur = str(cell.number_format or "General")
    if cur in ("General", ""):
        cell.number_format = fmt
    return cell


AMT_FMT = r'_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '


def _period_labels(period_from="1", balance_date="", pre_end=None):
    """回传给 info 的两个期间标签，便于调用方/界面显示。"""
    d1, bal = _split_dates(balance_date, pre_end)
    if d1 is None or bal is None:
        return {}
    try:
        m0 = int(str(period_from or "1"))
    except (TypeError, ValueError):
        m0 = 1
    return {"pre": "%d-%d月" % (m0, d1.month) if m0 != d1.month else "%d月" % d1.month,
            "rest": ("%d-%d月" % (d1.month + 1, bal.month)) if bal.month > d1.month else "—"}


def _relabel_periods(ws, period_from="1", d1=None, bal=None):
    """把「预审期间（如1-9月份）」「剩余期间（如10-12月份)」改成实际月份。

    模板里这两句是**举例**（默认 12-31 年审 + 9 月预审 → 1-9 / 10-12）。
    实际项目常常是 1-6 月预审、7-9 月出报告，照抄模板的「1-9 / 10-12」就会
    与实际执行的期间对不上，复核的人按表头理解抽样区间会理解错。
    按标签文字定位，不写死单元格 —— 模板改版、上期页插过行都还能改到。
    """
    if d1 is None or bal is None:
        return 0
    try:
        m0 = int(str(period_from or "1"))
    except (TypeError, ValueError):
        m0 = 1
    m1 = d1.month
    m2 = bal.month
    pre = ("%d-%d月" % (m0, m1)) if m0 != m1 else ("%d月" % m1)
    rest = ("%d-%d月" % (m1 + 1, m2)) if m2 > m1 else "—"
    n = 0
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if not isinstance(v, str) or len(v) > 40:
                continue
            if "预审期间" in v or "剩余期间" in v:
                head = "预审期间" if "预审期间" in v else "剩余期间"
                c.value = "%s（%s）" % (head, pre if head == "预审期间" else rest)
                n += 1
    return n


def _split_dates(balance_date, pre_end):
    """切成 (预审截止日, 资产负债表日)。

    预审段 = 期间起 .. pre_end 月月末；剩余段 = 次月 1 日 .. 资产负债表日。
    pre_end 缺省取资产负债表日所在月 → 剩余段为空（客户只有一个期间时就是这个样子，
    模板右边那半列会是 0，这是事实不是漏填）。
    """
    import calendar
    import datetime as _dt
    try:
        y, m, d = (int(x) for x in str(balance_date)[:10].split("-"))
        bal = _dt.date(y, m, d)
    except Exception:
        return None, None
    try:
        pm = int(pre_end) if pre_end not in (None, "") else m
    except (TypeError, ValueError):
        pm = m
    if pm >= m:
        return bal, bal
    if pm < 1:
        return bal - _dt.timedelta(days=1), bal
    return _dt.date(y, pm, calendar.monthrange(y, pm)[1]), bal


def _dcond(rng, op, day):
    return ',%s,"%s"&DATE(%d,%d,%d)' % (rng, op, day.year, day.month, day.day)


#: 1.2 的桶 → BKD 科目名称关键词（按实际科目名细化，避免「折旧」把累计折旧也算进来）。
#: BKD 里没有对应科目的桶（其他剔除、负值）不参与核对，写 "="-" 的公式而不是文本。
_CHECK_KW = {
    "固定资产折旧费": "折旧费",
    "使用权资产折旧费": "使用权资产折旧",
    "无形资产摊销费": "无形资产摊销",
    "长期待摊费用摊销费": "长期待摊",
    "工资薪酬": "薪酬",
    "租赁费": "租赁",
    "预提费用（含预提和冲销）": "预提",
}


def _fill_check(ws, wb, rows=None, layout="split"):
    """在每行「合计」右侧加 check 列，与本期 BKD 的未审汇总数核对。

    1.1 直接对 BKD 的合计行（本期未审 = SUM(明细)），销售/管理各对一次，
        总体金额对两者之和。
    1.2 **按池子里实际出现的科目名称精确引用**（V1.2.80 改）：
        以前用 `SUMIFS(...,"*薪酬*")` 这种关键词模糊匹配，只要 BKD 里有一条
        名字里带「薪酬」的科目就会被算进来，而池子里的分桶是按同一个关键词
        推的 —— 看着"勾稽通过"，其实是同一个模糊规则对了两遍，错也一起错。
        现在改成：先取该桶在池子里出现过哪些**科目名称**，逐个精确 SUMIFS 求和，
        两侧用的是同一批名字，差额才真的是差额。

    相等显示 "-"，不等显示差额，任一侧为空也显示 "-"（没有可比数，不是通过）。
    """
    from openpyxl.styles import Font, Alignment
    from openpyxl.comments import Comment
    try:
        from .expense import SHEET_VC, SHEET_VD
    except Exception:
        SHEET_VC, SHEET_VD = "VC.00 销售费用BKD", "VD.00 管理费用BKD"

    # 两页的金额列不同，check 列跟着金额列走：
    #   分段页 1.1 金额 F/G/H → check 写在 I（引用合计 H）；1.2 金额 G/H/I → check 写 J（引用 I）
    #   年审页 1.1 金额只有 F（另有 14 行「与 BKD 核对」总括行，不再重复加）；
    #          1.2 金额 G → check 写 H（引用 G）
    CFG = {
        # tot = (VC 行, VD 行, 合计行)；合计行给 None 表示不写（另有总括核对行代它）
        "split": dict(tot=(11, 12, 15), tot_hdr=10, tot_amt=8, tot_chk=9,
                      excl_amt=9, excl_chk=10, excl_neg=P_EXCL_NEG),
        # 年审 1.1 的 15 行不写 check：14 行「与 BKD 合计核对（应为 0）」等价，
        # 两处都写等于同一件事做两遍。
        "annual": dict(tot=(12, 13, None), tot_hdr=11, tot_amt=6, tot_chk=7,
                       excl_amt=7, excl_chk=8, excl_neg=R_EXCL_NEG),
    }
    cfg = CFG.get(layout) or CFG["split"]

    red = Font(name="宋体", size=9, bold=True, color="C00000")

    def _hdr(row, col, txt):
        c = ws.cell(row, col, txt)
        c.font = red
        c.alignment = Alignment(horizontal="center", vertical="center")

    # 数字列用「会计专用（无货币符号）」，零显示为右对齐的横杠：
    #   _ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_
    # 就是 Excel 单元格格式里的「会计专用」，货币符号选「无」。
    # 值始终是数值，能 SUM、能排序、能条件格式标黄。
    CHK_FMT = r'_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '

    def _diff(ref, other):
        """差额 = 左 − 右，取到分。

        取整到分是为了消掉浮点尾差：两侧加总顺序不同会差 3e-08，
        显示成横杠没问题，但脚本按 `check == 0` 机器核验会误判。
        真实差额（≥0.005）照样显示出来，不掩错。
        """
        return "=ROUND(%s-(%s),2)" % (ref, other)

    def _bkd_total(sheet):
        """BKD 表1「合计」行的**行号**。**不能写死**：VC 是 205、VD 是 150
        （管理费用明细少 55 行），写死会算出「减 0」的假差额。"""
        if sheet not in wb.sheetnames:
            return None
        w = wb[sheet]
        for r in range(20, min(w.max_row or 0, 260) + 1):
            label = str(w.cell(r, 4).value or "").replace(" ", "").replace("\u3000", "").strip()
            if label == "合计" and w.cell(r, 7).value not in (None, ""):
                return r
        return None

    def _bkd_col(sheet, key):
        """按**表头文字**找列号，不写死列位 —— 模板改版时列会整列右移。"""
        from openpyxl.utils import get_column_letter as _L
        if sheet not in wb.sheetnames:
            return None
        w = wb[sheet]
        k = str(key).replace(" ", "").replace("\n", "")
        for r in range(20, 36):
            for c in range(1, min(w.max_column or 0, 32) + 1):
                v = str(w.cell(r, c).value or "").replace(" ", "").replace("\n", "")
                if v and k in v:
                    return c
        return None

    vc_last = _bkd_total(SHEET_VC)
    vd_last = _bkd_total(SHEET_VD)
    from openpyxl.utils import get_column_letter as _CL
    _ta = _CL(cfg["tot_amt"])
    _tk = _CL(cfg["tot_chk"])
    vc_g = "'%s'!$G$%d" % (SHEET_VC, vc_last) if vc_last else None
    vd_g = "'%s'!$G$%d" % (SHEET_VD, vd_last) if vd_last else None
    if cfg["tot"]:
        r1, r2, rt = cfg["tot"]
        _hdr(cfg["tot_hdr"], cfg["tot_chk"], "check")
        for r, other in ((r1, vc_g), (r2, vd_g)):
            c = ws.cell(r, cfg["tot_chk"],
                        _diff("$%s%d" % (_ta, r), other) if other else '="-"')
            c.alignment = Alignment(horizontal="center")
            c.number_format = CHK_FMT
        if rt:
            c = ws.cell(rt, cfg["tot_chk"],
                        _diff("$%s%d" % (_ta, rt), "%s+%s" % (vc_g, vd_g))
                        if (vc_g and vd_g) else '="-"')
            c.alignment = Alignment(horizontal="center")
            c.number_format = CHK_FMT

    _hdr(17, cfg["excl_chk"], "check")
    # 明细范围到「合计行前一行」为止：写死 $210 的话，明细一多就漏算、
    # 明细一少就把合计算两遍，两种错都看不出来。
    from openpyxl.utils import get_column_letter as _L
    _vc_e = (vc_last - 1) if vc_last else 210
    _vd_e = (vd_last - 1) if vd_last else 210
    vc_rng = "'%s'!$G$30:$G$%d" % (SHEET_VC, _vc_e)
    vc_kw = "'%s'!$D$30:$D$%d" % (SHEET_VC, _vc_e)
    vd_rng = "'%s'!$G$30:$G$%d" % (SHEET_VD, _vd_e)
    vd_kw = "'%s'!$D$30:$D$%d" % (SHEET_VD, _vd_e)

    # 工资薪酬专区：BKD 的「OAR 口径」列 + 池子里出现过的工资类口径名
    _vc_o, _vd_o = _bkd_col(SHEET_VC, "OAR口径"), _bkd_col(SHEET_VD, "OAR口径")
    vc_oar = ("'%s'!$%s$30:$%s$%d" % (SHEET_VC, _L(_vc_o), _L(_vc_o), _vc_e)) if _vc_o else None
    vd_oar = ("'%s'!$%s$30:$%s$%d" % (SHEET_VD, _L(_vd_o), _L(_vd_o), _vd_e)) if _vd_o else None
    from .expense_cutoff import EQUITY_WORDS
    _all_pay = sorted({str(x.get("oarBucket") or "").strip()
                       for x in (rows or []) if is_pay_bucket(x.get("oarBucket"))})
    # 股权激励已经从「工资薪酬」拆成独立一项：口径名也要拆开，
    # 否则工资薪酬的 check 里还含着 385 万股权激励，与上一行的剔除金额差一整块。
    equity_buckets = [b for b in _all_pay if any(w in b for w in EQUITY_WORDS)]
    pay_buckets = [b for b in _all_pay if b not in equity_buckets]

    def _summary_only(kind):
        """这个桶是不是**完全靠摘要**判出来的（科目名认不出）。

        长期待摊费用摊销最典型：科目叫「装修费」「办公费」，只有摘要写着
        「摊销--N月--长期待摊费用-XX公司」。BKD 侧按科目名找不到对应金额，
        硬做 check 会显示一个巨大的假差额，所以不核对、写明原因。
        """
        rs = [x for x in (rows or []) if x.get("excl") == kind]
        return bool(rs) and all("摘要「" in str(x.get("note") or "") for x in rs)

    # 桶 → 池子里出现过的科目名称（去重、保序）。
    # 「其他剔除」与「负值」没有对应科目，仍然不适用。
    names_by_kind = {}
    for row in (rows or []):
        k = row.get("excl")
        nm = str(row.get("name") or "").strip()
        if not k or not nm or k in (KIND_NEG, KIND_XFER):
            continue
        lst = names_by_kind.setdefault(k, [])
        if nm not in lst:
            lst.append(nm)

    from .expense_cutoff import KIND_EQUITY
    _ea = _CL(cfg["excl_amt"])
    for r, kind in list(EXCL_ROWS) + [(cfg["excl_neg"], KIND_NEG)]:
        cell = ws.cell(r, cfg["excl_chk"])
        names = names_by_kind.get(kind) or []
        if kind not in (KIND_NEG, KIND_XFER) and _summary_only(kind):
            # 纯摘要判定：BKD 上按科目名取不到对应的钱，check 没有意义
            cell.value = "摘要判定，不核对"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            try:
                cell.comment = Comment(
                    "本项是按**摘要**判定的（科目名认不出，例如「管理费用_装修费」里"
                    "摘要写着「摊销--N月--长期待摊费用-XX公司」）。\n"
                    "BKD 侧没有同名科目可以加总，做差额只会得到一个巨大的假差额，"
                    "所以本行不做 check；如需核对，请到样本池按摘要筛选后与「长期待摊费用」"
                    "底稿的摊销额比对。", "PTW")
            except Exception:
                pass
            continue
        if kind == KIND_EQUITY and vc_oar and vd_oar and equity_buckets:
            # 股权激励：与工资薪酬同一套做法，按「OAR 口径」列整段取数
            parts = []
            for b in equity_buckets:
                q = str(b).replace('"', '""')
                parts.append('SUMIFS(%s,%s,"%s")' % (vc_rng, vc_oar, q))
                parts.append('SUMIFS(%s,%s,"%s")' % (vd_rng, vd_oar, q))
            cell.value = _diff("$%s%d" % (_ea, r), "+".join(parts))
            cell.number_format = CHK_FMT
        elif kind == KIND_PAY and vc_oar and vd_oar and pay_buckets:
            # 工资薪酬（V1.2.82）：**不再逐个科目名匹配**，直接对 BKD 的「OAR 口径」
            # 列按薪酬类口径名整段取数。池子分桶用的是同一份 mapping，两侧同一批科目，
            # 这样差额只可能是金额差，不会掺进分类差。
            parts = []
            for b in pay_buckets:
                q = str(b).replace('"', '""')
                parts.append('SUMIFS(%s,%s,"%s")' % (vc_rng, vc_oar, q))
                parts.append('SUMIFS(%s,%s,"%s")' % (vd_rng, vd_oar, q))
            cell.value = _diff("$%s%d" % (_ea, r), "+".join(parts))
            cell.number_format = CHK_FMT
        elif kind in (KIND_NEG, "其他剔除", "租赁费") or not names:
            cell.value = '="-"'
        elif kind == KIND_ACCRUAL:
            # 预提费用不是「科目维度」的剔除：它靠同主体、同科目、同金额一正一负
            # 配对识别，桶里只装配对成功的两笔，而 BKD 上那个科目是全年余额。
            # 按科目名去对必然差一个数量级（实测差 3,777 万），这种差额不是错误，
            # 是口径不可比 —— 写「-」并在批注里说清楚，不许硬凑。
            cell.value = '="-"'
        else:
            # 每个名称在两张 BKD 各求一次（同一科目可能同时在销售与管理费用）
            parts = []
            for nm in names:
                q = str(nm).replace('"', '""')
                parts.append('SUMIFS(%s,%s,"%s")' % (vc_rng, vc_kw, q))
                parts.append('SUMIFS(%s,%s,"%s")' % (vd_rng, vd_kw, q))
            cell.value = _diff("$%s%d" % (_ea, r), "+".join(parts))
            cell.number_format = CHK_FMT
        cell.alignment = Alignment(horizontal="center")
    try:
        ws.cell(17, 10).comment = Comment(
            "1.2 的 check 按「池子里该桶实际出现的科目名称」逐个精确 SUMIFS 到两张 BKD，"
            "与上方的剔除金额同口径、同一批科目名。\n"
            "不再用 *关键词* 模糊匹配 —— 那样两侧是同一个模糊规则，错也一起错。\n"
            "「-」的三种情况：① 两侧相等；② 该桶不是科目维度（预提费用按配对识别、"
            "负值按方向识别），与 BKD 科目余额天然不可比；③ 本期该桶无科目（如租赁费已不排除）。\n"
            "写「摘要判定，不核对」的行：该类费用是看摘要才认出来的（科目名认不出），"
            "BKD 侧没有同名科目可加总，硬做会得到假差额。", "PTW")
    except Exception:
        pass
    return {"sheet": ws.title, "cols": "I/J"}


def _first_sample_block(ws):
    """找第一张「样本明细」表的表头行与数据区（按表头文字定位，不写死行号）。

    两版模板行号不同（年审版表头 55、预审+剩余版 59），上期底稿还可能插过行，
    所以不能拿常量去套 —— 找 C 列写着「抽样ID」的那一行，往下读到「异常解释」。
    返回 (表头行, 首数据行, 末数据行)；找不到返回 None。
    """
    hdr = None
    for r in range(1, min(ws.max_row or 0, 200) + 1):
        if str(ws.cell(r, 3).value or "").strip() == "抽样ID":
            hdr = r
            break
    if hdr is None:
        return None
    start = hdr + 1
    last = start - 1
    blank = 0
    for r in range(start, (ws.max_row or start) + 1):
        c3 = str(ws.cell(r, 3).value or "").strip()
        if "异常解释" in c3:
            break
        vals = [ws.cell(r, c).value for c in range(3, 20)]
        if all(v in (None, "") for v in vals):
            blank += 1
            if blank >= 2:
                break
            continue
        blank = 0
        last = r
    return (hdr, start, last) if last >= start else None


def _copy_prior_samples(wb, ws, prior_sheet, start, cap, end_row, o_col=15):
    """把上期 TOD 页表1 的样本明细**原样**搬到本期预审段表1。

    以前这里只写一句「见上期页表1，此处不重复列示」。问题在于：底稿是**按本期
    这一页**看和归档的，跳页引用一多，样本、抽样ID、对方单位就在本页面上消失了，
    复核人无法在这张表里核对「预审到底抽了哪几笔」，也无法与上期底稿做增减对比。
    现在逐行搬值（含 L~S 那几列人工记录的 A/B/C 判断），本页自洽。

    返回搬了几行。
    """
    if prior_sheet not in wb.sheetnames:
        return 0
    src = wb[prior_sheet]
    blk = _first_sample_block(src)
    if not blk:
        return 0
    _hdr, s0, s1 = blk
    n = s1 - s0 + 1
    n_ins = max(0, n - cap)
    if n_ins:
        _insert_rows(ws, end_row + 1, n_ins)
    for i in range(n):
        sr, dr = s0 + i, start + i
        if i >= cap:
            _copy_style(ws, start, dr, cols=range(3, 20))
        for c in range(3, 20):
            sv = src.cell(sr, c)
            dc = ws.cell(dr, c)
            dc.value = sv.value
            try:
                if sv.number_format:
                    dc.number_format = sv.number_format
            except Exception:
                pass
        # O 列差异公式按新行号重写（搬过来的是旧行号的引用，会指错行）
        if ws.cell(dr, 10).value not in (None, ""):
            _fmt(ws.cell(dr, o_col, "=J%d-N%d" % (dr, dr)), AMT_FMT)
    return n, n_ins


def _bkd_aud_col(wb, sheet):
    """BKD 页「本期审定数」在第几列 —— 按表头文字找，不写死字母。

    V1.2.110 在「本年账面数」后插了汇率两列，审定数从 J 移到 L；
    1.1 的核对行如果还写死 J，就会去比合计行下面的 A3/Diff（值 2.00），
    表上显示「差 7,529 万」，看着像勾稽炸了。
    """
    if sheet not in wb.sheetnames:
        return None
    ws = wb[sheet]
    for r in range(1, min(ws.max_row or 0, 60) + 1):
        for c in range(1, min(ws.max_column or 0, 40) + 1):
            if "本期审定数" in str(ws.cell(r, c).value or ""):
                return c
    return None


def _write_sample_rows(ws, start, items, cap, end_row, kind_label="关键项", o_col=15):
    """写一张样本明细表（年审页表1 / 预审页表1、表1.2 共用同一套列）。

    超出行数就插行（模板只给 cap 行），插行后 O 列差异公式不会自动跟，统一重写。
    """
    from openpyxl.styles import PatternFill
    from .expense_oar import extract_counterparty

    yellow = PatternFill("solid", fgColor="FFF2CC")
    n_ins = max(0, len(items) - cap)
    if n_ins:
        _insert_rows(ws, end_row + 1, n_ins)
    for i, it in enumerate(items):
        r = start + i
        if i >= cap:
            _copy_style(ws, start, r)
        ws.cell(r, 3, it.get("sid") or "")
        ws.cell(r, 4, kind_label)
        vd = it.get("vdate")
        if hasattr(vd, "date"):
            vd = vd.date()
        _fmt(ws.cell(r, 5, vd), "mm-dd-yy")
        ws.cell(r, 6, it.get("vkey") or it.get("vno") or "")
        ws.cell(r, 7, str(it.get("summary") or "")[:120])
        ws.cell(r, 8, it.get("code") or "")
        ws.cell(r, 9, it.get("name") or "")
        # 金额列写「折合人民币」：与「2 关键项」的 SUMIFS(池!M) 同源，
        # 否则外币主体的明细合计与上面汇总对不上（差一个汇率）。
        _fmt(ws.cell(r, 10, round(float(it.get("rmb", it.get("amt")) or 0), 2)), AMT_FMT)
        try:
            cp, _biz = extract_counterparty(it.get("summary"))
        except Exception:
            cp = ""
        ws.cell(r, 11, cp)
        # 对齐不在这里设：模板原有对齐靠 _copy_style 继承，硬设会把模板格式改掉
        # L~S（12-19）是支持性证据与 A/B/C 判断，人工填
        for cidx in range(12, 20):
            ws.cell(r, cidx).fill = yellow
    for i in range(cap + n_ins):
        r = start + i
        _fmt(ws.cell(r, o_col, "=J%d-N%d" % (r, r)), AMT_FMT)
    return n_ins


def fill(wb, rows, stats=None, balance_date="", threshold=275000.0, pool_last=0):
    """填 TOD 年审页。所有金额都写成 SUMIFS / BKD 链接，改数据跟着变。"""
    from openpyxl.styles import Alignment, PatternFill

    if TOD_SHEET not in wb.sheetnames:
        return {"status": "skip", "reason": "底稿里没有「%s」" % TOD_SHEET}
    ws = wb[TOD_SHEET]
    # 先在「其他剔除」上面补出「股权激励（股份支付）」一行（幂等），
    # 之后整页都按 9 项剔除的行号写（常量已按新布局）。
    _ensure_equity_row(ws)
    d8 = _d8(balance_date)
    # 待填提示只用浅黄（FFF2CC）。纯黄（FFFF00）一屏两三块太抢眼，
    # 打印出来还会吃色，和 BKD 里「待人工补」的黄色也不一致。
    yellow = PatternFill("solid", fgColor="FFF2CC")
    # 池子的数据区：从第 4 行到最后一行。整列引用（$M:$M）能算但慢，
    # 一万多行 × 二十几个 SUMIFS 会明显拖累打开速度，所以给死区间。
    P = "'%s'" % TOD_POOL          # sheet 名含中文和「-」，公式里必须带单引号
    pl = max(int(pool_last or 0), 4)
    rng_m = "%s!$M$4:$M$%d" % (P, pl)
    rng_n = "%s!$N$4:$N$%d" % (P, pl)      # 是否排除（分类，TOD 直接吃这一列）
    rng_p = "%s!$P$4:$P$%d" % (P, pl)      # 是否关键项
    rng_q = "%s!$Q$4:$Q$%d" % (P, pl)      # 抽样ID
    rng_d = "%s!$F$4:$F$%d" % (P, pl)
    rng_e = "%s!$E$4:$E$%d" % (P, pl)      # 归属（销售费用 / 管理费用）
    # 截止性池子是**全量**序时账（截止性要看到资产负债表日之后 30 天的凭证），
    # TOD 年审的总体却必须停在本期。不加上界的话，次年 1 月的大额凭证会被算进
    # 剔除项，G31（剩余总体）就对不上 F15，而且没人看得出差在哪。
    dcond = ""
    try:
        y, m, dd = (int(x) for x in str(balance_date)[:10].split("-"))
        dcond = ',%s,"<="&DATE(%d,%d,%d)' % (rng_d, y, m, dd)
    except Exception:
        dcond = ""

    # ---- 1.1 总体：与分段页同一口径 —— 数字取样本池，再与 BKD 合计核对 ----
    # 统一从池子 SUMIFS 的好处：1.1、1.2 剔除、剩余总体（G31）三处同源，
    # 改分类、改汇率都一起动；直接链 BKD 的话，池子里一旦有分类差异，
    # 「总体 − 剔除」就和 BKD 对不上，而表上又看不出差在哪。
    bysheet = {}
    for st in (stats or []):
        if isinstance(st, dict) and st.get("sheet"):
            bysheet[st["sheet"]] = st
    tot_rows = [
        (12, "销售费用（VC）本期全部费用交易", "销售费用", "VC.00 销售费用BKD"),
        (13, "管理费用（VD）本期全部费用交易", "管理费用", "VD.00 管理费用BKD"),
    ]
    for r, label, l1, sheet in tot_rows:
        ws.cell(r, 4, label)
        ws.cell(r, 5, TOD_POOL)
        _fmt(ws.cell(r, 6, '=SUMIFS(%s,%s,"%s",%s,"<>%s"%s)'
                     % (rng_m, rng_e, l1, rng_n, KIND_XFER, dcond)),
             "#,##0.00;[Red]-#,##0.00")
    # 模板 F15 是 =SUM(F12:F14)；14 行现在做核对行，合计只算 12/13
    ws.cell(R_TOT0 + R_TOT_N, 6).value = "=SUM(F%d:F%d)" % (R_TOT0, R_TOT0 + 1)
    # 核对行（14 行）：总体 − BKD 本期审定合计，应为 0
    _chk = R_TOT0 + 2
    ws.cell(_chk, 4, "与 BKD 合计核对（应为 0）")
    from openpyxl.utils import get_column_letter as _glc
    _bkd_parts, _miss = [], []
    for _s in ("VC.00 销售费用BKD", "VD.00 管理费用BKD"):
        _st = bysheet.get(_s)
        _col = _bkd_aud_col(wb, _s)
        if _st and _st.get("sum_row") and _col:
            _bkd_parts.append("'%s'!%s%d" % (_s, _glc(_col), int(_st["sum_row"])))
        else:
            _miss.append(_s)
    if len(_bkd_parts) == 2:
        ws.cell(_chk, 5, "VC.00 + VD.00 本期审定数")
        _fmt(ws.cell(_chk, 6, "=F%d-(%s)" % (R_TOT0 + R_TOT_N, "+".join(_bkd_parts))),
             "#,##0.00;[Red]-##0.00")
    else:
        # 只找到一张 BKD 时不写半截公式：拿单边去比会显示一个巨大的假差额
        ws.cell(_chk, 5, "BKD 合计行未定位到（%s）" % "、".join(_miss))
        ws.cell(_chk, 6).fill = yellow

    # ---- 1.2 剔除：按 TOD 分类 SUMIFS ----
    for r, kind in EXCL_ROWS:
        ws.cell(r, 5, TOD_POOL)
        tag = EXCL_XREF.get(kind, "")
        # 「—」表示没有对应底稿，别给它拼日期（会变成「—20260930」很难看）
        ws.cell(r, 6, (tag + d8) if (tag and tag != "—" and d8) else tag)
        _fmt(ws.cell(r, 7, '=ROUND(SUMIFS(%s,%s,"%s"%s),2)' % (rng_m, rng_n, kind, dcond)),
             "#,##0.00;[Red]-#,##0.00")
    # b. 负值：性质归完之后剩下的负数，与表2 的 D82+D83 是两回事
    ws.cell(R_EXCL_NEG, 5, TOD_POOL)
    ws.cell(R_EXCL_NEG, 6, "见表2 负值分析")
    _fmt(ws.cell(R_EXCL_NEG, 7, '=ROUND(SUMIFS(%s,%s,"%s"%s),2)' % (rng_m, rng_n, KIND_NEG, dcond)),
         "#,##0.00;[Red]-#,##0.00")
    # 预提和冲销成对，净额必然是 0 —— 只显示一个 0 会让人以为「没有预提费用」，
    # 加批注把两侧金额写清楚（数字随数据变，批注是快照，所以只写口径不写结论）
    try:
        from openpyxl.comments import Comment
        _p = sum(float(r.get("amt") or 0) for r in rows
                 if r.get("excl") == KIND_ACCRUAL and float(r.get("amt") or 0) > 0)
        _n = sum(float(r.get("amt") or 0) for r in rows
                 if r.get("excl") == KIND_ACCRUAL and float(r.get("amt") or 0) < 0)
        for _r, _k in EXCL_ROWS:
            if _k == KIND_ACCRUAL:
                ws.cell(_r, 7).comment = Comment(
                    "本行为净额：预提 %s、冲销 %s，一正一负相抵为 0。\n"
                    "预提与冲销的分项金额见「表2 冲销预提费用」，逐笔见「%s」。"
                    % ("{:,.2f}".format(_p), "{:,.2f}".format(_n), TOD_POOL), "PTW")
                break
    except Exception:
        pass

    # ---- 2 关键项 ----
    ws.cell(R_KEY0, 5, '=COUNTIFS(%s,"是")' % rng_p)
    _fmt(ws.cell(R_KEY0, 6, '=SUMIFS(%s,%s,"是"%s)' % (rng_m, rng_p, dcond)),
         "#,##0.00;[Red]-#,##0.00")
    # 性质特殊 / 其他风险项是职业判断，工具不替项目组下结论，留空标黄
    for r in (R_KEY0 + 1, R_KEY0 + 2):
        for col in (5, 6):
            if ws.cell(r, col).value in (None, ""):
                ws.cell(r, col).fill = yellow

    # 1.2 剔除项的判定链路写进批注：复核的人（和外部 AI）看这一行就知道
    # 「为什么这笔没进抽样总体」，以及上年 TOD 该怎么比对。
    try:
        from openpyxl.comments import Comment
        ws.cell(R_EXCL0 - 2, 4).comment = Comment(
            "排除项判定顺序（命中即停）：\n"
            "1) 结转损益凭证；2) OAR 口径属人工成本 → 工资薪酬；\n"
            "3) 科目名关键词（折旧摊销按资产类型、薪酬类词）；\n"
            "   股权激励/股份支付**单列一行**（无现金流出，底稿是股份支付测算）；\n"
            "4) 科目名认不出时看摘要：含「长期待摊」→ 长期待摊费用摊销费，\n"
            "   含「股份支付/股权激励/限制性股票」→ 工资薪酬；\n"
            "5) 同主体同科目同金额一正一负 → 预提费用；6) 剩余负数 → 负值（进表2，不算排除）。\n"
            "租赁费不默认排除：短期/低价值租赁直接进费用，是否抽查交给抽样。\n\n"
            "上年 TOD 的 1.2 有金额、本年该行为 0 时，先回样本池按摘要找一遍再下结论。",
            "PTW")
    except Exception:
        pass

    # ---- 表1 明细：关键项，金额降序，超 14 行先插行 ----
    keys = [r for r in rows if r.get("todKey") == "是"]
    keys.sort(key=lambda r: -abs(float(r.get("rmb", r.get("amt")) or 0)))
    off = _write_sample_rows(ws, R_T1_0, keys, R_T1_CAP, R_T1_END)

    # ---- 表2 负值分析（行号随插行偏移）----
    o = off
    # 表2 分析的是**全部负值**（不含期末结账），不是「负值」那一个桶：
    # 负数现在按性质散在 a 段各桶（工资薪酬 -184 万、折旧 -212 万…），
    # 只统计「负值」桶会漏掉一大半，审计会以为负值只有 336 万。
    neg_all = ('SUMIFS(%s,%s,"<0"%s)-SUMIFS(%s,%s,"%s",%s,"<0"%s)'
               % (rng_m, rng_m, dcond, rng_m, rng_n, KIND_XFER, rng_m, dcond))
    # D82 冲销预提费用 = 「预提费用」桶里的负数（预提的负方）
    ws.cell(R_NEG_ACCR + o, 4, '=SUMIFS(%s,%s,"%s",%s,"<0"%s)'
            % (rng_m, rng_n, KIND_ACCRUAL, rng_m, dcond))
    # D83 其他性质 = 全部负值 − 冲销预提
    ws.cell(R_NEG_OTHER + o, 4, "=(%s)-D%d" % (neg_all, R_NEG_ACCR + o))
    ws.cell(R_NEG_SUM + o, 4, "=SUM(D%d:D%d)" % (R_NEG_ACCR + o, R_NEG_OTHER + o))
    # 核对行：表2 合计 应当正好等于全部负值 → 0
    ws.cell(R_NEG_CHK + o, 4, "=D%d-(%s)" % (R_NEG_SUM + o, neg_all))
    for r in (R_NEG_ACCR + o, R_NEG_OTHER + o, R_NEG_SUM + o, R_NEG_CHK + o):
        _fmt(ws.cell(r, 4), "#,##0.00;[Red]-#,##0.00")

    # 是否抽样：按绝对值与阈值比；策略本身人工写，留空标黄
    neg_accr = -sum(float(r.get("amt") or 0) for r in rows
                    if r.get("excl") == KIND_ACCRUAL and float(r.get("amt") or 0) < 0
                    and r.get("inPeriod", True))
    neg_all = -sum(float(r.get("amt") or 0) for r in rows
                   if float(r.get("amt") or 0) < 0 and r.get("excl") != KIND_XFER
                   and r.get("inPeriod", True))
    neg_other = neg_all - neg_accr
    ws.cell(R_NEG_ACCR + o, 5, "是" if neg_accr >= threshold else "否")
    ws.cell(R_NEG_OTHER + o, 5, "是" if neg_other >= threshold else "否")
    # b.1 抽样策略（87-91 行 → 插行后整体下移），人工填写。
    # 只给**空白**格标黄：表2 的数据行（冲销预提费用 / 其他性质 / 合计）就在上面
    # 两三行，无差别刷色会把「其他性质」这类已有文字的行盖成黄的。
    for r in range(R_SEC2 + 10 + o, R_SEC2 + 16 + o):
        cell = ws.cell(r, 3)
        if cell.value in (None, ""):
            cell.fill = yellow

    # ---- 负值明细：只列超阈值的（b.2 是"我们对抽出的样本执行以下工作"）----
    negs = [r for r in rows
            if r.get("excl") == KIND_NEG and r.get("inPeriod", True)]
    negs.sort(key=lambda r: float(r.get("amt") or 0))
    negs = [r for r in negs if abs(_base_amt(r)) >= threshold]
    hdr = R_NEG_HDR + o
    base = hdr + 2
    if len(negs) > 3:
        _insert_rows(ws, base + 3, len(negs) - 3)
    for i, it in enumerate(negs):
        r = base + i
        if i >= 3:
            _copy_style(ws, base, r, cols=range(3, 15))
        ws.cell(r, 3, it.get("vkey") or it.get("vno") or "")
        vd = it.get("vdate")
        if hasattr(vd, "date"):
            vd = vd.date()
        c = ws.cell(r, 4, vd)
        _fmt(c, "mm-dd-yy")
        ws.cell(r, 5, str(it.get("summary") or "")[:120])
        ws.cell(r, 6, it.get("code") or "")
        ws.cell(r, 7, it.get("name") or "")
        c = ws.cell(r, 8, round(float(it.get("amt") or 0), 2))
        _fmt(c, "#,##0.00;[Red]-#,##0.00")
        for cidx in (9, 10, 11, 12, 13):
            ws.cell(r, cidx).fill = yellow
    # 与分段页同一套 check：1.2 每一项右侧一列，对 BKD 的同一批科目/口径取数。
    # 年审页金额在 G 列，check 就写 H 列（配置见 _fill_check 的 CFG["annual"]）。
    try:
        _chk = _fill_check(ws, wb, rows, layout="annual")
    except Exception as _e:
        _chk = {"error": "%s: %s" % (type(_e).__name__, str(_e)[:120])}

    return {
        "status": "ok", "sheet": TOD_SHEET,
        "totals": len(tot_rows), "keys": len(keys), "inserted": off,
        "check": _chk,
        "negRows": len(negs), "negAccrual": round(neg_accr, 2),
        "negOther": round(neg_other, 2),
    }



# ---------------------------------------------------------------- 预审+剩余期间
def fill_pre(wb, rows, stats=None, balance_date="", threshold=275000.0,
             pool_last=0, pre_end=None, prior_tod=None, prior_kind=None,
             period_from="1"):
    """填「VC&VD.01.2 TOD 预审+剩余期间」：同一份池子，金额按两个期间拆两列。

    与年审页的区别只有一个：所有金额拆成「预审段 / 剩余段」两列再加总。
    预审段 = 期间起 .. pre_end 月月末，剩余段 = 其后 .. 资产负债表日。
    表1 装预审阶段执行的关键项，表1.2 装年审（剩余期间）阶段的。

    prior_tod 给了就**预审段整体改为引用上期 TOD 页**（本期只算剩余段）。
    原因：两期的重要性水平不同，用本期阈值重算上期期间，得到的数与上期底稿对不上。
    prior_kind 决定取上期哪一列：annual 取单列，split 取合计列（见 prior_links）。
    """
    from openpyxl.styles import Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    if PRE_SHEET not in wb.sheetnames:
        return {"status": "skip", "reason": "底稿里没有「%s」" % PRE_SHEET}
    ws = wb[PRE_SHEET]
    _ensure_equity_row(ws)          # 与年审页一致：1.2 多一行股权激励
    d8 = _d8(balance_date)
    yellow = PatternFill("solid", fgColor="FFF2CC")
    # 预审段是否改为引用上期 TOD：上期的重要性水平与本期不同，重算出来的数
    # 与上期底稿对不上，项目组复核时无法解释。
    use_prior = bool(prior_tod and str(prior_tod) in wb.sheetnames)
    prior_map = {(cr, cc): (pr, pc) for cr, cc, pr, pc in prior_links(prior_kind)}

    def _pre_formula(crow, ccol, fallback=None):
        """预审段的公式：有上期 TOD 就链接它，否则用传入的 SUMIFS/COUNTIFS。"""
        if not use_prior:
            return fallback
        hit = prior_map.get((crow, ccol))
        if not hit:
            return fallback
        return "='%s'!%s%d" % (prior_tod, get_column_letter(hit[1]), hit[0])
    P = "'%s'" % TOD_POOL
    pl = max(int(pool_last or 0), 4)
    rng_m = "%s!$M$4:$M$%d" % (P, pl)      # 折合人民币
    rng_n = "%s!$N$4:$N$%d" % (P, pl)      # 是否排除
    rng_p = "%s!$P$4:$P$%d" % (P, pl)      # 是否关键项
    rng_d = "%s!$F$4:$F$%d" % (P, pl)      # 日期
    rng_e = "%s!$E$4:$E$%d" % (P, pl)      # 归属（销售费用 / 管理费用）

    d1, bal = _split_dates(balance_date, pre_end)
    if d1 is None:
        return {"status": "skip", "reason": "没有资产负债表日，无法切预审/剩余期间"}
    # 表头按实际月份写（模板里的「如1-9月份」只是举例）
    _relabeled = _relabel_periods(ws, period_from, d1, bal)
    cd_bal = _dcond(rng_d, "<=", bal)
    cd_pre = _dcond(rng_d, "<=", d1)
    cd_rest = _dcond(rng_d, ">", d1) + _dcond(rng_d, "<=", bal)

    # ---- 1.1 总体：按归属拆两段。数据来源写池子，因为金额就是按日期从池子里切的。
    # **必须把「结转损益」排除在外**：期末结转是贷方，和全年借方正好相抵，
    # 整列加总等于 0（实测 6.1E-09）。排除后才是 BKD 的「本期账面数」口径。
    for r, label, l1 in ((P_TOT0, "销售费用（VC）本期全部费用交易", "销售费用"),
                         (P_TOT0 + 1, "管理费用（VD）本期全部费用交易", "管理费用")):
        ws.cell(r, 4, label)
        # 预审段来自上期 TOD、剩余段来自本期样本池，两个来源都写上
        ws.cell(r, 5, ("%s + %s" % (prior_tod, TOD_POOL)) if use_prior else TOD_POOL)
        f_pre = _pre_formula(
            r, 6, '=SUMIFS(%s,%s,"%s",%s,"<>%s"%s%s)'
                  % (rng_m, rng_e, l1, rng_n, KIND_XFER, cd_bal, cd_pre))
        _fmt(ws.cell(r, 6, f_pre), AMT_FMT)
        _fmt(ws.cell(r, 7, '=SUMIFS(%s,%s,"%s",%s,"<>%s"%s)'
                     % (rng_m, rng_e, l1, rng_n, KIND_XFER, cd_rest)), AMT_FMT)
    # 合计与「TOD 年审」页 1.1、以及 BKD 合计应当完全一致：池子已在源头掐掉上一年度凭证，
    # 这里两边取的都是「本期（含期后）」的口径。有差额先查池子的日期列。
    try:
        from openpyxl.comments import Comment
        ws.cell(15, 8).comment = Comment(
            "合计＝样本池「折合人民币」按日期切两段。\n"
            "若与「TOD 年审」页 1.1 的总体金额不等，先看样本池里有没有资产负债表日所在年度之外的日期。", "PTW")
    except Exception:
        pass

    # ---- 1.2 剔除：8 个桶 + 负值，每行拆两段；I 列模板自带 =G+H ----
    for r, kind in EXCL_ROWS:
        ws.cell(r, 5, ("%s + %s" % (prior_tod, TOD_POOL)) if use_prior else TOD_POOL)
        tag = EXCL_XREF.get(kind, "")
        ws.cell(r, 6, (tag + d8) if (tag and tag != "—" and d8) else tag)
        if kind == "租赁费":
            # V1.2.80：租赁费不再默认排除。本行必须强制 0 ——
            # 预审段是「链接上期 TOD」的，上期按旧口径剔除的租赁费会被原样引进来，
            # 表上就变成"口径里说不含租赁费、数字里还剔着租赁费"，自相矛盾。
            _fmt(ws.cell(r, 7, 0), AMT_FMT)
            _fmt(ws.cell(r, 8, 0), AMT_FMT)
            try:
                from openpyxl.comments import Comment as _Cm
                ws.cell(r, 7).comment = _Cm(
                    "V1.2.80 起租赁费不作默认排除项，本行保留模板行位、金额恒为 0。\n"
                    "原因：短期租赁与低价值租赁直接计入费用，没有对应的使用权资产底稿可承接；"
                    "一刀切排除会让样本池漏掉一整类高金额业务，且无法用其他底稿的测试结果替代。\n"
                    "租赁费现按普通费用进样本池，由金额阈值与抽样决定是否需要执行程序。", "PTW")
            except Exception:
                pass
            continue
        f_pre = _pre_formula(
            r, 7, '=ROUND(SUMIFS(%s,%s,"%s"%s%s),2)'
                  % (rng_m, rng_n, kind, cd_bal, cd_pre))
        _fmt(ws.cell(r, 7, f_pre), AMT_FMT)
        _fmt(ws.cell(r, 8, '=ROUND(SUMIFS(%s,%s,"%s"%s),2)'
                     % (rng_m, rng_n, kind, cd_rest)), AMT_FMT)
    ws.cell(P_EXCL_NEG, 5, ("%s + %s" % (prior_tod, TOD_POOL)) if use_prior else TOD_POOL)
    ws.cell(P_EXCL_NEG, 6, "见表2 负值分析")
    f_pre = _pre_formula(
        P_EXCL_NEG, 7, '=ROUND(SUMIFS(%s,%s,"%s"%s%s),2)'
                       % (rng_m, rng_n, KIND_NEG, cd_bal, cd_pre))
    _fmt(ws.cell(P_EXCL_NEG, 7, f_pre), AMT_FMT)
    _fmt(ws.cell(P_EXCL_NEG, 8, '=ROUND(SUMIFS(%s,%s,"%s"%s),2)'
                 % (rng_m, rng_n, KIND_NEG, cd_rest)), AMT_FMT)
    try:
        from openpyxl.comments import Comment
        _p = sum(float(r.get("amt") or 0) for r in rows
                 if r.get("excl") == KIND_ACCRUAL and float(r.get("amt") or 0) > 0)
        _n = sum(float(r.get("amt") or 0) for r in rows
                 if r.get("excl") == KIND_ACCRUAL and float(r.get("amt") or 0) < 0)
        for _r, _k in EXCL_ROWS:
            if _k == KIND_ACCRUAL:
                ws.cell(_r, 7).comment = Comment(
                    "本行为净额：预提 %s、冲销 %s，一正一负相抵为 0。\n"
                    "预提与冲销的分项金额见「表2 冲销预提费用」，逐笔见「%s」。"
                    % ("{:,.2f}".format(_p), "{:,.2f}".format(_n), TOD_POOL), "PTW")
                break
    except Exception:
        pass

    # ---- 2 关键项：E/F 预审数量金额、G/H 剩余数量金额（I/J 模板自带加总）----
    ws.cell(P_KEY0, 5, _pre_formula(
        P_KEY0, 5, '=COUNTIFS(%s,"是"%s%s)' % (rng_p, cd_bal, cd_pre)))
    _fmt(ws.cell(P_KEY0, 6, _pre_formula(
        P_KEY0, 6, '=SUMIFS(%s,%s,"是"%s%s)' % (rng_m, rng_p, cd_bal, cd_pre))), AMT_FMT)
    # 剩余段永远自己算（本期新增期间，上期 TOD 里没有）
    ws.cell(P_KEY0, 7, '=COUNTIFS(%s,"是"%s)' % (rng_p, cd_rest))
    _fmt(ws.cell(P_KEY0, 8, '=SUMIFS(%s,%s,"是"%s)'
                 % (rng_m, rng_p, cd_rest)), AMT_FMT)
    # 「性质特殊[款项性质]」与「[其他项目组认为存在风险的样本]」默认填 0（V1.2.80）。
    # 这两项是**审计师的判断结果**，工具不替人决定往里放什么，也不从上期套用
    # （上期的特殊款项跟本期未必同类，套过来反而误导）。写 0 表示"本项未识别到"：
    # 比留白清楚 —— 空白读起来像"忘了填"，还会让合计行看着像缺数。
    # 项目组实际识别到时直接改数字，合计（E39/F39 是模板公式）自动跟着变。
    for r in (P_KEY0 + 1, P_KEY0 + 2):
        for col in (5, 6, 7, 8):
            _fmt(ws.cell(r, col, 0), AMT_FMT if col in (6, 8) else "0")

    # ---- 表1（预审阶段）/ 表1.2（年审阶段）----
    keys = [r for r in rows if r.get("todKey") == "是"]
    k1, k2 = [], []
    for r in keys:
        v = r.get("vdate")
        if hasattr(v, "date"):
            v = v.date()
        (k1 if hasattr(v, "year") and v <= d1 else k2).append(r)
    if use_prior:
        # 预审段已经整体引用上期 TOD，明细也以上期那张表为准。
        # 这里若仍按**本期阈值**重挑，会出现「2 段写 7 个、表1 列了 12 行」的自相矛盾。
        k1 = []
    k1.sort(key=lambda r: -abs(float(r.get("amt") or 0)))
    k2.sort(key=lambda r: -abs(float(r.get("amt") or 0)))
    n1 = _write_sample_rows(ws, P_T1_0, k1, P_T1_CAP, P_T1_END)
    prior_sample_rows = 0
    if use_prior and prior_tod:
        # 上期样本**原样搬进来**：本页要能自己看明白预审抽了哪几笔。
        # 只搬了 k1=[] 的空位，行数超模板容量时按需插行，插行数回传给表1.2 定位。
        prior_sample_rows, _ins = _copy_prior_samples(
            wb, ws, prior_tod, P_T1_0, P_T1_CAP, P_T1_END)
        n1 = max(n1, _ins)
    n2 = _write_sample_rows(ws, P_T12_0 + n1, k2, P_T12_CAP, P_T12_END + n1)
    o = n1 + n2

    # ---- 4 分配策略：写在模板 D48:F50 的输入框里，数字取实际条数 ----
    try:
        import datetime as _dt2
        if k2:
            span = ("预审期间（至 %s）%d 个，剩余期间（%s 至 %s）%d 个"
                    % (d1.isoformat(), len(k1),
                       (d1 + _dt2.timedelta(days=1)).isoformat(), bal.isoformat(), len(k2)))
        else:
            span = ("预审期间（至 %s）%d 个（本次未切分剩余期间，全部样本在预审段执行）"
                    % (d1.isoformat(), len(k1)))
        txt = ("本次将 TOD 样本分配到预审和年审两个阶段执行。剔除已在其他底稿执行程序的项目后，"
               "按单笔金额是否达到测试阈值（TE×50%%）确定关键项：%s。"
               "代表性样本按安永抽样工具从扣除剔除项与关键项后的剩余总体中抽取，"
               "抽样结果分别填入表1、表1.2。" % span)
        # 1.2 插了一行股权激励，模板的 48 行（输入框首行）下移到 49
        c = ws.cell(49, 4, txt)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        # D48:F50 是模板的输入框（3 行），默认行高放不下 5 行字，会把后半句切掉。
        # 跟 01.4 表2 的做法一致：只调这三行的高度，不碰其它格式。
        for _r in (48, 49, 50):
            ws.row_dimensions[_r].height = 30
    except Exception:
        pass

    # ---- 表2 负值（行号随插行偏移；模板 D111/D112 的公式要按新行号重写）----
    neg_all = ('SUMIFS(%s,%s,"<0"%s)-SUMIFS(%s,%s,"%s",%s,"<0"%s)'
               % (rng_m, rng_m, cd_bal, rng_m, rng_n, KIND_XFER, rng_m, cd_bal))
    ws.cell(P_NEG_ACCR + o, 4, '=SUMIFS(%s,%s,"%s",%s,"<0"%s)'
            % (rng_m, rng_n, KIND_ACCRUAL, rng_m, cd_bal))
    ws.cell(P_NEG_OTHER + o, 4, "=(%s)-D%d" % (neg_all, P_NEG_ACCR + o))
    ws.cell(P_NEG_SUM + o, 4, "=SUM(D%d:D%d)" % (P_NEG_ACCR + o, P_NEG_OTHER + o))
    # 模板原式是 `=D111-I27`（表2 合计 与 1.2 的「负值」桶比）。但 1.2 的负值桶只装
    # 「说不出性质」的负数，工资/折旧那些负数都在 a 段各桶里，两者天然不等。
    # 表2 要分析的是**全部负数**，所以核对口径改成「合计 − 全部负值」，与年审页一致。
    ws.cell(P_NEG_CHK + o, 4, "=D%d-(%s)" % (P_NEG_SUM + o, neg_all))
    for r in (P_NEG_ACCR + o, P_NEG_OTHER + o, P_NEG_SUM + o, P_NEG_CHK + o):
        _fmt(ws.cell(r, 4), AMT_FMT)

    neg_accr = -sum(float(r.get("amt") or 0) for r in rows
                    if r.get("excl") == KIND_ACCRUAL and float(r.get("amt") or 0) < 0
                    and r.get("inPeriod", True))
    neg_total = -sum(float(r.get("amt") or 0) for r in rows
                     if float(r.get("amt") or 0) < 0 and r.get("excl") != KIND_XFER
                     and r.get("inPeriod", True))
    neg_other = neg_total - neg_accr
    ws.cell(P_NEG_ACCR + o, 5, "是" if neg_accr >= threshold else "否")
    ws.cell(P_NEG_OTHER + o, 5, "是" if neg_other >= threshold else "否")
    # b.1 抽样策略留人工（C 级：抽样结论不代下）
    for r in range(P_NEG_OTHER + 2 + o, P_NEG_HDR + o):
        cell = ws.cell(r, 3)
        if cell.value in (None, ""):
            cell.fill = yellow

    # ---- b.2 负值明细（128-130 三行，超出插行）----
    negs = [r for r in rows
            if r.get("excl") == KIND_NEG and r.get("inPeriod", True)]
    negs.sort(key=lambda r: float(r.get("amt") or 0))
    negs = [r for r in negs if abs(_base_amt(r)) >= threshold]
    base = P_NEG_0 + o
    if len(negs) > P_NEG_CAP:
        _insert_rows(ws, base + P_NEG_CAP, len(negs) - P_NEG_CAP)
    for i, it in enumerate(negs):
        r = base + i
        if i >= P_NEG_CAP:
            _copy_style(ws, base, r, cols=range(3, 14))
        ws.cell(r, 3, it.get("vkey") or it.get("vno") or "")
        vd = it.get("vdate")
        if hasattr(vd, "date"):
            vd = vd.date()
        _fmt(ws.cell(r, 4, vd), "mm-dd-yy")
        ws.cell(r, 5, str(it.get("summary") or "")[:120])
        ws.cell(r, 6, it.get("code") or "")
        ws.cell(r, 7, it.get("name") or "")
        _fmt(ws.cell(r, 8, round(float(it.get("amt") or 0), 2)), AMT_FMT)
        for cidx in (9, 10, 11, 12, 13):
            ws.cell(r, cidx).fill = yellow

    # ---- CHECK：每行合计后面与本期 BKD 未审汇总核对 ----
    chk = {}
    try:
        chk = _fill_check(ws, wb, rows, layout="split")
    except Exception as e:
        chk = {"error": "%s: %s" % (type(e).__name__, str(e)[:120])}

    return {
        "status": "ok", "sheet": PRE_SHEET,
        "preEnd": str(d1), "balanceDate": str(bal),
        "keyPre": len(k1), "keyRest": len(k2),
        "priorSamples": prior_sample_rows,
        "inserted": o, "negRows": len(negs),
        "negAccrual": round(neg_accr, 2), "negOther": round(neg_other, 2),
        "check": chk,
    }



# ---------------------------------------------------------------- 总入口
def build(wb, rows, stats=None, balance_date="", threshold=275000.0,
          tpl_dir=None, tpl_name=None, pre_end=None, prior_tod_bytes=None,
          prior_tod_sheet=None, period_from="1"):
    """复制两张 TOD 页 → 打 TOD 标记 → 填表。

    模板里本来就有两版：`TOD 年审`（全年一张）与 `TOD 预审+剩余期间`（金额拆两段）。
    项目组两种做法都在用，所以两份一起生成，谁用哪份自己删。
    pool 由调用方用同一个 rows 另写一份（两份底稿看的是同一份池子）。

    prior_tod_bytes 给了就把上一期 TOD 整页搬进来，
    「预审+剩余期间」的预审段改引用它 —— 两期重要性水平不同，不能重算。
    prior_tod_sheet 可在上期文件里手动指定用哪张表（页签不规范时用）。
    """
    if not tpl_dir:
        return {"status": "skip", "reason": "没有模板目录"}
    info = mark_rows(rows, threshold, balance_date)

    # 先搬上期 TOD：fill_pre 要用它的 sheet 名与模板类型做链接
    prior_info = {}
    if prior_tod_bytes:
        try:
            prior_info = copy_prior_tod_as_sheet(wb, prior_tod_bytes,
                                                 want_sheet=prior_tod_sheet)
        except Exception as e:
            prior_info = {"error": "%s: %s" % (type(e).__name__, str(e)[:160])}
    # 期间衔接校验：本期预审截止日应与上期资产负债表日相等，不等就明确提示
    if prior_info.get("priorDate") and pre_end not in (None, ""):
        try:
            import datetime as _dt3
            d1, _b = _split_dates(balance_date, pre_end)
            py, pm_, pd_ = (int(x) for x in prior_info["priorDate"][:10].split("-"))
            if d1 and _dt3.date(py, pm_, pd_) != d1:
                prior_info["warning"] = (
                    "本期预审段截至 %s，上期 TOD 覆盖到 %s，两段不衔接；"
                    "上期页已原样搬入，接缝期间的金额需单独说明" % (d1, prior_info["priorDate"]))
        except Exception:
            pass

    ws, err = copy_sheet(tpl_dir, wb, TOD_SHEET)
    if ws is None:
        return {"status": "error", "reason": err, "kindCount": info["kindCount"]}
    wsp, err2 = copy_sheet(tpl_dir, wb, PRE_SHEET)
    # 位置：紧跟在模板自带的「VC&VD.01.2 详细测试 TOD」后面，年审在前、预审在次；
    # 上期 TOD 页跟在预审页后面（它只服务于预审段）
    try:
        want = wb.sheetnames.index("VC&VD.01.2 详细测试 TOD") + 1
    except ValueError:
        want = len(wb.sheetnames)
    _seq = [n for n in (TOD_SHEET, PRE_SHEET) if n in wb.sheetnames]
    if prior_info.get("sheet"):
        _seq.append(prior_info["sheet"])
    for _i, _nm in enumerate(_seq):
        try:
            wb.move_sheet(_nm, offset=(want + _i) - wb.sheetnames.index(_nm))
        except Exception:
            pass
    try:
        finfo = fill(wb, rows, stats=stats, balance_date=balance_date,
                     threshold=threshold, pool_last=len(rows) + 3)
    except Exception as e:
        return {"status": "error", "reason": "%s: %s" % (type(e).__name__, str(e)[:200]),
                "kindCount": info["kindCount"], "keys": info["keys"]}
    # 预审+剩余期间页：复制失败不影响年审页，单独报
    if wsp is None:
        finfo["pre"] = {"status": "error", "reason": err2}
    else:
        try:
            finfo["pre"] = fill_pre(wb, rows, stats=stats, balance_date=balance_date,
                                    threshold=threshold, pool_last=len(rows) + 3,
                                    pre_end=pre_end,
                                    prior_tod=prior_info.get("sheet"),
                                    prior_kind=prior_info.get("kind"),
                                    period_from=period_from)
        except Exception as e:
            finfo["pre"] = {"status": "error",
                            "reason": "%s: %s" % (type(e).__name__, str(e)[:200])}
    finfo.update({"kindCount": info["kindCount"], "keyAmount": info["keyAmount"],
                  "poolIds": info.get("poolIds", 0), "priorTod": prior_info,
                  "periodLabels": _period_labels(period_from, balance_date, pre_end)})
    return finfo
