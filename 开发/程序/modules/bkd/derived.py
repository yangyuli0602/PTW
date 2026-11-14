# -*- coding: utf-8 -*-
"""货币资金派生表：C.01b / C.00b / C.02 / C.02a

设计原则（来自《审计上数Skill》§4.5 与 §7）：

1. **按表号解析、不按序号**。两个版本的 C.01b / C.00b 在 workbook 里位置互换过，
   按 sheet 顺序取必然会取错，而且错得安静。
2. **模板里没有这张表就跳过并记理由**，不新建 sheet。Skill 明令「改非目标 sheet」
   属 D 级禁止，凭空加一张表同样是改结构。
3. **只填能由账面直接派生的**。对账单金额、现金等价物分类、调节项结论、汇率来源
   都是 C 级人工判断，工具只留空标黄 + 写清楚缺什么，不替人下结论。
"""
import os
import re
import copy as _copy

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

try:                                       # 作为包的一部分
    from .core import (T, r2, KZ_DATA, cur_code, _locate_bkd_cols, _locate_restricted_col,
                       _locate_fx_cols, _find_sum_row, _norm_subject_l1,
                       locate_cols, C01B_COLS, C01B_REQUIRED, C02A_COLS, C02A_REQUIRED,
                       ACCT_FMT)
except ImportError:                        # 脚本把 modules/bkd 直接加进 sys.path 时
    from core import (T, r2, KZ_DATA, cur_code, _locate_bkd_cols, _locate_restricted_col,
                      _locate_fx_cols, _find_sum_row, _norm_subject_l1,
                      locate_cols, C01B_COLS, C01B_REQUIRED, C02A_COLS, C02A_REQUIRED,
                      ACCT_FMT)

#: 金额格式跟着 core 走（会计专用、无货币符号、零显示 -），派生表不另立一套
MONEY_FMT = ACCT_FMT


try:                                       # 作为包的一部分
    from .tasks import add_task as _add_task
except ImportError:                        # 脚本把 modules/bkd 直接加进 sys.path 时
    from tasks import add_task as _add_task

GAP_FILL = PatternFill("solid", fgColor="FFFF00")
_THIN = Side(style="thin", color="808080")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


# ---------------------------------------------------------------- 工具
def _sheet_head(name):
    """取表号段：'C.02 Bank reconciliations' -> 'C.02'，'C.02a 外币测算' -> 'C.02A'。"""
    m = re.match(r"^([A-Za-z0-9.&]+)", re.sub(r"[\s\u3000]+", " ", str(name or "").strip()))
    return (m.group(1) if m else "").upper()


def find_sheet(wb, code, keywords=()):
    """按表号精确匹配，再退回关键词模糊匹配。

    必须精确比表号段：直接用 startswith("C.02") 会把 C.02a 也吃掉，
    于是外币测算表被当成余额调节表填，两处数据一起错。
    """
    want = code.upper()
    for n in wb.sheetnames:
        if _sheet_head(n) == want:
            return n
    for n in wb.sheetnames:
        flat = re.sub(r"[\s\u3000]", "", str(n)).upper()
        for kw in keywords:
            if re.sub(r"[\s\u3000]", "", kw).upper() in flat:
                return n
    return None


def _find_header(ws, must=(), limit=40):
    """找同时含 must 里所有关键词的那一行（表头行）。返回行号或 None。"""
    for r in range(1, min(ws.max_row, limit) + 1):
        txt = "".join(T(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 30) + 1))
        if all(k in txt for k in must):
            return r
    return None


def _find_label_row(ws, label, col=2, start=1):
    """在指定列里找文字等于/包含 label 的行（用于「合计」「Rx」这类锚点）。"""
    for r in range(start, min(ws.max_row, start + 400) + 1):
        if label in T(ws.cell(r, col).value).replace(" ", ""):
            return r
    return None


#: 本位币与外币的各种写法视为同一币种。BKD 里常见写「人民币 / 美元」，模板里写
#: 「CNY / USD」；不归一会导致「本期明明有美元账户，却判断成模板行没被使用」，
#: 于是美元汇率不标黄，C.02a 静默算出一个差额为 0 的假象。
_CUR_ALIAS = {
    # 本位币
    "人民币": "CNY", "RMB": "CNY", "CNY": "CNY", "本币": "CNY", "本位币": "CNY",
    # 主要外币（中英文写法都收，账套里两种都有）
    "美元": "USD", "美金": "USD", "USD": "USD",
    "港币": "HKD", "港元": "HKD", "香港元": "HKD", "HKD": "HKD",
    "欧元": "EUR", "EUR": "EUR",
    "日元": "JPY", "日圆": "JPY", "日本日圆": "JPY", "JPY": "JPY",
    "英镑": "GBP", "GBP": "GBP",
    "新加坡元": "SGD", "新加坡币": "SGD", "SGD": "SGD",
    "新台币": "TWD", "TWD": "TWD",
    "澳币": "AUD", "澳大利亚元": "AUD", "AUD": "AUD",
    "墨西哥币": "MXN", "墨西哥比索": "MXN", "MXN": "MXN",
    "巴西雷亚尔": "BRL", "BRL": "BRL",
    "挪威克朗": "NOK", "NOK": "NOK",
    "韩元": "KRW", "KRW": "KRW",
    "泰铢": "THB", "THB": "THB",
    "加元": "CAD", "加拿大元": "CAD", "CAD": "CAD",
    "瑞士法郎": "CHF", "CHF": "CHF",
    "印度卢比": "INR", "INR": "INR",
}


def _cur_norm(v):
    t = T(v).strip().upper()
    return _CUR_ALIAS.get(t, t)


def _put(ws, r, c, v, fmt=None, fill=None, bold=False, right=False):
    cell = ws.cell(r, c, v)
    cell.font = Font(name="Arial", size=10, bold=bold)
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    cell.alignment = Alignment(
        horizontal="right" if (right or fmt == MONEY_FMT) else "left",
        vertical="center")
    if fill is not None:
        cell.fill = fill
    return cell


def _bkd_span(ws_bkd):
    """返回 (列字典, 数据首行, 数据末行, 合计行)。派生表要按实际区间取值，不用整列。"""
    col = _locate_bkd_cols(ws_bkd)
    if not col:
        return None, 0, 0, None
    dstart = col["_header_row"] + 1
    srow = _find_sum_row(ws_bkd, dstart, col["localEnd"])
    dend = (srow - 1) if srow and srow - 1 >= dstart else min(ws_bkd.max_row, dstart + 500)
    if srow and len(T(ws_bkd.cell(srow - 1, col["company"]).value)) > 12:
        dend = max(dstart, srow - 2)
    return col, dstart, dend, srow


def _copy_row_style(ws, src, dst, max_col):
    """把 src 行的字体/边框/填充/数字格式/行高复制到 dst 行。

    openpyxl 的 insert_rows 只在 XML 层面插空行，**不复制样式**，
    新行会变成没边框的白板，打印出来一眼就看得出是程序插的。
    """
    from copy import copy
    for c in range(1, max_col + 1):
        s, d = ws.cell(src, c), ws.cell(dst, c)
        if s.has_style:
            d.font = copy(s.font)
            d.border = copy(s.border)
            d.fill = copy(s.fill)
            d.alignment = copy(s.alignment)
            d.number_format = s.number_format
    try:
        h = ws.row_dimensions[src].height
        if h:
            ws.row_dimensions[dst].height = h
    except Exception:
        pass


def _ensure_rows(ws, hdr, total_row, need, max_col):
    """把数据区从 (hdr, total_row) 之间扩到 need 行，返回新的合计行号。

    为什么必须扩：V7 模板 C.01b 只预置 4 行、C.00b 3 行、C.02a 4 个币种，
    而真实客户一家就有 53 个其他货币资金账户、5 种以上币种。不扩行的话
    生成出来的底稿**少一半账户而且不报错**——这是最危险的一种"成功"。

    插行时必须手动搬合并单元格：openpyxl 不会动它们，不搬就会把下面的
    说明框（B19:F22 之类）压到数据区上。
    """
    preset = max(0, (total_row or hdr + 1) - hdr - 1)
    if need <= preset or not total_row:
        return total_row
    n = need - preset
    shifted = []
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row >= total_row:
            shifted.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
            ws.unmerge_cells(str(mc))
    ws.insert_rows(total_row, n)
    for (a, b, c, d) in shifted:
        ws.merge_cells(start_row=a + n, start_column=c, end_row=b + n, end_column=d)
    for r in range(total_row, total_row + n):
        _copy_row_style(ws, total_row - 1, r, max_col)
    return total_row + n


def _paint_like_first_row(ws, first, last, c1, c2):
    """用 first 行的格式刷 first..last：字体/边框/对齐/数字格式/行高，不含值和底色。

    为什么以第一行为准：表头下面那一行是模板给的"标准数据行"，客户在模板里调过的
    字体、边框、金额格式都落在这行上，自己另造一套样式会和模板打架。

    **底色不复制**：空白提示交给条件格式（填上值自动消失）。复制底色会把第一行的
    黄刷到每一行，填了值也不消失——那正是之前 BKD 看着假的原因。
    """
    from copy import copy
    h = ws.row_dimensions[first].height
    for r in range(first, last + 1):
        for c in range(c1, c2 + 1):
            try:
                src, dst = ws.cell(first, c), ws.cell(r, c)
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.alignment = copy(src.alignment)
                dst.number_format = src.number_format
                dst.fill = PatternFill(fill_type=None)
            except Exception:
                continue
        if h:
            ws.row_dimensions[r].height = h


def _note_boxes(ws, after_row):
    """数据区下方第一个合并区（Notes 说明框）的左上格，给条件格式用。"""
    for mc in sorted(ws.merged_cells.ranges, key=lambda x: x.min_row):
        if mc.min_row > after_row:
            return [(mc.min_row, mc.min_col)]
    return []


def _restore_blank_cf(ws, data_box, extra_boxes=()):
    """把模板自带的「空白高亮」条件格式重铺到新的数据范围。

    模板把范围写死了（C.01b 是 B10:L12 + B19，BKD 是 B10:K15 + B22）。
    插行插列之后它既盖不到新行、也盖不到新列，空白格就全变白了——
    而这些表的"哪里要补"全靠这层浅黄提示，白掉等于把提示弄丢了。

    data_box=(r1, r2, c1, c2)；extra_boxes=[(r, c)] 是 Notes 框之类的单格。
    颜色/图案一律 deepcopy 模板原规则，不自己造。
    """
    from openpyxl.formatting.formatting import ConditionalFormattingList
    src = None
    for cf in ws.conditional_formatting:
        for r in cf.rules:
            if r.dxf is not None and r.dxf.fill is not None:
                src = r
                break
        if src is not None:
            break
    if src is None:
        return 0

    r1, r2, c1, c2 = data_box
    ws.conditional_formatting = ConditionalFormattingList()
    rule = _copy.deepcopy(src)
    rule.formula = ["LEN(TRIM(%s%d))=0" % (get_column_letter(c1), r1)]
    ws.conditional_formatting.add(
        "%s%d:%s%d" % (get_column_letter(c1), r1, get_column_letter(c2), r2), rule)
    for (rr, cc) in extra_boxes:
        sub = _copy.deepcopy(src)
        sub.formula = ["LEN(TRIM(%s%d))=0" % (get_column_letter(cc), rr)]
        ws.conditional_formatting.add("%s%d" % (get_column_letter(cc), rr), sub)
    return 1 + len(extra_boxes)


def _find_total_row(ws, hdr, label_col=7, start=None):
    """找合计行：优先「合计」文字，退回到带 SUM 公式的第一行。

    名字不能叫 _find_sum_row——那个是从 core 导入的（按金额列找 SUM 行），
    覆盖掉会让 _bkd_span 拿着一份不存在的签名去调用。
    """
    s = start or (hdr + 1)
    r = _find_label_row(ws, "合计", label_col, s)
    if r:
        return r
    for rr in range(s, min(ws.max_row, s + 400) + 1):
        for c in range(1, min(ws.max_column, 15) + 1):
            v = ws.cell(rr, c).value
            if isinstance(v, str) and v.upper().startswith("=SUM("):
                return rr
    return None


# ---------------------------------------------------------------- C.01b
def fill_c01b(wb, res, meta=None):
    """C.01b 特殊账户程序：取 BKD 里「其他货币资金、有余额、且不执行函证」的账户。

    口径（2026-11-08 用户定）：C.01b 是「未函证账户的替代程序」，
    所以放「是否函证 = 否」的户；函证为「是」的走函证，不在这里重复做一遍。
    函证列**空着的也放**（还没判断完也得有替代程序兜着），是否函证到底怎么定
    由最右侧新增的「是否函证」列 XLOOKUP 回 BKD 实时显示——在 BKD 改了判断，
    这里立刻跟着变，一眼能看出哪几户还没定。
    对账单金额（I 列）本次拿不到 —— 那是客户资料，不是账面能派生的，
    所以留空标黄并进 gaps，不填 0（填 0 会让「差异 = 账面」看起来像真差异）。
    """
    name = find_sheet(wb, "C.01b", ["特殊账户"])
    if not name:
        return {"sheet": None, "status": "skip", "reason": "模板里没有 C.01b 表"}
    ws = wb[name]
    hdr = _find_header(ws, ["科目名称", "账号"])
    if not hdr:
        return {"sheet": name, "status": "skip",
                "reason": "C.01b 表头识别失败（找不到「科目名称 + 账号」那一行），未写入"}
    # 列按表头文字定位。模板多插一列时不再把公司名写进科目列，
    # 缺必需列就整表跳过，让人看得见，不写到隔壁列。
    cmap, cmiss = locate_cols(ws, hdr, C01B_COLS)
    bad = [k for k in C01B_REQUIRED if k not in cmap]
    if bad:
        return {"sheet": name, "status": "skip",
                "reason": "C.01b 列识别失败，未写入：缺 %s（表头第 %d 行）"
                          % ("、".join(bad), hdr)}

    def _col(key):
        return cmap.get(key)

    def put(r, key, v, fmt=None):
        c = _col(key)
        if c:
            _put(ws, r, c, v, fmt)

    rows = [r for r in (res.get("rows") or [])
            if _norm_subject_l1(r.get("subject")) == "其他货币资金"
            and abs(r.get("local") or 0) > 0.004
            and T(r.get("confirm")) in ("否", "")]

    total_row = _find_total_row(ws, hdr) or (hdr + 1 + max(1, len(rows)))
    total_row = _ensure_rows(ws, hdr, total_row, len(rows), 15)
    MONEY = MONEY_FMT
    for i, row in enumerate(rows):
        r = hdr + 1 + i
        put(r, "company", T(row.get("company")))
        put(r, "subject", _norm_subject_l1(row.get("subject")) or T(row.get("subject")))
        put(r, "bank", T(row.get("bankName") or row.get("bank")))
        put(r, "acct", T(row.get("acctNo")))
        put(r, "currency", T(row.get("currency")) or "CNY")
        put(r, "orig", r2(row.get("orig")), MONEY)
        put(r, "local", r2(row.get("rmb") if row.get("rmb") is not None else row.get("local")), MONEY)
        # 对账单金额：拿不到 → 留空。不涂色，交给模板的「空白高亮」
        # 条件格式（填上客户给的数就自动消失，硬涂的黄不会消失）。
        put(r, "statement", None, MONEY)
        # 差异：模板原本是 =I-G，I 空时差异 = -G，是假差异。改成 I 空则空。
        c_diff, c_st, c_orig = _col("diff"), _col("statement"), _col("orig")
        if c_diff and c_st and c_orig:
            ws.cell(r, c_diff).value = '=IF(%s%d="","",%s%d-%s%d)' % (
                get_column_letter(c_st), r, get_column_letter(c_st), r,
                get_column_letter(c_orig), r)
            ws.cell(r, c_diff).number_format = MONEY
        put(r, "nature", T(row.get("nature")))
        # 本期末实际受限金额：直接取 BKD 受限列（一笔只数一次）
        put(r, "restricted",
            r2(row.get("restricted")) if row.get("restricted") is not None else None, MONEY)
        # 右侧说明列跟着补边框，否则插出来的行是断的
        c_last = max(cmap.values())
        for c in range(cmap.get("reason", c_last) + 1, c_last + 1):
            _put(ws, r, c, None)

    # ---- 最右侧新增「是否函证」列：XLOOKUP 回 C.00 BKD，BKD 改了判断这里立刻跟着变 ----
    # 为什么用公式而不是抄值：这一列就是给复核人"一眼看出哪几户还没定"用的，
    # 抄成常量的话，BKD 里把「否」改成「是」之后这里还停在旧值，比不放更误导。
    c_flag = None
    bkd_name = next((n for n in wb.sheetnames
                     if "BKD" in n.upper() and _sheet_head(n) == "C.00"), None)
    bcol, bstart, bend, _bsrow = (None, 0, 0, None)
    if bkd_name:
        bcol, bstart, bend, _bsrow = _bkd_span(wb[bkd_name])
    c_last = max(cmap.values())
    # 匹配键：优先用「银行/存款机构名称」（= 唯一主键，BKD 里每行都有）；
    # 不能用账号——V1.2.61 起 BKD 账号列只对上年匹配上的户填值，本年新增户是空的，
    # 拿账号当键会整片取不到。
    c_key = cmap.get("bank") if bcol and bcol.get("bankName") else cmap.get("acct")
    b_key = bcol.get("bankName") or bcol.get("acctNo") if bcol else None
    if bcol and bcol.get("confirm") and bend >= bstart and c_key and b_key:
        c_flag = c_last + 1
        # 表头沿用右侧邻居的样式，保持"整块表格"的观感
        hcell = ws.cell(hdr, c_flag)
        try:
            hcell._style = _copy.copy(ws.cell(hdr, c_last)._style)
        except Exception:
            pass
        hcell.value = "是否函证"
        hcell.font = Font(name="Arial", size=10, bold=True)
        hcell.alignment = Alignment(horizontal="center", vertical="center")
        hcell.border = BORDER
        b_acct = get_column_letter(b_key)
        b_flag = get_column_letter(bcol["confirm"])
        c_acct = c_key
        for i in range(len(rows)):
            r = hdr + 1 + i
            cell = ws.cell(r, c_flag)
            cell.value = (
                '=IF(%s%d="","",XLOOKUP(%s%d,\'%s\'!$%s$%d:$%s$%d,'
                '\'%s\'!$%s$%d:$%s$%d,"")&"")'
                % (get_column_letter(c_acct), r,
                   get_column_letter(c_acct), r,
                   bkd_name, b_acct, bstart, b_acct, bend,
                   bkd_name, b_flag, bstart, b_flag, bend))
            cell.font = Font(name="Arial", size=10)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
        try:
            ws.column_dimensions[get_column_letter(c_flag)].width = 11
        except Exception:
            pass

    # 模板预置了 3 行样例、实际账户可能更少：多出来的行会留着模板的 `=I-G`
    # （I 列为空时算出 0.00，看着像"差异为零"的假结论）。清掉未使用的行，只留格式。
    for r in range(hdr + 1 + len(rows), (total_row or hdr + 1)):
        for c in cmap.values():
            ws.cell(r, c).value = None

    if total_row:
        last = hdr + max(1, len(rows))
        for key in ("orig", "local", "restricted"):
            c = _col(key)
            if not c:
                continue
            letter = get_column_letter(c)
            ws.cell(total_row, c).value = "=SUM(%s%d:%s%d)" % (letter, hdr + 1, letter, last)
            ws.cell(total_row, c).number_format = MONEY
    # ---- 用第一行数据行的格式刷全表 ----
    # 插出来的行没有行高、账号列是 General（会同模板第一行的文本格式打架），
    # 逐行看是"有的行是模板、有的行是程序插的"。统一刷成第一行的样子。
    # 数据行是 hdr+1 .. hdr+len(rows)（写入循环里 r = hdr+1+i，i 到 len-1）
    last_data = hdr + max(1, len(rows))
    c_first, c_last = min(cmap.values()), max(cmap.values())
    _paint_like_first_row(ws, hdr + 1, last_data, c_first, c_last)

    # ---- 空白高亮铺满数据区 ----
    # 模板的 CF 只盖 B10:L12 三行，插到 26 行之后后面全是白的；
    # Notes 框那条 B19 也会因为插行错位到别处。两条一起重铺。
    c_hl_end = cmap.get("restricted") or c_last
    res["c01bBlankCf"] = _restore_blank_cf(
        ws, (hdr + 1, last_data, c_first, c_hl_end), _note_boxes(ws, total_row or last_data))

    gaps = res.setdefault("gaps", [])
    if rows:
        _add_task(res, id="c01b-statement", type="document",
                  title="C.01b 需向客户取得对账单/资产证明书（%d 户）" % len(rows),
                  target="%s（I 列 对账单/资产证明书金额）" % name,
                  saveTo="", url="",
                  prompt=("逐户向客户索取银行对账单或资产证明书，金额填入 %s 的 I 列；"
                          "J 列差异自动计算，差异不为 0 的要解释。" % name),
                  blocking=True)
    c_stmt = cmap.get("statement") or c_first
    for i in range(len(rows)):
        r = hdr + 1 + i
        gaps.append({"sheet": name, "row": r, "col": c_stmt, "field": "对账单/资产证明书金额",
                     "company": T(rows[i].get("company")),
                     "subject": "其他货币资金", "acct": T(rows[i].get("acctNo")),
                     "reason": "无数据源：需向客户取得对账单或资产证明书"})
    out = {"sheet": name, "status": "ok", "rows": len(rows),
           "total": r2(sum((w.get("rmb") if w.get("rmb") is not None else w.get("local") or 0)
                           for w in rows)),
           "gaps": len(rows)}
    if c_flag:
        out["confirmCol"] = get_column_letter(c_flag) + " 是否函证（XLOOKUP → C.00 BKD）"
    if not rows:
        out["reason"] = ("本期没有「不执行函证（含未判断）且余额非零」的其他货币资金账户，"
                         "C.01b 无需填明细")
    return out


# ---------------------------------------------------------------- C.00b
def fill_c00b(wb, res, meta=None):
    """C.00b Other classification：表1 填受限货币资金明细。

    表1 以下的两块（现金等价物识别总体、现金等价物明细）需要判断「哪些是现金等价物」，
    属 C 级判断，工具不填，只在返回里说明。
    """
    name = find_sheet(wb, "C.00b", ["Otherclassification", "其他分类"])
    if not name:
        return {"sheet": None, "status": "skip", "reason": "模板里没有 C.00b 表"}
    ws = wb[name]
    hdr = _find_header(ws, ["科目名称", "账号"])
    if not hdr:
        return {"sheet": name, "status": "skip",
                "reason": "C.00b 表头识别失败（找不到「科目名称 + 账号」那一行），未写入"}
    rows = [r for r in (res.get("rows") or []) if abs(r.get("restricted") or 0) > 0.004]
    total_row = _find_total_row(ws, hdr) or (hdr + 1 + max(1, len(rows)))
    total_row = _ensure_rows(ws, hdr, total_row, len(rows), 15)
    for i, row in enumerate(rows):
        r = hdr + 1 + i
        _put(ws, r, 2, T(row.get("company")))
        _put(ws, r, 3, _norm_subject_l1(row.get("subject")) or T(row.get("subject")))
        _put(ws, r, 4, T(row.get("bankName") or row.get("bank")))
        _put(ws, r, 5, T(row.get("acctNo")))
        _put(ws, r, 6, T(row.get("currency")) or "CNY")
        _put(ws, r, 7, r2(row.get("orig")), MONEY_FMT)
        _put(ws, r, 8, r2(row.get("rmb") if row.get("rmb") is not None else row.get("local")), MONEY_FMT)
        _put(ws, r, 9, T(row.get("nature")))
        _put(ws, r, 10, r2(row.get("restricted")), MONEY_FMT)
        for c in (11, 12, 13, 14, 15):
            _put(ws, r, c, None)
    if total_row:
        last = hdr + max(1, len(rows))
        for c, letter in ((8, "H"), (10, "J")):
            ws.cell(total_row, c).value = "=SUM(%s%d:%s%d)" % (letter, hdr + 1, letter, last)
            ws.cell(total_row, c).number_format = MONEY_FMT
    _add_task(res, id="c00b-equiv", type="judgement",
              title="C.00b 需判断哪些账户属于现金等价物",
              target="C.00b Other classification",
              saveTo="", url="",
              prompt=("按期限与流动性判断现金等价物范围（一般为 3 个月内到期、"
                      "易于转换为已知金额现金、价值变动风险很小）；工具不代填。"),
              blocking=False)
    if not rows:
        _add_task(res, id="c00b-restricted", type="input",
                  title="BKD 受限资金列全空，C.00b 无明细",
                  target="C.00 BKD 受限资金列",
                  saveTo="", url="",
                  prompt="在 C.00 BKD 的「受限资金」列逐户填写本期末受限金额后重新生成。",
                  blocking=False)
    out = {"sheet": name, "status": "ok", "rows": len(rows),
           "total": r2(sum(r.get("restricted") or 0 for r in rows))}
    if not rows:
        out["status"] = "skip"
        out["reason"] = ("BKD 受限资金列全为空，表1 无明细可填。"
                         "受限金额需人工在 C.00 BKD 的「受限资金」列填写后重跑")
    out["pending"] = ("表1 的「现金 / 现金等价物」分类与表2 现金等价物明细属人工判断，"
                      "工具未填")
    return out


# ---------------------------------------------------------------- C.02
def fill_c02(wb, res, meta=None):
    """C.02 银行余额调节：只出待做清单，不落笔。

    「本期是否存在调节项」必须有客户的对账单/余额调节表才能判断。Skill 把
    「自动下无调节项结论」列为 D 级禁止 —— 工具在这里落一个「无调节项」，
    等于凭空出具了一个审计结论。所以只统计有多少账户需要做调节，写进返回。
    """
    name = find_sheet(wb, "C.02", ["Bankreconciliation", "余额调节"])
    if not name:
        return {"sheet": None, "status": "skip", "reason": "模板里没有 C.02 表"}
    cand = [r for r in (res.get("rows") or [])
            if _norm_subject_l1(r.get("subject")) in ("银行存款", "其他货币资金")
            and abs(r.get("local") or 0) > 0.004]
    _add_task(res, id="c02-recon", type="judgement",
              title="C.02 银行余额调节：%d 个账户待判断是否存在调节项" % len(cand),
              target="C.02 Bank reconciliations",
              saveTo="", url="",
              prompt=("按对账单/余额调节表逐户判断是否存在未达账项；工具不代下"
                      "「无调节项」结论，需人工判断并在表内填列。"),
              blocking=False)
    return {"sheet": name, "status": "manual",
            "accounts": len(cand),
            "accountList": ["%s/%s/%s" % (T(x.get("company")),
                                          T(x.get("bankName") or x.get("bank")),
                                          T(x.get("acctNo"))) for x in cand[:50]],
            "reason": ("是否存在调节项需对账单/余额调节表，属 C 级判断："
                       "工具未写结论、未填对账单金额，只列候选账户")}


# ---------------------------------------------------------------- C.02a
def fill_c02a(wb, res, meta=None):
    """C.02a 外币测算。

    三条硬规则（都是明确要求过的）：

      1. **只列本期真实出现的币种**。模板预置 USD/CNY/JPY 三行，本期没有日元就
         一行都不留——空行会被当成"漏填"，还会把合计行往下推。
      2. **币种写 ISO 代码，与 BKD 的「币种」列完全一致**。BKD 写「美元」、这里写
         「USD」的话，SUMIF 一个字都对不上：原币 / 账面本位币 / 差异三列静默全 0，
         看着像"本期没有汇兑差异"，实际是根本没取到数。
      3. **表内不画格线、只留外框**，每行格式与第一行数据行完全一致。
    """
    name = find_sheet(wb, "C.02a", ["外币测算", "Foreignexchange"])
    if not name:
        return {"sheet": None, "status": "skip", "reason": "模板里没有 C.02a 表"}
    ws = wb[name]
    bkd = wb["C.00 BKD"]
    col, dstart, dend, _srow = _bkd_span(bkd)
    if not col:
        return {"sheet": name, "status": "skip", "reason": "C.00 BKD 表头识别失败"}

    # 汇率取不到不能拖垮整张表：以前 `from core import fxrate` 写在 try 外面，
    # 分发包少两个 OpenSSL DLL 时这行直接 ImportError，C.02a 整张表没生成
    # （2026-11-07 客户实测）。现在 import + 抓取一起兜，表照出、汇率留空标黄。
    FX, fx, fx_err = None, {}, ""
    try:
        from core import fxrate as FX
        fx = FX.fetch((meta or {}).get("balance_date") or "",
                      url=(meta or {}).get("fx_url") or "",
                      cache_dir=KZ_DATA)
    except Exception as e:
        fx_err = "%s: %s" % (type(e).__name__, str(e)[:160])
        fx = {}

    c_cur = get_column_letter(col.get("currency") or 6)
    c_orig = get_column_letter(col.get("origEnd") or col["localEnd"])
    # F 列（账面本位币金额）取 BKD 的「期末账面金额（人民币）」，**不是「本币」**。
    # 本币是各主体账套自己的本位币：香港/新加坡主体记的是美元，拿它去比
    # 「原币 × 人民币中间价」等于拿美元比人民币，会差出上千万的假差异。
    # 人民币列 = 本币 × 主体折算率，和 E 列（折合本位币）同为人民币口径才可比。
    # 前提是 BKD 折算率按主体填对（美元本位币主体要填美元汇率，不能留 1）。
    _c_rate, c_rmb_fx = _locate_fx_cols(bkd, col["_header_row"])
    if c_rmb_fx:
        c_money = get_column_letter(c_rmb_fx)
    else:
        c_money = get_column_letter(col["localEnd"])
    bkd_cur = "'C.00 BKD'!$%s$%d:$%s$%d" % (c_cur, dstart, c_cur, dend)
    bkd_orig = "'C.00 BKD'!$%s$%d:$%s$%d" % (c_orig, dstart, c_orig, dend)
    bkd_money = "'C.00 BKD'!$%s$%d:$%s$%d" % (c_money, dstart, c_money, dend)

    # 1) 本期有金额的币种（人民币始终保留作本位币基准）
    #    只按"账上有这个币种"判断不够：账上挂着 6 个零余额的外币账户，
    #    表里就会多出 6 行 0.00，看着像漏填。外币测算的目的是折算与差异，
    #    零余额没有测算对象，一行都不出——BKD 里仍然保留这些账户。
    seq = []
    for r in (res.get("rows") or []):
        code = cur_code(r.get("currency"))
        if not code or code in seq:
            continue
        amt = abs(r.get("orig") or 0) + abs(r.get("local") or 0)
        if amt > 0.004 or code == "CNY":
            seq.append(code)
    if "CNY" in seq:
        seq.remove("CNY")
        seq.insert(0, "CNY")
    if not seq:
        seq = ["CNY"]

    hdr = _find_header(ws, ["原币", "汇率"]) or 8
    # 列按表头文字定位（模板列序变了也不会写错列）
    cmap, cmiss = locate_cols(ws, hdr, C02A_COLS)
    bad = [k for k in C02A_REQUIRED if k not in cmap]
    if bad:
        return {"sheet": name, "status": "skip",
                "reason": "C.02a 列识别失败，未写入：缺 %s（表头第 %d 行）"
                          % ("、".join(bad), hdr)}
    c_cur_a, c_orig_a = cmap["currency"], cmap["orig"]
    c_rate_a, c_conv_a = cmap["rate"], cmap["conv"]
    c_book_a, c_diff_a = cmap["book"], cmap["diff"]
    L_cur_a, L_orig_a = get_column_letter(c_cur_a), get_column_letter(c_orig_a)
    L_conv_a, L_book_a = get_column_letter(c_conv_a), get_column_letter(c_book_a)
    L_rate_a, L_diff_a = get_column_letter(c_rate_a), get_column_letter(c_diff_a)

    total = _find_total_row(ws, hdr, 6) or (hdr + 4)
    # 2) 行数对齐实际币种数：模板 4 行，实际可能 11 种，也可能只有 1 种
    total = _fit_rows(ws, hdr, total, len(seq), 9)

    # 表头口径改名：「账面本位币金额」→「账面人民币金额」。
    # 这张表 F 列取的是 BKD 的折算人民币列，本来就是人民币口径；
    # 叫「本位币」会让人以为是各主体自己的本位币（香港主体记的是美元），理解反了会查错方向。
    _book_hdr = ws.cell(hdr, c_book_a)
    if T(_book_hdr.value) and "人民币" not in T(_book_hdr.value):
        _book_hdr.value = "账面人民币金额"

    got, need_rate = [], []
    for i, code in enumerate(seq):
        r = hdr + 1 + i
        _put(ws, r, c_cur_a, code)
        ws.cell(r, c_orig_a).value = "=SUMIF(%s,$%s%d,%s)" % (bkd_cur, L_cur_a, r, bkd_orig)
        # 汇率没填时折合/差异必须留空，不能算：汇率空 → 折合=0 → 差异=0-账面 = 负的账面金额，
        # 那是纯算法产物，会被当成"巨额汇兑差异"报出去（C.01b 的 =I-G 踩过同一个坑）。
        ws.cell(r, c_conv_a).value = '=IF({rt}{r}="","",{ro}{r}*{rt}{r})'.format(
            r=r, rt=L_rate_a, ro=L_orig_a)
        ws.cell(r, c_book_a).value = "=SUMIF(%s,$%s%d,%s)" % (bkd_cur, L_cur_a, r, bkd_money)
        ws.cell(r, c_diff_a).value = '=IF({rt}{r}="","",{rc}{r}-{rb}{r})'.format(
            r=r, rt=L_rate_a, rc=L_conv_a, rb=L_book_a)
        if code == "CNY":
            # 人民币折算率恒为 1：这不是判断，是本位币的定义
            _put(ws, r, c_rate_a, 1, "0.0000")
            continue
        rt = FX.rate_for(code, fx) if FX else None
        if rt:
            _put(ws, r, c_rate_a, rt, "0.0000")
            got.append("%s=%.4f" % (code, rt))
        else:
            _put(ws, r, c_rate_a, None, "0.0000", GAP_FILL)
            ws.cell(r, c_rate_a).comment = Comment(
                "未取到该币种的人民币中间价（网络/公告页/日期原因）。\n"
                "请手工填入资产负债表日汇率，并在 Notes 写明来源。", "PTW")
            need_rate.append(code)

    # 3) 合计行：**标签放最前面**（模板把「合计」写在「账面本位币金额」那一列上，
    #    把该列的金额挤掉了）。原币 / 折合 / 账面人民币 / 差异四列都要有合计。
    ws.cell(total, c_cur_a).value = "合计"
    for _c in (c_orig_a, c_conv_a, c_book_a):
        _L = get_column_letter(_c)
        ws.cell(total, _c).value = "=SUM(%s%d:%s%d)" % (_L, hdr + 1, _L, total - 1)
    ws.cell(total, c_diff_a).value = "=SUM(%s%d:%s%d)" % (L_diff_a, hdr + 1, L_diff_a, total - 1)

    # 3.1) Check 行：原币、人民币各对一次 BKD 合计（都应为 0）
    #   原币合计只作完整性控制——这一列混着 USD/HKD/EUR，相加本身没有金额意义；
    #   人民币合计才是真正的金额勾稽（BKD 折算人民币合计）。
    chk = total + 1
    ws.cell(chk, c_cur_a).value = "Check with BKD"
    if T(ws.cell(chk, c_diff_a).value).strip().upper() == "RX":
        ws.cell(chk, c_diff_a).value = None          # 模板占位标记，别留在差异列里
    _sr = _find_sum_row(bkd, dstart, col["localEnd"])
    if _sr:
        _L_orig_bkd = get_column_letter(col.get("origEnd") or col["localEnd"])
        c1_chk = ws.cell(chk, c_orig_a,
                         "=%s%d-'C.00 BKD'!%s%d" % (L_orig_a, total, _L_orig_bkd, _sr))
        c1_chk.comment = Comment("原币合计 − BKD 原币合计 = 0（完整性控制，不构成金额结论）", "PTW")
        c2_chk = ws.cell(chk, c_book_a,
                         "=%s%d-'C.00 BKD'!%s%d" % (L_book_a, total, c_money, _sr))
        c2_chk.comment = Comment("人民币合计 − BKD 折算人民币合计 = 0（金额勾稽）", "PTW")

    # 4) 只留外框 + 每行格式统一（合计行与 Check 行一起刷，底框画在 Check 行上）
    _normalize_c02a_table(ws, hdr, chk, min(cmap.values()), max(cmap.values()))

    # 5) Notes 写进下面那个大合并框（标签行只留 "Notes"）
    if fx.get("rates"):
        res["c02aNotesEndRow"] = _write_fx_note(ws, fx, got)
        # 截图三步，顺序固定：
        #   ① data/fx/pbc_YYYYMMDD.png 已就位 → 直接用（人工/AI 补的图，不覆盖）
        #   ② 按**资产负债表日**去人行栏目页定位该日公告页（会往前翻页，只认当天）
        #   ③ 用本机 Edge/Chrome 无头截该页
        # 不再看汇率来源：交易中心只决定"汇率从哪儿取"，证据图永远是公告原文那一页。
        # 用户口径（2026-11-11）：2025-09-30 就查 2025-09-30，绝不借用相邻交易日。
        _bd = (meta or {}).get("balance_date") or ""
        sp = FX.shot_path(_bd, KZ_DATA) if FX else ""
        _why = ""
        _ann = {}
        _auto = (os.environ.get("KZ_FX_AUTOSHOT") or "1").strip() != "0"
        if not sp and FX and _auto:
            try:
                _ann = FX.find_announcement_on(_bd, KZ_DATA) or {}
            except Exception as _e:
                _ann = {}
                _why = "定位人行公告页出错：%s" % str(_e)[:100]
            if _ann.get("url"):
                res["fxAnnounceUrl"] = _ann["url"]
                sp = FX.shot(_bd, _ann["url"], KZ_DATA)
                if not sp:
                    _why = ("已定位到该日公告页 %s，但无头截图失败"
                            "（浏览器未启动 / 页面白屏 / 超时），请手工打开该页截图"
                            % _ann["url"])
                    res["c02aShotError"] = _why
            elif not _why:
                _why = ("人行栏目页里没有 %s 的中间价公告（该日未发布 / 非交易日，"
                        "或站点访问失败）；不能借用相邻日期的截图，请手工到 %s 找该日公告"
                        % (_bd, FX_PBC_INDEX))
        elif not sp:
            _why = ("自动截图已关闭（KZ_FX_AUTOSHOT=0），请手工截 %s 的人行公告页"
                    % _bd)
        if sp:
            res["c02aShot"] = sp
        else:
            # 平台截不了图（分发包没有浏览器），就把这一步钉成一张任务卡交给使用它的 AI。
            task = _fx_shot_task(ws, fx, meta, got, res.get("c02aNotesEndRow"),
                                 why=_why, ann_url=(_ann.get("url") or ""))
            if task:
                res["fxShotTask"] = task
                _add_task(res, id="c02a-fx-shot", type="screenshot",
                          title="C.02a 汇率公告截图待补（%s）" % task["date"],
                          target="C.02a 外币测算", saveTo=task["saveTo"],
                          url=task.get("url") or "", prompt=task.get("prompt") or "",
                          blocking=True)
                res.setdefault("warnings", []).append(
                    "C.02a 汇率公告截图待补（%s）：见该表 B%d 批注的操作步骤，"
                    "AI 截好图重新生成一次即自动插入%s"
                    % (task["date"], task["row"],
                       ("（原因：%s）" % _why) if _why else ""))
                # 同时进「检查表 → 待确认清单」：截图缺失是证据缺失，
                # 只写在批注里，复核人翻不到。
                res.setdefault("gaps", []).append({
                    "sheet": name, "row": task["row"], "col": 2,
                    "field": "汇率公告截图", "company": "", "subject": "", "acct": "",
                    "reason": ("汇率公告截图未取得：" + (_why or "请按 C.02a 批注步骤补齐")
                              + "；目标文件 "
                              + (os.path.basename(task.get("saveTo") or "") or "（见批注）")),
                })

    out = {"sheet": name, "status": "ok", "currencies": seq,
           "needRate": need_rate, "rows": len(seq),
           "fxSource": fx.get("url") or "", "fxDate": fx.get("date") or "",
           # 来源种类：PBC（人行公告）/ CFETS（交易中心兜底）。截图能不能自动截取决于它，
           # 结果里直接给出来，省得排查时只看到"没有图"。
           "fxSourceKind": fx.get("source") or "",
           # 按资产负债表日定位到的公告页（没定位到就是空），排查"为什么没截图"时先看它
           "fxAnnounceUrl": res.get("fxAnnounceUrl") or "",
           "fxRates": got}
    if fx_err:
        # 环境缺件/网络不通都要说清楚，不能只是"汇率列空着"让人猜
        out["fxError"] = fx_err
        res.setdefault("warnings", []).append(
            "C.02a 汇率未自动取到：%s（表已生成，汇率列留空标黄，请人工填当日中间价）" % fx_err)
    if res.get("fxShotTask"):
        out["fxShotTask"] = res["fxShotTask"]
    if need_rate:
        out["reason"] = "以下币种需人工填汇率：%s" % "、".join(need_rate)
    elif len(seq) == 1 and seq[0] == "CNY":
        out["reason"] = "全部账户为本位币，C.02a 无外币需测算"
    elif fx.get("rates"):
        # 汇率时点指纹：差异明显时，反推账面用的是哪一天的中间价。
        # 这一次就是靠人工比对才查出"账套其实是 6 月末、资产负债表日填的 9-30"，
        # 机器能算的事不该留给人。
        hint = _fx_date_hint(res.get("rows") or [], fx,
                             (meta or {}).get("balance_date") or "", KZ_DATA)
        if hint:
            out["reason"] = hint
            out["fxHint"] = hint
    return out


def _fx_shot_task(ws, fx, meta, got, anchor_row, why="", ann_url=""):
    """汇率公告截图没拿到时，在 C.02a 留一张"任务卡 + 提示词"，交给使用 PTW 的 AI。

    **为什么必须交给 AI**：分发给客户的 ptW 运行时是 python embeddable，没有浏览器驱动、
    也没有 Pillow，平台既打不开网页也嵌不了图。这一步天生要一个能上网、能截屏的智能体来做。

    平台能做的是把「取哪天的、去哪取、截什么、存到哪个路径、怎么自检、做完做什么」
    钉成一段不用猜的话，写在底稿里——接手的人（或 AI）照着做就行，不需要回头问。
    下一条命令有浏览器时，`FX.shot()` 会直接捡起这个文件，插到 Notes 下方，本文随之消失。
    """
    # 本函数是自己一段逻辑，拿不到 fill_c02a 里的局部 FX，必须自己导入一次；
    # 少这一句时光是"留任务卡"这一步就会 NameError，C.02a 整表算失败（V1.2.86-89 bug）。
    try:
        from core import fxrate as FX
    except Exception:
        FX = None
    bd = str((meta or {}).get("balance_date") or "")
    if not bd:
        return None
    ymd = bd.replace("-", "")
    save_to = os.path.join(KZ_DATA, "fx", "pbc_%s.png" % ymd)
    # 日期一律用**资产负债表日**：截图要证明的是这一天的汇率。
    # 不能用 fx.get("date")——那是汇率实际取自哪一天（非交易日时会退到上一个
    # 交易日，比如 10-01 拿到的是 9-30 的公告），写成 9 月 30 日会让补图的人
    # 存错文件名，程序下次还是找不到。
    day = bd
    # 入口优先级：已定位到的**该日公告页** > 人行公告原文（PBC）> 栏目页。
    # 绝不能把交易中心页面当入口：那是访问当天的数据，截出来日期对不上。
    url = ann_url or (fx.get("url") if (fx.get("source") or "") == "PBC" else "")
    url = url or FX_PBC_INDEX
    rates = "、".join(got[:9]) if got else "见上表 D 列"

    prompt = (
        "【汇率公告截图 · 待补任务】\n"
        "对象：C.02a 外币测算 —— %s 人民币汇率中间价公告截图\n"
        "现状：汇率已自动取到（%s），但公告截图未取得。\n"
        "未自动截的原因：%s\n"
        "平台自身不一定能联网截图，需由使用 PTW 的 AI（或人工）完成。\n\n"
        "一、去哪里取\n"
        "  人民银行「人民币汇率中间价」栏目：\n"
        "  %s\n"
        "  找到标题含「%s」的《中国外汇交易中心受权公布人民币汇率中间价公告》并打开。\n"
        "  栏目页只列最近约 20 个工作日，日期较早时会翻页翻不到——改用站内搜索\n"
        "  「%s 人民币汇率中间价」。备用数据页（优先级低于人行原文）：\n"
        "  https://www.chinamoney.com.cn/chinese/bkccpr/\n\n"
        "二、截什么\n"
        "  整页截图，必须同时满足两条：标题里能看到日期「%s」；正文里能看到币种与汇率。\n\n"
        "三、存哪里（路径和文件名都不能改）\n"
        "  %s\n"
        "  没有浏览器自动化工具时，装 Edge/Chrome 的机器可直接跑这一条（路径已按本机生成）：\n"
        "  %s\n\n"
        "四、截完自检（三条全过才算完成）\n"
        "  1) 文件存在且大于 20KB；\n"
        "  2) 图上日期 = %s；\n"
        "  3) 图上能看清与下表一致的汇率（%s）。\n"
        "  对不上就重截。**不许改本表的汇率，也不许拿别的日期、别的网站截图充数。**\n\n"
        "五、收尾（这一步不做等于没做）\n"
        "  重新生成一次货币资金底稿。程序发现该文件后会自动把截图插到 Notes 下方，\n"
        "  本占位提示随之消失，不需要手工删除。\n\n"
        "做不到时：把这一行文字改成「截图未取得：<具体原因>」，并在生成结果里报告，\n"
        "不要留空、不要用相似日期的截图代替。"
    ) % (day, rates, why or "程序未尝试自动截图", url, day, day, day, save_to,
         (FX.shot_cmd(day, url, KZ_DATA) if FX else ""), day, rates)

    r = int(anchor_row or 21) + 1
    cell = ws.cell(r, 2)
    cell.value = ("【待补】汇率公告截图（%s）—— 操作步骤见本格批注；"
                  "目标文件：data\\fx\\pbc_%s.png" % (day, ymd))
    cell.font = Font(name="Arial", size=10, bold=True, color="C00000")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.fill = PatternFill("solid", fgColor="FFF2CC")
    cell.comment = Comment(prompt, "PTW")
    try:
        cell.comment.width = 520
        cell.comment.height = 380
    except Exception:
        pass

    return {"id": "fx_shot", "sheet": ws.title, "row": r, "file": "C.02a 外币测算",
            "title": "补汇率公告截图（%s）" % day,
            "date": day, "url": url, "saveTo": save_to, "prompt": prompt,
            "afterDone": "重新生成一次货币资金底稿，截图会自动插入 Notes 下方"}


#: 人行「人民币汇率中间价」栏目页（任务卡里给 AI 的入口）
FX_PBC_INDEX = ("https://www.pbc.gov.cn/zhengcehuobisi/125207/125217/125925/index.html")


def _fx_date_hint(rows, fx, balance_date, cache_dir, max_back=4):
    """反推账面折算汇率对应哪一天。返回提示文字，判断不了返回空串。

    只在一个前提下做：外币账面金额按币种汇总后，隐含汇率与资产负债表日的中间价
    偏差超过 0.3%。偏差很小时是正常的期末重估误差，不值得提示。

    候选日期取资产负债表日所在月及往前各月的月末（最多 max_back 个），
    命中判据是隐含汇率与候选日中间价的相对误差 < 0.1%。
    """
    if not balance_date:
        return ""
    from core import fxrate as FX

    agg = {}
    for r in (rows or []):
        code = cur_code(r.get("currency"))
        if not code or code == "CNY":
            continue
        orig = r.get("orig") or 0
        loc = r.get("local") or 0
        if abs(orig) < 0.005 or abs(loc) < 0.005:
            continue
        a = agg.setdefault(code, [0.0, 0.0])
        a[0] += orig
        a[1] += loc
    if not agg:
        return ""

    # 取账面金额最大的那个币种做指纹，避免小币种的零星账户干扰
    code = max(agg, key=lambda k: abs(agg[k][1]))
    orig, loc = agg[code]
    book = loc / orig
    bal = FX.rate_for(code, fx)
    if not bal or abs(book / bal - 1) <= 0.003:
        return ""

    # 候选：往前的月末
    try:
        import datetime as _dt
        y, m, _d = [int(x) for x in str(balance_date).split("-")[:3]]
        cands = []
        for i in range(max_back):
            mm = m - 1 - i
            yy = y
            while mm <= 0:
                mm += 12
                yy -= 1
            nxt = _dt.date(yy + (1 if mm == 12 else 0), 1 if mm == 12 else mm + 1, 1)
            cands.append((nxt - _dt.timedelta(days=1)).strftime("%Y-%m-%d"))
    except Exception:
        return ""

    for iso in cands:
        try:
            # 指纹识别只是"猜它用的是哪天"，不值得为它回退 7 天、把一次生成拖成几分钟；
            # 精确日期不中就换下一个候选。
            other = FX.fetch(iso, cache_dir=cache_dir, timeout=8, back_days=0)
        except Exception:
            other = {}
        v = FX.rate_for(code, other)
        if v and abs(book / v - 1) < 0.001:
            return ("账面折算汇率（%s %s）与 %s 的中间价一致，与资产负债表日 %s 的 %s 不同"
                    "（差异 %+.1f%%）：请确认账套期末是否就是资产负债表日，"
                    "或期末是否已按即期汇率重估。"
                    % (code, ("%.6g" % book), other.get("date") or iso,
                       balance_date, ("%.6g" % bal), (book / bal - 1) * 100))

    return ("账面折算汇率（%s %s）与资产负债表日 %s 的中间价 %s 相差 %+.1f%%，"
            "且与近 %d 个月月末的中间价都不一致：可能是期末未按即期汇率重估，"
            "请确认是否需要调整汇兑损益。"
            % (code, ("%.6g" % book), balance_date, ("%.6g" % bal),
               (book / bal - 1) * 100, max_back))


def _fit_rows(ws, hdr, total_row, need, max_col):
    """把 [hdr+1, total_row) 之间的数据行数量调成 need：不够插、多了删。

    为什么必须能删：模板给 C.02a 预置 USD/CNY/JPY 三行，本期没有日元时留着那行
    就是个空行——用户会以为是漏填，合计行还被往下推。只插不删是半个功能。

    openpyxl 的 insert_rows / delete_rows **都不移动合并单元格**，
    数据区下面的 Notes 说明框必须自己按 ±n 搬，否则会压到数据上。
    """
    if not total_row:
        return total_row
    preset = total_row - hdr - 1
    delta = need - preset
    if delta == 0:
        return total_row
    merges, heights = [], {}
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row > total_row:
            merges.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col))
            ws.unmerge_cells(str(mc))
    for r in list(ws.row_dimensions.keys()):
        h = ws.row_dimensions[r].height
        if r > total_row and h:
            heights[r] = h
    if delta > 0:
        ws.insert_rows(total_row, delta)
        for r in range(total_row, total_row + delta):
            _copy_row_style(ws, total_row - 1, r, max_col)
    else:
        # 删数据区末尾多出来的行（不删合计行本身）
        ws.delete_rows(total_row + delta, -delta)
    for (a, b, c, d) in merges:
        ws.merge_cells(start_row=a + delta, start_column=c,
                       end_row=b + delta, end_column=d)
    for r, h in heights.items():
        try:
            ws.row_dimensions[r + delta].height = h
        except Exception:
            pass
    return total_row + delta


def _normalize_c02a_table(ws, hdr, total, c1, c2):
    """C.02a 整块：表内不画格线，只留最外一圈；数据行格式逐行统一。

    模板每格都带细边框，插行改币种之后线框不连续，看着像两张表拼起来。
    这里把内部边框全部清掉，只给外圈加框——"看起来是一张表"，而不是"一堆格子"。
    """
    thin = Side(style="thin", color="808080")
    none = Side()          # 空 Side 而不是 None：None 读回来会让 border.left.style 直接报错
    body_h = ws.row_dimensions[hdr + 1].height or 13.85

    # 表头：保留模板字体，只重画外圈上/左右
    for c in range(c1, c2 + 1):
        ws.cell(hdr, c).border = Border(
            left=thin if c == c1 else none,
            right=thin if c == c2 else none,
            top=thin)

    for r in range(hdr + 1, total + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name="Arial", size=10)
            cell.border = Border(
                left=thin if c == c1 else none,
                right=thin if c == c2 else none,
                bottom=thin if r == total else none)
            if c == c1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            if c != c1:
                cell.number_format = "0.0000" if c == c1 + 2 else MONEY_FMT
            else:
                cell.number_format = "General"
        ws.row_dimensions[r].height = body_h


def _write_fx_note(ws, fx, got):
    """把汇率来源写进 Notes 下面那个大合并框，返回合并框末行（供插图定位）。

    为什么不写在标签行：标签行只有一行高，长文本会溢出、盖住下面的公告截图；
    模板本来就在 Notes 下面留了一个合并框（B19:F21 之类），那里才是放说明的地方。
    图片要锚在**合并框下面**，否则会压在框上。
    """
    url = fx.get("url") or ""
    day = fx.get("date") or ""
    rates = "、".join(got[:8]) if got else "见下表"
    src = fx.get("source") or ""
    note = ("Notes 我们使用%s中国人民银行授权中国外汇交易中心公布的人民币汇率中间价"
            "进行汇率折算（%s）。来源：%s" % (day or "资产负债表日", rates, url))
    # 该日没有发布中间价（周末/休市）时，取数会往前退到最近一个工作日。
    # 换日期是重大判断，必须写在表上，不能只留在程序里。
    if fx.get("fallbackDays"):
        note = ("Notes 资产负债表日（%s）无人民币汇率中间价，采用最近一个工作日 %s 的中间价"
                "（%s）进行汇率折算（%s）。来源：%s"
                % (fx.get("requestedDate") or "", day, src or "中国外汇交易中心", rates, url))

    label = 0
    for r in range(1, min(ws.max_row, 40) + 1):
        if "Notes" in T(ws.cell(r, 2).value):
            label = r
            break

    box = None
    for mc in sorted(ws.merged_cells.ranges, key=lambda x: x.min_row):
        if mc.min_row >= max(label, 1) and (mc.max_col - mc.min_col) >= 2:
            box = mc
            break

    if box:
        cell = ws.cell(box.min_row, box.min_col)
        cell.value = note
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        if label:
            ws.cell(label, 2).value = "Notes"
        return box.max_row

    # 没有合并框就退回标签行（老模板可能没有那个框）
    if label:
        cell = ws.cell(label, 2)
        cell.value = note
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    return label


# ---------------------------------------------------------------- 入口
def fill(wb, res, meta=None):
    """按 Skill 的四张派生表依次执行，每张独立 try。

    一张表失败不该让另外三张也白做——派生表之间没有依赖，
    但共用同一个 workbook，所以异常要就地吞掉并记进返回。
    """
    out = {}
    for key, fn in (("c01b", fill_c01b), ("c00b", fill_c00b),
                    ("c02", fill_c02), ("c02a", fill_c02a)):
        try:
            out[key] = fn(wb, res, meta)
        except Exception as e:
            out[key] = {"status": "error", "reason": "%s: %s" % (type(e).__name__, str(e)[:160])}
    return out
