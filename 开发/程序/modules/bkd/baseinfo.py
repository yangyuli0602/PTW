# -*- coding: utf-8 -*-
"""项目基础信息（客户名称 / 资产负债表日 / 会计准则 / 记账本位币 / TE / SAD）。

**为什么要单独立一份**：这些是"全底稿共用"的抬头信息，货币资金、销管费用、
长期资产每份底稿的 Lead 封面都要写同一套值。原先各模块各传各的参数：
货币资金这边只接到客户名和日期两个参数、还都没落地，封面四格全空；
费用那边自己写了一套 TE 落盘逻辑。同一个数两处实现，迟早对不上。

**存放**：``<数据目录>/project_info.json``。改一次，所有底稿下次生成都跟着变。

**写入位置不写死**：按标签文字（客户名称 / 资产负债表日 / 可容忍误差 / 名义金额 /
适用会计准则 / 记账本位币）在 Lead 页里现场找，值写在标签右侧第一个空格。
模板改版、行号下移都不影响；找不到某个标签就记进返回结果，不猜。
"""

import io
import json
import os
import time

#: 字段 → (中文名, 提示语)
FIELDS = (
    ("customer", "客户名称", "写进 Lead 封面；表 1 的账套名称列也跟着它"),
    ("balance_date", "资产负债表日", "所有日期的唯一来源：表头、上期列头、C.03 截止期间"),
    ("standard", "适用会计准则", "如「中国企业会计准则」"),
    ("currency", "记账本位币", "如「人民币」"),
    ("te", "可容忍误差(TE)", "认定阈值、截止测试阈值的推导起点"),
    ("sad", "名义金额(SAD)", "留空按 TE×10% 推"),
)

#: 首次运行时的默认值：**全部留空**。
#: 不编客户名、不编准则、不硬编码 TE —— 猜错比空着更难发现：
#: 空着时 read_te_sad 会走它自己的兜底并留 TE_NOT_READ 告警，看得见；
#: 写一个假值进去，告警就没了，人还以为项目组填过。
DEFAULTS = {
    "customer": "",
    "balance_date": "",
    "standard": "",
    "currency": "",
    "te": None,
    "sad": None,
    "cra": {},
}

#: 标签文字（同一含义的几种写法都认）
LABELS = {
    "customer": ("客户名称", "被审计单位"),
    "balance_date": ("资产负债表日",),
    "te": ("可容忍误差",),
    "sad": ("名义金额",),
    "standard": ("适用会计准则", "会计准则"),
    "currency": ("记账本位币",),
}

# ---------------------------------------------------------------- CRA（认定风险）
#
# 模板 D14:D18 原来是写死的 `=IF(C14="Minimal",$C$4*100%,IF(C14="Low",$C$4*75%,...))`，
# 百分比固化在公式里，换个项目组/换类底稿就得逐格改公式 —— 用户要求改成可配置。
#
# 现在：百分比存在 CRA 表里（每类底稿一张），公式改成 VLOOKUP 查这张表。
# 改百分比只改表，公式自动跟随；不同科目各有各的表，互不影响。

#: 五种 CRA 等级（与模板 C14:C18 的数据验证列表一致）
CRA_TYPES = ("Minimal", "Low", "Moderate", "High", "N/A")

#: 模板原公式里的默认百分比，作为未配置时的初始值
CRA_DEFAULT_PCT = {"Minimal": 1.00, "Low": 0.75, "Moderate": 0.50, "High": 0.25}

#: 五项认定：内部键 → (模板标签关键字, 显示名)
ASSERTIONS = (
    ("completeness", "完整", "完整性 (C)"),
    ("existence", "存在", "存在性 (E)"),
    ("valuation", "计价", "计价（V）"),
    ("rights", "权利", "权利和义务（R&O）"),
    ("presentation", "列报", "列报和披露（P&D）"),
)

#: 销管费用 BKD 的四项认定。与货币资金不是同一套：费用用「发生」不用「存在」，
#: 用「计量」不用「计价/权利」，模板 C8:C11 的标签就是这么写的。
EXP_ASSERTIONS = (
    ("completeness", "完整", "完整性 (C)"),
    ("occurrence", "发生", "发生 (O)"),
    ("measurement", "计量", "计量（M）"),
    ("presentation", "列报", "列报和披露（P&D）"),
)

#: 所有合法的认定键（落盘校验用；各模块只展示自己那几项）
ASSERTION_KEYS = tuple(dict.fromkeys(
    [k for k, _w, _l in ASSERTIONS] + [k for k, _w, _l in EXP_ASSERTIONS]))

#: 默认科目类型（本期只做货币资金；费用、长期资产接进来时在这里加）
CRA_MODULE_DEFAULT = "货币资金"

#: 已接入 CRA 配置的底稿类型。每类一张 CRA 表，互不影响
CRA_MODULES = ("货币资金", "销管费用")

#: 各模块的默认百分比。**不要强行统一**：模板原公式就是两套
#: （货币资金 100/75/50/25，费用 50/25/15/10），改默认值等于改现有底稿的阈值。
CRA_DEFAULT_PCT_BY_MODULE = {
    "货币资金": {"Minimal": 1.00, "Low": 0.75, "Moderate": 0.50, "High": 0.25},
    "销管费用": {"Minimal": 0.50, "Low": 0.25, "Moderate": 0.15, "High": 0.10},
}


def default_pcts(module=CRA_MODULE_DEFAULT):
    """某模块的默认百分比（没配过时用）。"""
    return dict(CRA_DEFAULT_PCT_BY_MODULE.get(module) or CRA_DEFAULT_PCT)


def assertions_of(module=CRA_MODULE_DEFAULT):
    """某模块的认定项：销管费用四项，其余走通用五项。"""
    return EXP_ASSERTIONS if module == "销管费用" else ASSERTIONS


def cra_sheet_name(module):
    """每类底稿一张 CRA 表，表名固定可推导。"""
    return "CRA_%s" % module


def cra_of(info, module=CRA_MODULE_DEFAULT):
    """取某类底稿的 CRA 配置：{pcts: {...}, assertions: {key: 类型}}。

    没配过就返回模板默认百分比 + 空选择（不替项目组判断风险等级）。
    """
    all_cfg = (info or {}).get("cra") or {}
    cfg = all_cfg.get(module) or {}
    pcts = default_pcts(module)
    for k, v in (cfg.get("pcts") or {}).items():
        try:
            pcts[k] = float(v)
        except (TypeError, ValueError):
            pass
    pcts["N/A"] = None
    assertions = {}
    # 按**该模块**的认定项过滤：费用是 完整/发生/计量/列报，
    # 拿货币资金那五项来过滤会把 occurrence/measurement 直接丢掉（存了等于没存）
    for key, _kw, _label in assertions_of(module):
        v = (cfg.get("assertions") or {}).get(key)
        if v in CRA_TYPES:
            assertions[key] = v
    return {"module": module, "pcts": pcts, "assertions": assertions,
            "configured": bool(cfg)}


def _norm_cra(raw):
    """把落盘的 CRA 配置规整成 {模块: {pcts: {...}, assertions: {...}}}。

    容错比严格重要：这是人手改过的 json，可能缺字段、写成字符串百分比、
    甚至模块名写错。规整不了的项直接丢，不抛异常——否则一份坏配置会让整次生成失败。
    """
    out = {}
    if not isinstance(raw, dict):
        return out
    for mod, cfg in raw.items():
        if not isinstance(cfg, dict):
            continue
        pcts = {}
        for k, v in (cfg.get("pcts") or {}).items():
            if k not in CRA_TYPES or k == "N/A":
                continue
            try:
                f = float(str(v).replace("%", "").strip())
            except (TypeError, ValueError):
                continue
            # 允许写成 75 或 0.75 两种习惯
            pcts[k] = f / 100.0 if f > 1 else f
        asserts = {}
        # 用「所有合法认定键」校验，不能只认货币资金那五项 —— 费用的
        # occurrence/measurement 会被当成非法键丢掉，配置存了等于没存。
        for key in ASSERTION_KEYS:
            v = (cfg.get("assertions") or {}).get(key)
            if v in CRA_TYPES:
                asserts[key] = v
        if pcts or asserts:
            out[str(mod)] = {"pcts": pcts, "assertions": asserts}
    return out


def _merge_cra(old, patch):
    """按模块合并，不清掉没提到的模块。"""
    out = dict(old or {})
    for mod, cfg in (patch or {}).items():
        if not isinstance(cfg, dict):
            continue
        cur = dict(out.get(mod) or {})
        cur["pcts"] = dict(cur.get("pcts") or {})
        cur["pcts"].update(cfg.get("pcts") or {})
        cur["assertions"] = dict(cur.get("assertions") or {})
        cur["assertions"].update(cfg.get("assertions") or {})
        out[mod] = cur
    return _norm_cra(out)


def path_of(data_dir):
    return os.path.join(data_dir, "project_info.json")


def load(data_dir):
    """读基础信息；文件不存在就返回默认值（不落盘，避免启动时写数据目录）。"""
    info = dict(DEFAULTS)
    p = path_of(data_dir)
    try:
        if os.path.isfile(p):
            with io.open(p, encoding="utf-8") as fh:
                d = json.load(fh) or {}
            for k in DEFAULTS:
                if k == "cra":
                    continue
                if d.get(k) not in (None, ""):
                    info[k] = d[k]
            info["cra"] = _norm_cra(d.get("cra"))
            info["updatedAt"] = d.get("updatedAt", "")
    except Exception:
        pass
    return info


def save(data_dir, patch):
    """合并写入。传 None 表示不改；传空串表示清空该字段。"""
    info = load(data_dir)
    for k, v in (patch or {}).items():
        if k == "cra":
            continue
        if k in DEFAULTS and v is not None:
            info[k] = v
    if isinstance((patch or {}).get("cra"), dict):
        info["cra"] = _merge_cra(info.get("cra") or {}, patch["cra"])
    for k in ("te", "sad"):
        try:
            info[k] = float(info[k]) if info.get(k) not in (None, "") else None
        except (TypeError, ValueError):
            info[k] = None
    # TE/SAD 联动：只给 TE 时 SAD 按 10% 跟着走，除非显式指定了 SAD
    if patch and "te" in patch and not (patch or {}).get("sad"):
        if info.get("te"):
            info["sad"] = round(info["te"] * 0.1, 2)
    info["updatedAt"] = time.strftime("%Y-%m-%d %H:%M:%S")
    os.makedirs(data_dir, exist_ok=True)
    with io.open(path_of(data_dir), "w", encoding="utf-8") as fh:
        json.dump(info, fh, ensure_ascii=False, indent=1)
    return info


def _cell_text(v):
    if v is None:
        return ""
    return str(v).replace(" ", "").replace("\u3000", "").replace("\n", "")


def _col_letter(c):
    s = ""
    while c > 0:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def fill_lead_header(wb, info, only=None, overwrite=False):
    """把基础信息写进每张 Lead 页的封面区。返回 {sheet: {cells, missing}}。

    - 找标签：扫前 12 行、前 4 列，命中标签文字后取**同行右侧第一个空格**；
    - ``overwrite=False`` 时不覆盖已有值（保护项目组手填的内容）；
    - 找不到标签 → 记进 missing，不猜位置。
    """
    done = {}
    if not info:
        return done
    names = [n for n in wb.sheetnames
             if "Lead" in n and (only is None or n in only)]
    for name in names:
        ws = wb[name]
        got, missing = {}, []
        for key, kws in LABELS.items():
            val = info.get(key)
            if val in (None, ""):
                continue
            if key in ("te", "sad") and not isinstance(val, (int, float)):
                try:
                    val = float(val)
                except (TypeError, ValueError):
                    continue
            hit = None
            for r in range(1, min(ws.max_row, 12) + 1):
                for c in range(1, 5):
                    t = _cell_text(ws.cell(r, c).value)
                    if not t or not any(k in t for k in kws):
                        continue
                    for cc in range(c + 1, min(ws.max_column, c + 4) + 1):
                        v = ws.cell(r, cc).value
                        if v in (None, "") or overwrite:
                            hit = (r, cc)
                            break
                    if hit:
                        break
                if hit:
                    break
            if not hit:
                missing.append(key)
                continue
            r, c = hit
            cell = ws.cell(r, c, val)
            if key in ("te", "sad"):
                # 金额格式与 core 同一套（会计专用、无货币符号、零显示 -）。
                # 这里局部导入而不是模块顶层导：core 是在函数里延迟导 baseinfo 的，
                # 顶层互导容易踩循环导入。
                try:
                    from .core import ACCT_FMT as _AF
                except ImportError:
                    from core import ACCT_FMT as _AF
                cell.number_format = _AF
            elif key == "balance_date":
                cell.number_format = "yyyy-mm-dd"
            got[key] = "%s%d" % (_col_letter(c), r)
        done[name] = {"cells": got}
        if missing:
            done[name]["missing"] = missing
    return done


def merged(data_dir, **override):
    """基础信息 + 本次传入的覆盖值（传了就以传的为准）。"""
    info = load(data_dir)
    for k, v in override.items():
        if k in DEFAULTS and v not in (None, ""):
            info[k] = v
    return info
