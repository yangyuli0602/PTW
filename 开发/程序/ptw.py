#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""PTW 命令行入口 —— 给 Agent（Qoder 等）用，不需要起 HTTP 服务。

为什么要有它：Agent 做「起服务 → 记端口 → 拼 JSON → 解析返回 → 下载文件」这一串，
每一步都可能失败，而且端口占用、单实例锁这类失败跟业务毫无关系。一条命令解决。

用法（在 程序 目录的上一级，或直接用绝对路径调用）：

    python 程序/ptw.py state
    python 程序/ptw.py import 科目余额表.xlsx
    python 程序/ptw.py cash  --date 2026-10-31 --prior 上年BKD.xlsx
    python 程序/ptw.py exp   --from 1 --to 6
    python 程序/ptw.py check --from 1 --to 6

统一输出单行 JSON：
    成功 {"ok": true, ...}
    失败 {"ok": false, "code": "TB_EMPTY", "error": "...", "next": "..."}

code 是稳定的英文标识，Agent 按 code 分支；next 是给人看的一句话下一步。
所有命令都接受 --data-dir 指定数据目录（默认 程序/../data），--pretty 格式化输出。
"""
import argparse
import json
import os
import sys

APP = os.path.dirname(os.path.abspath(__file__))

#: 把底层的中文报错归到稳定 code。Agent 靠 code 分支，不靠理解中文。
ERR_RULES = [
    ("科目余额表（tb 表）是空的", "TB_EMPTY", "先跑 `ptw.py import <科目余额表>`"),
    ("科目余额表里没有「库存现金", "NO_CASH_ACCT",
     "清洗时加「提取一级科目」，或检查一级科目是否识别成货币资金"),
    ("科目余额表里没有货币资金账户", "NO_CASH_ACCT",
     "确认 tb 已合并，且清洗后一级科目是 库存现金/银行存款/其他货币资金"),
    ("未找到内置模板", "TPL_MISSING", "data/bkd_tpl 缺失，从交付包拷回来"),
    ("找不到费用模板", "TPL_MISSING", "data/bkd_tpl 缺失，从交付包拷回来"),
    ("序时账里没有取到「销售费用", "NO_EXP_DATA",
     "检查 --from/--to 期间设置，或确认序时账已导入并合并"),
    ("没有取到", "NO_DATA", "先 `ptw.py state` 看数据在不在"),
    ("另一个程序正在使用此文件", "DB_LOCKED",
     "PTW 平台正开着同一个数据目录。先关掉平台（或改用 --data-dir 指向别的目录）"),
    ("already open in", "DB_LOCKED",
     "PTW 平台正开着同一个数据目录。先关掉平台再跑 CLI"),
    ("受限金额的账号不唯一", "RESTRICTED_AMBIGUOUS",
     "同一账号同时挂在银行存款和其他货币资金下，改成 `--restricted 其他货币资金|账号=金额`"),
    ("受限资金没能写入", "RESTRICTED_AMBIGUOUS",
     "写成 `--restricted 一级科目|账号=金额`，科目名用 BKD 里那一列的值"),
]


class CliError(Exception):
    def __init__(self, code, msg, next_=""):
        Exception.__init__(self, msg)
        self.code, self.msg, self.next = code, msg, next_


def emit(obj, pretty=False):
    # default=str 兜底：数据里混进 date/Decimal 时不能让整个命令挂掉，
    # 宁可字段变成字符串，也不能让 Agent 拿不到结果。
    print(json.dumps(obj, ensure_ascii=False, default=str,
                     indent=2 if pretty else None))


def classify(err_text):
    """中文报错 -> (code, next)。没命中规则就给一个通用码。"""
    s = str(err_text or "")
    for key, code, nxt in ERR_RULES:
        if key in s:
            return code, nxt
    return "UNKNOWN", "看 error 原文；若与文件/权限相关，先 `ptw.py state` 确认数据在不在"


def _platform_version():
    """读 server.py 的 VERSION —— 版本号只有一个源头，不在别处写第二份。"""
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "server.py")
    try:
        with open(p, encoding="utf-8") as f:
            for line in f:
                if line.startswith("VERSION"):
                    return line.split("=", 1)[1].strip().strip("\"'")
    except Exception:
        pass
    return ""


def boot(data_dir=None):
    """按 --data-dir 设好 KZ_DATA_DIR 再 import core。

    core.py 的 KZ_DATA / OUT / TPL_DIR 是模块级常量，import 那一刻就定死了，
    所以环境变量必须先设，import 必须放在函数里（延迟导入）。
    """
    if data_dir:
        os.environ["KZ_DATA_DIR"] = os.path.abspath(data_dir)
    if APP not in sys.path:
        sys.path.insert(0, APP)
    from core.db import Ledger
    return Ledger(APP)


# ---------------------------------------------------------------- state
def cmd_state(a):
    lg = boot(a.data_dir)
    st = lg.state() or {}
    stats = st.get("stats") or {}
    rows = {}
    for k, v in stats.items():
        rows[k] = (v.get("rows") if isinstance(v, dict) else v) or 0
    merged = {k: rows.get(k, 0) > 0
              for k in ("tb", "journal", "flow", "fa", "fa_beg")}
    if merged["tb"]:
        hint = "可以直接 ptw.py cash"
        if merged["fa"] and merged["fa_beg"]:
            hint += " / ptw.py asset（长期资产·SWP）"
        elif merged["fa"]:
            hint += "；固定资产台账只导了期末，缺期初会让 CHECK 不平，请补 fa_beg"
    else:
        hint = "tb 还没合并：先 ptw.py import <科目余额表>"
    return {
        "ok": True,
        "data_dir": lg.data_dir,
        "db": lg.db_path,
        "rows": rows,
        "merged": merged,
        "ready": {"cash": merged["tb"], "exp": merged["journal"],
                  "fa": bool(merged["fa"] and merged["fa_beg"])},
        "sheets": [{"sheet": s.get("sheet_name"), "file": s.get("file_name"),
                    "target": s.get("target"), "rows": s.get("n_rows")}
                   for s in (st.get("sheets") or [])],
        "hint": hint,
    }


# ---------------------------------------------------------------- import
def _best_target(names):
    """一份表可能同时有科目余额表和凭证两个 sheet，所以逐 sheet 判定而不是按文件判定。

    用各 target 的字段映射命中数打分，最高分且 >=4 才认，否则跳过并如实报告。
    """
    from core.fields import auto_map
    best, best_map, best_n = None, {}, 0
    for t in ("tb", "journal", "flow", "fa"):
        try:
            m = auto_map(t, names) or {}
        except Exception:
            continue
        n = len([k for k, v in m.items() if v])
        if n > best_n:
            best, best_map, best_n = t, m, n
    return (best if best_n >= 4 else None), best_map, best_n


def cmd_import(a):
    lg = boot(a.data_dir)
    if not a.file or not os.path.isfile(a.file):
        raise CliError("FILE_NOT_FOUND",
                       "文件不存在：%s" % a.file,
                       "检查路径；路径含空格或中文时用引号包起来")
    with open(a.file, "rb") as fh:
        f = lg.add_file(os.path.basename(a.file), fh)
    fid = f["file_id"]
    if f.get("duplicate"):
        # 同 md5 复用 file_id：重复导入同一份文件不会产生两份数据
        pass
    insp = lg.inspect(fid, a.target)
    done, skipped = [], []
    used = set()
    for sh in (insp.get("sheets") or []):
        if sh.get("empty"):
            skipped.append({"sheet": sh.get("sheet"), "why": sh.get("note") or "空表"})
            continue
        hr = int(sh.get("header_row") or 1)
        hrows = int(sh.get("header_rows") or 1)
        names = lg._header_names(sh.get("preview") or [], hr, hrows)
        tgt, mp, score = a.target, (sh.get("mapping") or {}), 0
        if not tgt:
            tgt, mp, score = _best_target(names)
        if not tgt:
            skipped.append({"sheet": sh.get("sheet"),
                            "why": "表头不像任何一类数据（%d 列，最高匹配 0）" % len(names)})
            continue
        # 固定资产台账：一个文件常带「期末台账 / 期初台账」两张表，
        # 都导成 fa 会互相覆盖，所以按 sheet 名分流到 fa_beg
        if tgt == "fa" and "期初" in str(sh.get("sheet") or ""):
            tgt = "fa_beg"
        r = lg.import_sheet(fid, sh.get("sheet"), hr, hrows, tgt)
        sid = r["sid"]
        lg.set_mapping(sid, tgt, mp or {})
        # 套默认清洗配方并保存：不做这步，后面的「合并」拿不到干净的列
        try:
            steps = lg.clean_preset(sid, "default")["steps"]
            lg.clean_save(sid, steps)
            lg.clean_apply(sid)
        except Exception as e:
            skipped.append({"sheet": sh.get("sheet"), "why": "清洗失败：%s" % e})
            continue
        used.add(tgt)
        done.append({"sheet": sh.get("sheet"), "target": tgt,
                     "rows": r.get("n_rows"), "score": score,
                     "mapped": len([k for k, v in (mp or {}).items() if v])})
    merged = {}
    for t in sorted(used):
        try:
            merged[t] = lg.merge(t)
        except Exception as e:
            merged[t] = {"error": str(e)}
    if not done:
        raise CliError("NO_SHEET_IMPORTED", "没有导入任何工作表",
                       "看 skipped 的原因；可能是表头行识别错，或文件格式不支持")
    return {"ok": True, "file": f.get("file_name"), "file_id": fid,
            "imported": done, "skipped": skipped, "merged": merged}


# ---------------------------------------------------------------- 生成
def _save(data, name):
    from modules.bkd import core as BKD
    os.makedirs(BKD.OUT, exist_ok=True)
    path = os.path.join(BKD.OUT, name)
    with open(path, "wb") as fh:
        fh.write(data)
    # 固定名副本：`_最新` 永远指向最后一次生成的底稿。
    # 输出目录里堆着几十份带时间戳的文件，按名字分不清哪份是刚生成的，
    # 打开了几小时前那份还会以为"改了没生效"。时间戳文件留作历史。
    # 被 Excel 占用时（同目录有 ~$ 锁文件）Windows 拒绝写入，退到 _最新(2)…
    stem = name.split("_")[0]
    cands = ["%s_最新.xlsx" % stem] + ["%s_最新(%d).xlsx" % (stem, i) for i in range(2, 10)]
    for cand in cands:
        try:
            with open(os.path.join(BKD.OUT, cand), "wb") as fh:
                fh.write(data)
            _save.latest = cand
            return path
        except OSError:
            continue
    _save.latest = ""
    return path


def _load_prior_from(lg, R, path):
    """上年 BKD 底稿 -> 账户清单。没给就返回空，不阻断（上年数是可选信息）。"""
    if not path:
        return [], {"loaded": False, "reason": "未提供上年底稿"}, ""
    if not os.path.isfile(path):
        raise CliError("FILE_NOT_FOUND", "上年底稿不存在：%s" % path, "检查 --prior 路径")
    with open(path, "rb") as fh:
        f = lg.add_file(os.path.basename(path), fh)
    rows, info = R._load_prior(f["file_id"], None)
    return rows, info, f["file_id"]


def _rates(pairs):
    out = {}
    for kv in pairs or []:
        k, _, v = str(kv).partition("=")
        k = k.strip()
        if not k:
            continue
        try:
            out[k] = float(v)
        except ValueError:
            raise CliError("BAD_RATE", "折算率格式不对：%s" % kv, "写成 主体名=7.1 的形式")
    return out


def _restricted(pairs):
    """--restricted "账号=金额"，可重复。只认账号，认不出就报错，不猜。"""
    out = {}
    for kv in pairs or []:
        k, _, v = str(kv).partition("=")
        k = k.strip()
        if not k:
            continue
        try:
            out[k] = float(v)
        except ValueError:
            raise CliError("BAD_RESTRICTED",
                           "受限金额格式不对：%s" % kv,
                           "写成 账号=金额 的形式，如 --restricted 755932641710201=500000")
    return out


def FP_blocking_brief(fp):
    """把阻塞级漂移压成一行，方便 Agent 一眼看到是哪张表。"""
    out = []
    for d in (fp.get("diffs") or []):
        if d.get("level") != "block":
            continue
        out.append("%s:%s" % (d.get("sheet") or "-", d.get("kind")))
    return "、".join(out[:5])


def cmd_cash(a):
    import time as _t
    lg = boot(a.data_dir)
    from modules.bkd import routes as R
    prior, pinfo, _pfid = _load_prior_from(lg, R, a.prior)
    rates = _rates(a.rate)
    cf = None
    if not a.no_cutoff and a.date:
        cf = {"cutoff": "1", "cutoff_date": a.date,
              "cutoff_before": a.cutoff_before, "cutoff_after": a.cutoff_after,
              "cutoff_threshold": a.cutoff_amount}
    _pl = None
    # 上年 Lead 就在上年底稿里：没单独指定时，直接用 --prior 那个文件去找 Lead 表
    _plp = getattr(a, "prior_lead", "") or (a.prior if getattr(a, "prior", "") else "")
    if _plp:
        if not os.path.isfile(_plp):
            raise CliError("PRIOR_LEAD", "上年 Lead 文件不存在：%s" % _plp, "检查 --prior-lead 路径")
        with open(_plp, "rb") as _fh:
            _raw = _fh.read()
        _pl = R._parse_prior_lead_bytes(_raw)
        # 整表复制要原始字节（保留格式），解析结果只是索引
        if isinstance(_pl, dict):
            _pl["data"] = _raw
        if _pl and _pl.get("error"):
            raise CliError("PRIOR_LEAD", "上年 Lead 解析失败：%s" % _pl["error"],
                           "确认该文件里有 C.00 Lead 表，且表1 含「科目名称」「期末审定数」表头")
    data, info = R._run_c00(lg.conn, prior=prior, balance_date=a.date,
                            customer=a.customer, rates=rates, cutoff_fields=cf,
                            strict=a.strict, restricted=_restricted(a.restricted),
                            standard=getattr(a, "standard", "") or "",
                            currency=getattr(a, "currency", "") or "",
                            te=getattr(a, "te", None), sad=getattr(a, "sad", None),
                            prior_lead=_pl)
    if data is None:
        raise CliError(*((lambda c: (c[0], info.get("error") or "生成失败", c[1]))(
            classify(info.get("error")))))
    fp = info.get("fingerprint") or {}
    # TE 门禁（2026-11-12）：读不到就停，不再静默按默认值出稿。
    # 静默兜底的后果是整份底稿的判断阈值集体失准，而结果看起来完全正常。
    if (info.get("teSource") or "") == "default" and not getattr(a, "allow_default_te", False):
        raise CliError(
            "TE_NOT_READ",
            "可容忍误差（TE）没读到：模板没有 PMTE 表、C.00 Lead!C4 为空、命令行也没给 --te。",
            "按 00_总则.md 阶段3 补 --te/--sad 后重跑；确实只能用默认值就加 "
            "--allow-default-te，并在交付说明里写明「TE 为默认值」。")
    if a.strict and fp.get("blocking"):
        bad = FP_blocking_brief(fp)
        raise CliError("FP_MISMATCH",
                       "模板结构有 %d 处漂移，--strict 下不落盘：%s" % (fp["blocking"], bad),
                       "先跑 ptw.py fingerprint 看 diff；确认无误后去掉 --strict 重跑")
    stamp = _t.strftime("%Y%m%d_%H%M%S")
    name = "货币资金BKD_%s.xlsx" % stamp
    path = _save(data, name)
    out = {"ok": True,
           "main": {"file": name, "path": path,
                    "rows": info.get("rows"), "gaps": info.get("gaps"),
                    "amount": info.get("amount"), "rmbAmount": info.get("rmbAmount"),
                    "stats": info.get("stats"), "cutoff": info.get("cutoff"),
                    "bySubject": info.get("bySubject"),
                    "restricted": {"total": info.get("restrictedTotal"),
                                   "mode": info.get("restrictedMode"),
                                   "target": info.get("restrictedTarget")},
                    "derived": info.get("derived") or {},
                    "teSource": info.get("teSource"),
                    "te": info.get("te"), "sad": info.get("sad"),
                    "leadRows": info.get("leadRows") or {},
                    "c02aShotEmbedded": info.get("c02aShotEmbedded"),
                    "c02aShotError": info.get("c02aShotError") or "",
                    "gapList": info.get("gapList") or []},
           "checks": {"fingerprint": info.get("fingerprint"),
                      "invariants": info.get("invariants"),
                      "baseline": info.get("baseline")},
           "prior": pinfo, "warnings": []}
    _latest = getattr(_save, "latest", "") or ""
    out["main"]["latestFile"] = _latest
    if not _latest:
        out["warnings"].append("「_最新」副本写入失败（目录里所有候选名都被占用），"
                              "请先关掉 Excel 里打开的那份底稿再生成。")
    elif not _latest.endswith("BKD_最新.xlsx"):
        out["warnings"].append("「%s」正被 Excel 打开，已另存为 %s；原文件名可能仍是旧内容。"
                               % (_latest.split("_最新")[0] + "_最新.xlsx", _latest))
    if info.get("cutoffError"):
        out["warnings"].append(info["cutoffError"])
    # 一眼可判的验收建议，Agent 不用再读文档
    inv = info.get("invariants") or {}
    bad = [i["id"] for i in (inv.get("items") or []) if i.get("status") == "fail"]
    if bad:
        out["warnings"].append("不变量不通过：%s" % "、".join(bad))
    fp = info.get("fingerprint") or {}
    if fp.get("blocking"):
        out["warnings"].append("模板结构相对上次有 %d 处变化，建议人工复核后再用" % fp["blocking"])
    dv = info.get("derived") or {}
    miss = [k for k, v in dv.items() if isinstance(v, dict) and v.get("status") in ("partial", "manual")]
    if miss:
        out["warnings"].append("派生表需人工补：%s（原因见 main.derived 各表 reason）"
                               % "、".join(miss))
    for w in (info.get("teWarnings") or []):
        out["warnings"].append(w)
    # 主底稿生成过程中的告警（Lead 行号定位失败会退回 V7 兜底，必须让人看见）
    for w in (info.get("buildWarnings") or []):
        out["warnings"].append(w)
    out["tasks"] = info.get("tasks") or []
    out["blockingTasks"] = [t.get("id") for t in out["tasks"] if t.get("blocking")]
    out["next"] = ("；".join([
        "gaps=%s（0 表示无待确认格）" % info.get("gaps"),
        "不变量 %d 项不通过" % len(bad) if bad else "不变量全通过",
        (info.get("baseline") or {}).get("summary") or "已建立基准",
    ]))
    # ---- 文档指针：AI 不用猜「现在该读哪份」。
    # 一致性不能只靠"记得读文档"，把下一步该查的章节直接写进结果里。
    read = ["00_总则.md（阶段状态机）"]
    if (info.get("teSource") or "") == "default":
        read.append("00_总则.md · 阶段3（TE 未确认，本次按 --allow-default-te 放行）")
    if out["blockingTasks"]:
        read.append("00_总则.md · 阶段5（blocking 任务未结清：%s）"
                    % "、".join(out["blockingTasks"]))
    if miss:
        read.append("底稿规则/10_表族与顺序.md §5（派生表未落笔：%s）" % "、".join(miss))
    if bad:
        read.append("底稿规则/12_Lead.md §6（不变量不通过：%s）" % "、".join(bad))
    out["docHint"] = {"entry": "SKILL.md",
                      "read": read,
                      "verify": "程序/tools/wp_diff.py 对 PBC/项目/基线.xlsx"}
    out["version"] = _platform_version()
    return out


def _exp_ai_tasks(tasks):
    """把费用三张任务卡整理成结构化 aiTasks：清单、结果文件、应有/已完成条数、状态。

    为什么单列一份：`tasks[]` 是"要做什么"，这张表是"做到哪一步了"。
    外部 AI 拿它就能直接循环（读 readPath → 写 savePath → 重跑），
    也能用 status 判断结清，不用去解析 warnings 的散文。
    """
    out = []
    for t in (tasks or []):
        rp = t.get("readPath") or ""
        sp = t.get("saveTo") or ""
        exp_n = t.get("expectCount")
        done = None
        if sp and os.path.exists(sp):
            try:
                with open(sp, encoding="utf-8") as fh:
                    d = json.load(fh)
                items = d.get("items") if isinstance(d, dict) else d
                if isinstance(items, list):
                    done = len(items)
            except Exception:
                done = None
        if exp_n is None:
            status = "unknown"
        elif done is None:
            status = "pending"
        elif done >= int(exp_n):
            status = "done"
        else:
            status = "partial"
        out.append({"id": t.get("id") or "", "title": t.get("title") or "",
                    "target": t.get("target") or "",
                    "readPath": rp, "savePath": sp,
                    "expectCount": exp_n, "doneCount": done,
                    "status": status,
                    "afterDone": t.get("afterDone") or ""})
    return out


def cmd_exp(a):
    import time as _t
    lg = boot(a.data_dir)
    from modules.bkd import routes as R
    pm, perr = {}, ""
    prior_oar, prior_shorts, prior_note = {}, {}, ""
    if a.prior:
        if not os.path.isfile(a.prior):
            raise CliError("FILE_NOT_FOUND", "上年费用底稿不存在：%s" % a.prior, "检查 --prior 路径")
        from modules.bkd import expense as EXP
        from modules.bkd import oar_map as OM
        with open(a.prior, "rb") as fh:
            _raw = fh.read()
        pm, perr = EXP.parse_prior_expense(_raw)
        prior_oar, prior_shorts, prior_note = OM.read_prior(_raw)
        if perr:
            raise CliError("PRIOR_UNUSABLE", perr,
                           "确认这是上年费用底稿（含 VC.00/VD.00 BKD 与『上期审定数』列）")
    ev = a.oar_evidence or os.path.join(
        __import__("modules.bkd.core", fromlist=["x"]).OUT, "OAR_evidence_费用.json")
    # 上一期 TOD：整页搬入，预审段引用它（不解析成结构，样式要原样保留）
    # 上一期 TOD 就在上年费用底稿这同一册 WP 里，--prior 给了就用它，不必再传一次。
    _ptod = None
    _pt = getattr(a, "prior_tod", "") or ""
    if _pt:
        if not os.path.isfile(_pt):
            raise CliError("FILE_NOT_FOUND", "上一期 TOD 底稿不存在：%s" % _pt,
                           "检查 --prior-tod 路径")
        with open(_pt, "rb") as fh:
            _ptod = fh.read()
    elif a.prior:
        _ptod = _raw
    # 排除口径：V1.2.79 起回到关键词规则（不再读「排除Mapping」配置表）
    data, info = R._run_exp(lg.conn, a.from_, a.to,
                            balance_date=a.date, customer=a.customer,
                            prior_map=pm or None, pre_end=a.pre_end,
                            oar_evidence=ev,
                            oar_chg_min=a.oar_chg_min,
                            oar_chg_pct=float(a.oar_chg_pct or 0.10),
                            te=a.te, sad=a.sad,
                            standard=a.standard or "", currency=a.currency or "",
                            assurance=[x.strip() for x in (a.assurance or "").split(",")
                                       if x.strip()] or None,
                            prior_oar=prior_oar, prior_shorts=prior_shorts,
                            prior_note=prior_note, prior_tod=_ptod,
                            prior_tod_sheet=(getattr(a, "prior_tod_sheet", "") or None))
    if data is None:
        c = classify(info.get("error"))
        raise CliError(c[0], info.get("error") or "生成失败", c[1])
    name = "费用BKD_%s.xlsx" % _t.strftime("%Y%m%d_%H%M%S")
    path = _save(data, name)
    # `_最新` 副本的实际文件名：被 Excel 打开时会退到 _最新(2)，必须告诉用户看哪份，
    # 否则又会「改了没生效」——货币资金那边踩过一次。
    _latest = getattr(_save, "latest", "") or ""
    _warn = []
    if not _latest:
        _warn.append("「_最新」副本写入失败（候选名全被占用），请先关掉 Excel 里的底稿再生成。")
    elif not _latest.endswith("BKD_最新.xlsx"):
        _warn.append("「%s」正被 Excel 打开，已另存为 %s；原文件名可能仍是旧内容。"
                     % (_latest.split("_最新")[0] + "_最新.xlsx", _latest))
    _cut = info.get("cutoff") or {}
    if _cut.get("warning"):
        _warn.append(_cut["warning"])
    # AI 待办清单的说明（口径待判定 / 表2 / OAR 分析）：让跑命令行的人也知道
    # "还有几件事没做、清单在哪"，不用去翻 bkd_work 目录
    for _m in (info.get("warnings") or []):
        if _m not in _warn:
            _warn.append(_m)

    # ---- AI 通道：导出待判断清单（平台本身没有模型，由外部 AI 读这份文件）----
    _ai = {}
    if getattr(a, "ai_out", ""):
        import json as _json
        _om = info.get("oarMap") or {}
        try:
            from modules.bkd import expense_oar as _EOAR1
            _reason_items = _EOAR1.ai_reason_items(info.get("oar") or {})
        except Exception as _e:
            _reason_items = []
            _warn.append("OAR 分析清单导出失败：%s" % str(_e)[:100])
        pack = {
            "task": "exp_ai",
            "meta": {"customer": a.customer, "balanceDate": a.date,
                     "period": "%s-%s" % (a.from_, a.to)},
            "oar": {
                "instruction": "上年无对应项的科目，请按业务实质映射到 existingBuckets 里已有口径；"
                               "每项必须给 reason（依据）与 confidence（0~1 或 高/中/低）。",
                "existingBuckets": _om.get("bucketSamples") or {},
                "choices": ["薪酬", "折旧摊销", "广告费", "股权激励", "其他"],
                "items": _om.get("pending") or [],
            },
            "oarReason": {
                "instruction": (
                    "为 OAR 页每个口径写一段变动分析（80~140 字，审计师视角）：\n"
                    "1) 先讲清钱花在哪类业务、由什么推动，用 evidence 里的 key/nature/amt，"
                    "点出金额最大的对手方与其占本口径比重；\n"
                    "2) 讲变动的性质——是普遍性规模增长，还是个别事项（新签合同、一次性支出、"
                    "口径重分类），尽量落到具体事项，不要只说「增加/减少」；\n"
                    "3) 只写证据支持的事实，不得编造未出现的公司、业务或金额；证据不足就直说"
                    "「序时账未见可归类明细，需人工补充」；\n"
                    "4) 不要重复表格已有的金额与比例（那是公式算的），正文提到金额用「千元」整数；\n"
                    "5) 不写「需进一步调查」「建议关注」这类空话；确有关注点就写一句"
                    "「宜确认……」，指明确认什么；\n"
                    "6) 输出 JSON：{\"oarReason\":{\"items\":[{\"label\":\"销售费用\","
                    "\"bucket\":\"广告费\",\"text\":\"……\"}]}}"
                ),
                "items": _reason_items,
            },
            "table2": {
                "instruction": "对金额或变动率超过波动范围的口径，结合 evidence 里的明细构成与"
                               "对手方，写一段审计师视角的变动分析结论（50~120 字，说清"
                               "变动由什么驱动、是否一次性、是否需要追加程序），"
                               "只写判断，不要重复数字——数字由公式给。",
                "items": info.get("table2") or [],
            },
        }
        try:
            with open(a.ai_out, "w", encoding="utf-8") as fh:
                _json.dump(pack, fh, ensure_ascii=False, indent=1)
            _ai = {"file": os.path.abspath(a.ai_out),
                   "oarPending": len(pack["oar"]["items"]),
                   "oarReason": len(pack["oarReason"]["items"]),
                   "table2": len(pack["table2"]["items"]),
                   "next": "把 file 交给平台 AI 判断，再跑 ptw.py ai-fill --file <AI结果.json> "
                           "--xlsx <本底稿>"}
        except Exception as _e:
            _warn.append("AI 清单导出失败：%s" % _e)
    _out = {"ok": True,
            "warnings": _warn,
            "main": {"file": name, "path": path, "latestFile": _latest,
                     "vc": info.get("vc"),
                     "vd": info.get("vd"), "gaps": info.get("gaps"),
                     "amount": info.get("amount"), "period": info.get("period"),
                     "bySegment": info.get("bySegment"),
                     "leadLinks": info.get("leadLinks") or {},
                     "oar": info.get("oar") or {},
                     "cutoff": info.get("cutoff") or {},
                     "priorMatched": info.get("priorMatched"),
                     "priorOnlyAdded": info.get("priorOnlyAdded") or {},
                     "zeroRowsDropped": info.get("zeroRowsDropped") or {},
                     "rateMissing": info.get("rateMissing") or [],
                     "oarMap": info.get("oarMap") or {},
                     # AI 待办：结构化，调用方（含外部 AI）可机器核对"结清没有"
                     "tasks": info.get("tasks") or [],
                     "gapList": info.get("gapList") or []},
            "checks": {"fingerprint": info.get("fingerprint"),
                       "invariants": info.get("invariants"),
                       "baseline": info.get("baseline")},
            "ai": _ai,
            "next": ("K 列（上期审定数）给了 --prior 才会回捞；没给就留空并记进 gapList。"
                     "OAR 页的明细表与 BKD 合计数直接勾稽，不需要另存文件")}

    # ---- 契约字段：让外部 AI 一眼看到「还没做完什么」，不用翻 warnings ----
    # 与货币资金（cmd_cash）保持同一套：tasks / blockingTasks / docHint / version。
    # 只新增字段，不改任何计算与文件；ok 的语义不变（文件已生成），
    # 「能不能交付」单独用 deliverableReady 表达。
    _tasks = info.get("tasks") or []
    _ai_tasks = _exp_ai_tasks(_tasks)
    _pending = [a["id"] for a in _ai_tasks if a["status"] in ("pending", "partial")]
    _block = [t.get("id") for t in _tasks if t.get("blocking")]
    _out["tasks"] = _tasks
    _out["blockingTasks"] = _block
    _out["aiTasks"] = _ai_tasks
    _out["deliverableReady"] = (not _block) and (not _pending)
    _read = ["文档/00_总则.md · 阶段5（结清任务卡）"]
    if _pending:
        _read.insert(0, "文档/10_费用.md · AI 三件事（%d 项未写：%s）"
                     % (len(_pending), "、".join(_pending)))
    if (info.get("rateMissing") or []):
        _read.append("文档/10_费用.md · 折算汇率（%d 个主体缺汇率）"
                     % len(info["rateMissing"]))
    _out["docHint"] = {"entry": "SKILL.md", "read": _read,
                       "verify": "main.aiTasks 全部 done 且 deliverableReady=true"}
    _out["version"] = _platform_version()
    # main 里也放一份：老调用方只看 main 时同样读得到
    _out["main"]["aiTasks"] = _ai_tasks
    _out["main"]["blockingTasks"] = _block
    _out["main"]["deliverableReady"] = _out["deliverableReady"]
    return _out


def cmd_asset(a):
    """生成固定资产 SWP 底稿（长期资产）。

    V1.2.98 起这条命令跑的是 SWP 引擎（``modules/fa_swp``），不再是旧的
    K1_100Breakdown 链路：数据只从科目余额表 + 序时账来，不再要求客户另给
    FA list 或分类对照表。
    """
    import datetime as _dt

    lg = boot(a.data_dir)
    from modules.fa_swp import build as FB

    ds = (a.date or "").strip()
    if not ds:
        raise CliError("BAD_DATE", "必须给 --date（资产负债表日，如 2026-09-30）",
                       "如 `ptw.py asset --date 2026-09-30`")
    try:
        bd = _dt.date.fromisoformat(ds)
    except Exception:                                       # noqa: BLE001
        raise CliError("BAD_DATE", "资产负债表日格式不对：%r" % ds,
                       "用 YYYY-MM-DD，例如 2026-09-30")

    warn = []
    if a.map or a.from_ or a.to:
        warn.append("--map / --from / --to 已随旧 K1 链路下线，本次忽略："
                    "SWP 从科目余额表 + 序时账取数，不需要资产台账或分类对照表。")

    data_dir = os.environ.get("KZ_DATA_DIR") or ""
    info = {"customer": a.customer or "", "balance_date": ds,
            "te": 0, "sad": 0,
            "standard": "中国企业会计准则", "currency": "人民币元"}
    try:
        res = FB.build(lg.conn, data_dir, os.path.join(data_dir, "bkd_out"),
                       info, bd, "ptw")
    except FileNotFoundError as e:
        raise CliError("TPL_MISSING", str(e),
                       "把 K1 SWP 模板放回 data/bkd_tpl/ 再试")
    except ValueError as e:
        raise CliError("BUILD_FAILED", str(e),
                       "数据或模板不满足条件，看控制台回溯定位")

    tr = res["trace"]
    return {"ok": True,
            "warnings": warn,
            "main": {"file": os.path.basename(res["path"]),
                     "path": res["path"],
                     "latestFile": os.path.basename(res["latest"]),
                     "priorFile": os.path.basename(res["prior"]),
                     "cats": tr["cat_order"],
                     "totals": tr["totals"], "net": tr["net"],
                     "k01": tr["k01"], "table4": tr["table4"],
                     "rowsInserted": tr["rows_inserted"],
                     "tbDup": tr["tb_dup"], "checkItems": tr["check_items"]},
            "next": ("K.01 的 TB 核对块与 Lead 互校必须全 0；非零要解释，不许调平。"
                     "上年 K.01、上年 Lead、TB 都是隐藏 sheet，取消隐藏可查来源。")}


def cmd_check(a):
    lg = boot(a.data_dir)
    r = lg.recon(period_from=a.from_, period_to=a.to, tol=a.tol)
    summ = (r or {}).get("summary") or {}
    bad = [i for i in ((r or {}).get("items") or []) if not i.get("ok")]
    return {"ok": True,
            "summary": summ,
            "failed": [{"no": i.get("no"), "title": i.get("title"),
                        "bad": i.get("bad"), "error": i.get("error")} for i in bad],
            "next": ("全部通过" if not bad else
                     "有 %d 项不通过：先看 error 字段，是「数据没准备好」还是「真的不平」" % len(bad))}


def cmd_importpack(a):
    """导入标准表数据包（kanzhang/v1）。客户资料已在别处清洗成标准表时走这条路，
    不必再绕回 xlsx。"""
    lg = boot(a.data_dir)
    try:
        r = lg.load_pack(a.file, replace=not a.keep)
    except Exception as e:
        raise CliError("PACK_BAD", "数据包导入失败：%s" % e,
                       "确认是 kanzhang/v1 数据包（zip 内有 tb.parquet / journal.parquet）")
    tabs = {d["table"]: d["rows"] for d in r["tables"]}
    if not tabs:
        raise CliError("PACK_EMPTY", "包里没有可识别的表（要 tb.parquet 或 journal.parquet）",
                       "检查 zip 结构")
    return {"ok": True, "pack": r, "tables": tabs,
            "next": "接下来直接 ptw.py cash --date ... 或 exp --from 1 --to 9"}


# ---------------------------------------------------------------- 指纹 / 基准 / 配置包
def cmd_oarmap_draft(a):
    """导出口径 mapping 里「本年新增、上年找不到」的科目，交给 AI 判断。

    为什么要有这一步：口径是**判断**，不能全靠关键词表。关键词表认不出的科目
    （新客户常有的「人力成本_年终奖」「库宝软件」）以前直接落成科目名，
    OAR 汇总就散了。现在统一落到「待判断」，由 AI 起草、人工确认，确认后
    写进 oar_map_ai.json，明年这家客户再来就自动沿用。
    """
    lg = boot(a.data_dir)
    from modules.bkd import expense as EXP
    from modules.bkd import oar_map as OM

    rows, rmeta = EXP.read_expense(lg.conn, a.from_, a.to)
    if rmeta.get("error") or not rows:
        raise CliError("NO_DATA", rmeta.get("error") or "序时账里没有销售费用/管理费用行",
                       "先导入并合并序时账")

    prior_oar, prior_shorts, note = {}, {}, ""
    if a.prior:
        if not os.path.isfile(a.prior):
            raise CliError("FILE_NOT_FOUND", "上年费用底稿不存在：%s" % a.prior, "检查 --prior")
        with open(a.prior, "rb") as fh:
            prior_oar, prior_shorts, note = OM.read_prior(fh.read())

    book_map = EXP.read_book_map(lg.conn)
    map_rows = []
    for r in rows:
        comp, _how = EXP.resolve_company(r["book"], book_map)
        map_rows.append({"company": comp, "name": r["name"], "amount": r["amount"]})

    from openpyxl import Workbook
    info = OM.build(Workbook(), prior_oar, prior_shorts, map_rows,
                    OM.load_ai_map(OM.ai_map_path()), book_map=book_map)

    out = a.out or "oar_map_pending.json"
    OM.export_pending(info, {"period": rmeta.get("period"),
                             "prior": os.path.basename(a.prior or ""),
                             "note": note}, out)
    return {"ok": True, "pending": info.get("pendingCount", 0),
            "file": os.path.abspath(out),
            "oarRows": info.get("oarRows"), "companyRows": info.get("companyRows"),
            "next": "把 file 交给 AI 按 03_文档/AI提示词_口径判断.md 判断，写回 items[].bucket，"
                    "再跑 ptw.py oarmap-fill --file <AI结果.json>"}


def cmd_oarmap_fill(a):
    """把 AI/人工确认的口径写进 oar_map_ai.json，下次生成自动生效。"""
    if APP not in sys.path:
        sys.path.insert(0, APP)
    # 这个命令不开库，但要把结果写进当前数据目录，自己设一下 env（不设就落到真实 data/）
    if getattr(a, "data_dir", ""):
        os.environ["KZ_DATA_DIR"] = os.path.abspath(a.data_dir)
    from modules.bkd import oar_map as OM
    if not os.path.isfile(a.file):
        raise CliError("FILE_NOT_FOUND", "AI 结果文件不存在：%s" % a.file, "检查 --file")
    items = OM.read_items(a.file)
    if not items:
        raise CliError("EMPTY", "没读到任何 {company,name,bucket} 三件套",
                       "格式见 oarmap-draft 导出的 JSON，bucket 只能是 "
                       "薪酬/折旧摊销/广告费/股权激励/其他 或单列口径名")
    r = OM.save_ai_map(items)
    return {"ok": True, **r,
            "next": "重新跑 ptw.py exp ... 生成，BKD 的 OAR 口径列会 XLOOKUP 到这里"}


def cmd_oar_fill(a):
    """把 AI 写的 OAR 结论回填进已生成的费用底稿。

    为什么单独一个命令：结论是「判断」，不能让程序套模板生成。
    流程是 —— 生成时导出 evidence JSON → AI 读它写结论 →
    本命令回填。只换文字，不动公式，所以金额/比例仍然是活的引用。
    """
    import json as _json
    # 这个命令只动文件、不碰数据库，但 modules 包要靠 sys.path 才能 import ——
    # 其它命令都是 boot() 里设的，这里得自己补一句（不能 boot，会白开一次库）。
    if APP not in sys.path:
        sys.path.insert(0, APP)
    from modules.bkd import expense_oar as EOAR

    if not os.path.isfile(a.file):
        raise CliError("FILE_NOT_FOUND", "底稿不存在：%s" % a.file, "检查 --file 路径")
    if not os.path.isfile(a.conclusions):
        raise CliError("FILE_NOT_FOUND", "结论文件不存在：%s" % a.conclusions,
                       "先看 evidence JSON，写成 {'科目名称': '结论文字'} 存盘")
    try:
        with open(a.conclusions, encoding="utf-8") as fh:
            concl = _json.load(fh)
    except Exception as e:
        raise CliError("BAD_JSON", "结论文件读不出来：%s" % e, "要 UTF-8 的 JSON 对象")
    if not isinstance(concl, dict) or not concl:
        raise CliError("BAD_JSON", "结论文件要是一个非空 JSON 对象",
                       "格式：{'科目名称': '结论文字'}")

    with open(a.file, "rb") as fh:
        data = fh.read()
    new_data, info = EOAR.apply_conclusions(data, concl)
    if not info.get("applied"):
        raise CliError("OAR_NO_MATCH", "没有匹配上的科目（%s）"
                       % "、".join(info.get("unmatched") or [])[:120],
                       "科目名要与 OAR 页 C 列一字不差；先跑 exp --oar-evidence 看账户清单")

    out = a.file if a.inplace else os.path.join(
        os.path.dirname(a.file), os.path.basename(a.file).replace(".xlsx", "_已补结论.xlsx"))
    with open(out, "wb") as fh:
        fh.write(new_data)
    # 落盘，下次重新生成自动带上（否则改个 TE 就被默认描述顶掉）
    try:
        saved = EOAR.save_conclusions(concl)
    except Exception as e:
        saved = {"error": "%s: %s" % (type(e).__name__, str(e)[:120])}
    return {"ok": True, "file": out, "applied": info.get("applied"),
            "accounts": info.get("accounts"),
            "unmatched": info.get("unmatched"),
            "saved": saved,
            "next": "打开底稿看 OAR 页：结论已写入且底色转白；金额/比例仍是公式"}


def cmd_ai_fill(a):
    """把平台 AI 的判断结果回填：OAR 口径进配置，BKD 表2 的文字进底稿。

    为什么必须分两步：PTW 平台自己**没有模型**。命令行负责「把该判断的东西摊平
    成 JSON」，平台 AI（或人）负责判断，本命令负责把判断写回该在的位置：
      · oar.items  → oar_map_ai.json（下次生成时 mapping 的「来源」列写「由AI判断」）
      · table2     → 直接改底稿 xlsx 里表2 每行公式的最后一段文字，金额仍走公式
    """
    import json as _json
    if APP not in sys.path:
        sys.path.insert(0, APP)
    if getattr(a, "data_dir", ""):
        os.environ["KZ_DATA_DIR"] = os.path.abspath(a.data_dir)
    if not os.path.isfile(a.file):
        raise CliError("FILE_NOT_FOUND", "AI 结果文件不存在：%s" % a.file, "检查 --file")
    try:
        with open(a.file, encoding="utf-8") as fh:
            d = _json.load(fh)
    except Exception as e:
        raise CliError("BAD_JSON", "AI 结果读不出来：%s" % e, "要 UTF-8 JSON")

    from modules.bkd import oar_map as OM
    from modules.bkd import expense_oar as EOAR
    from modules.bkd import expense as EXP

    oar = d.get("oar") if isinstance(d, dict) else None
    oar_items = (oar or {}).get("items") if isinstance(oar, dict) else None
    oar_items = oar_items or d.get("oarItems") or []
    # OAR 页「二、变动分析」的文字（与口径映射是两件事：一个是科目归哪类，
    # 一个是这一类为什么变）。键名 oarReason，也认 oar.reason / oar.reasons。
    _or = d.get("oarReason") if isinstance(d, dict) else None
    oar_reason = (_or.get("items") if isinstance(_or, dict) else _or) or []
    if not oar_reason and isinstance(oar, dict):
        oar_reason = oar.get("reasons") or oar.get("reason") or []
    t2 = d.get("table2") if isinstance(d, dict) else None
    t2_items = (t2 or {}).get("items") if isinstance(t2, dict) else None
    t2_items = t2_items or d.get("table2") or []

    out = {"ok": True, "oar": None, "table2": None}
    if oar_items:
        out["oar"] = OM.save_ai_map(oar_items)

    if oar_reason:
        # 先落盘：以后每次重新生成都会自动带上，不会被默认描述顶掉
        out["oarReasonSaved"] = EOAR.save_conclusions(oar_reason)
        if not a.xlsx:
            out["oarReason"] = {"error": "没给 --xlsx，OAR 分析只写进了配置，"
                                         "下次生成会自动带上"}
        elif not os.path.isfile(a.xlsx):
            raise CliError("FILE_NOT_FOUND", "底稿不存在：%s" % a.xlsx, "检查 --xlsx")
        else:
            with open(a.xlsx, "rb") as fh:
                raw = fh.read()
            new, info = EOAR.apply_conclusions(raw, EOAR.load_conclusions())
            dst = a.xlsx if a.inplace else os.path.join(
                os.path.dirname(a.xlsx),
                os.path.basename(a.xlsx).replace(".xlsx", "_已补结论.xlsx"))
            with open(dst, "wb") as fh:
                fh.write(new)
            out["oarReason"] = dict(info, file=dst)

    if t2_items:
        # 先落盘：以后每次重新生成都会自动带上这些结论，不会被默认描述顶掉
        out["table2Saved"] = EXP.save_table2_concl(t2_items)
        if not a.xlsx:
            out["table2"] = {"error": "没给 --xlsx，表2 结论只写进了配置，"
                                      "下次生成会自动带上"}
        elif not os.path.isfile(a.xlsx):
            raise CliError("FILE_NOT_FOUND", "底稿不存在：%s" % a.xlsx, "检查 --xlsx")
        else:
            with open(a.xlsx, "rb") as fh:
                raw = fh.read()
            new, info = EXP.apply_table2(raw, {"items": t2_items})
            dst = a.xlsx if a.inplace else os.path.join(
                os.path.dirname(a.xlsx),
                os.path.basename(a.xlsx).replace(".xlsx", "_已补结论.xlsx"))
            with open(dst, "wb") as fh:
                fh.write(new)
            out["table2"] = dict(info, file=dst)

    if not oar_items and not t2_items and not oar_reason:
        raise CliError("EMPTY", "既没有 oar.items、oarReason 也没有 table2，无从回填",
                       "格式见 exp --ai-out 导出的 JSON")
    out["next"] = ("重新跑一次 exp 生成，mapping 的「来源」列会写「由AI判断」、"
                   "备注列带上依据与置信度；表2 已直接改在 xlsx 里")
    return out


def cmd_fingerprint(a):
    lg = boot(a.data_dir)
    from core import fingerprint as FP
    kind = a.subject or "cash"
    p = FP.path_for(lg.data_dir, kind)
    fp = FP.load(p)
    if not fp:
        raise CliError("FP_MISSING", "还没有指纹：%s" % p, "先跑一次 cash 或 exp")
    return {"ok": True, "subject": kind, "path": p,
            "params": fp.get("params"),
            "sheets": [{"name": s.get("name"), "headerRow": s.get("headerRow"),
                        "rows": s.get("rows"), "cols": s.get("cols")}
                       for s in fp.get("sheets") or []],
            "links": fp.get("links"),
            "next": "指纹只记结构（表名/表头/行列数/链接），数据变化不影响它"}


def cmd_baseline(a):
    lg = boot(a.data_dir)
    from core import baseline as BL
    kind = a.subject or "cash"
    p = BL.path_for(lg.data_dir, kind)
    base = BL.load(p)
    if not base:
        raise CliError("BASE_MISSING", "还没有基准：%s" % p, "先跑一次生成，会自动建立")
    from core import invariants as INV
    return {"ok": True, "subject": kind, "path": p,
            "capturedAt": base.get("capturedAt"),
            "values": base.get("values"),
            "next": "下次生成会自动和这份比；连续两次同参数跑，第二次应报『完全一致』"}


def cmd_project(a):
    lg = boot(a.data_dir)
    from core import project as PJ
    act = a.action or "show"
    if act == "show":
        cfg = PJ.load_config(lg.data_dir)
        return {"ok": True, "path": PJ.config_path(lg.data_dir), "config": cfg}
    if act == "export":
        cust = a.customer or ""
        nm = "客户配置_%s_%s.ptwpack" % (cust or "未命名", a.period or "noPeriod")
        out = a.out or os.path.join(lg.data_dir, "exports", nm)
        r = PJ.export_pack(lg, lg.data_dir, out, customer=cust, period=a.period or "")
        return {"ok": True, "file": r["path"], "size": r["size"], "meta": r["meta"],
                "next": "包里含结构配置、配方、已确认判断；判断导入后会变成待复核"}
    if act == "import":
        if not a.file:
            raise CliError("NO_FILE", "要指定 --file 配置包路径",
                           "ptw.py project import --file 客户配置_xx.ptwpack")
        r = PJ.import_pack(lg, lg.data_dir, a.file)
        return {"ok": True, **r}
    raise CliError("BAD_ACTION", "未知动作：%s" % act, "可用：show / export / import")


# ---------------------------------------------------------------- main
def build_parser():
    p = argparse.ArgumentParser(
        prog="ptw",
        description="PTW 命令行入口（不起服务，直接调 core）")
    p.add_argument("--data-dir", default=None,
                   help="数据目录，默认 程序/../data")
    p.add_argument("--pretty", action="store_true", help="格式化 JSON 输出")
    sub = p.add_subparsers(dest="cmd")

    sub.add_parser("state", help="看数据状态：哪些表已合并、能不能直接生成")

    pi = sub.add_parser("import", help="导入 Excel/CSV 并自动映射、清洗、合并")
    pi.add_argument("file", help="要导入的文件路径")
    pi.add_argument("--target", choices=["tb", "journal", "flow", "fa", "fa_beg"], default=None,
                    help="强制指定数据类别；不给则按表头自动判定")

    pc = sub.add_parser("cash", help="生成货币资金 C.00 BKD + Lead + Cutoff + 检查表")
    pc.add_argument("--date", default="", help="资产负债表日，如 2026-10-31")
    pc.add_argument("--customer", default="", help="客户名称，写进 Lead")
    pc.add_argument("--te", type=float, default=None,
                    help="可容忍误差 TE（元），写进 Lead C4；不传则读共用基础信息")
    pc.add_argument("--sad", type=float, default=None, help="名义金额 SAD（元），默认 TE×10%%")
    pc.add_argument("--standard", default="", help="适用会计准则，写进 Lead C6")
    pc.add_argument("--currency", default="", help="记账本位币，写进 Lead C7")
    pc.add_argument("--prior", default="", help="上年底稿路径（可选，BKD 明细）")
    pc.add_argument("--prior-lead", default="",
                    help="上年 Lead 路径；默认与 --prior 同一个文件（上年 Lead 就在上年底稿里）")
    pc.add_argument("--rate", action="append", metavar="主体=汇率",
                    help="折算率，可重复，如 --rate 甲公司=7.1")
    pc.add_argument("--no-cutoff", action="store_true", help="不做截止性测试")
    pc.add_argument("--cutoff-before", type=int, default=1)
    pc.add_argument("--cutoff-after", type=int, default=1)
    pc.add_argument("--cutoff-amount", type=float, default=100000)
    pc.add_argument("--strict", action="store_true",
                    help="模板结构与上次不一致时不落盘（只报差异）")
    pc.add_argument("--restricted", action="append", default=[],
                    help="受限资金 账号=金额，可重复（人工判断结果，工具不猜）")
    # TE 缺失默认报错（2026-11-12）：静默按 10 万算会让整份底稿的"是否进一步调查"
    # 判断集体失准且看不出来。确实只能用默认值时显式放行，并在交付说明里写明。
    pc.add_argument("--allow-default-te", action="store_true",
                    help="TE 读不到时允许按默认值继续（默认报错；用了要在交付说明里写明）")

    pe = sub.add_parser("exp", help="生成销管费用 VC/VD 底稿")
    pe.add_argument("--from", dest="from_", default="1", help="期间起，1~12 或 2026-01")
    pe.add_argument("--to", dest="to", default="6", help="期间止")
    pe.add_argument("--date", default="")
    pe.add_argument("--customer", default="")
    pe.add_argument("--prior", default="", help="上年费用底稿路径（用于回捞上期审定数）")
    pe.add_argument("--prior-tod", default="",
                    help="上一期 TOD 底稿路径：整页搬入本年底稿，「预审+剩余期间」的"
                         "预审段改为引用它（两期重要性水平不同，不重算）")
    pe.add_argument("--prior-tod-sheet", default="",
                    help="上一期 TOD 用哪张工作表（页签不规范时手动指定；留空自动识别）")
    pe.add_argument("--ai-out", default="",
                    help="把 OAR 待判断科目 + BKD 表2 待分析口径导出成一份 JSON，"
                         "交给平台 AI 写结论；填完用 ptw.py ai-fill 回填")
    pe.add_argument("--pre-end", default=None,
                    help="预审段结束月份（如 6）；只用于分段金额，不影响 BKD 输出")
    pe.add_argument("--oar-evidence", default="",
                    help="OAR 变动证据导出路径（JSON），交给 AI 写结论用")
    pe.add_argument("--oar-chg-min", type=float, default=None,
                    help="OAR 分析对象的最小变动金额（默认取 SAD）")
    pe.add_argument("--oar-chg-pct", type=float, default=0.10,
                    help="OAR 分析对象的变动率门槛，默认 0.10（10%%）")
    pe.add_argument("--te", type=float, default=None,
                    help="可容忍误差 TE（元）。写进 Lead C4，波动幅度阈值=C22=TE×50%%")
    pe.add_argument("--sad", type=float, default=None,
                    help="名义金额 SAD（元），默认 TE×10%%")
    pe.add_argument("--standard", default="",
                    help="适用会计准则，写进 Lead C6（如：中国企业会计准则）")
    pe.add_argument("--currency", default="",
                    help="记账本位币，写进 Lead C7（如：人民币 / CNY）")
    pe.add_argument("--assurance", default="",
                    help="四项认定等级，逗号分隔（完整性,发生,计量,列报披露）；"
                         "可选 Minimal/Low/Moderate/High，默认全 Minimal")

    pfill = sub.add_parser("oar-fill", help="把 AI 写的 OAR 结论回填进费用底稿")
    pfill.add_argument("--file", required=True, help="已生成的费用BKD xlsx")
    pfill.add_argument("--conclusions", required=True, help="结论 JSON：{科目名称: 结论文字}")
    pfill.add_argument("--inplace", action="store_true", help="直接覆盖原文件（默认另存 _已补结论）")

    pmap1 = sub.add_parser("oarmap-draft", help="导出口径 mapping 里待 AI 判断的科目")
    pmap1.add_argument("--from", dest="from_", default="1")
    pmap1.add_argument("--to", dest="to", default="12")
    pmap1.add_argument("--prior", default="", help="上年费用底稿路径（有它才能沿用上年口径）")
    pmap1.add_argument("--out", default="", help="输出 JSON 路径，默认 oar_map_pending.json")

    pmap2 = sub.add_parser("oarmap-fill", help="把 AI 判断的口径写进 oar_map_ai.json")
    pmap2.add_argument("--file", required=True, help="AI 返回的 JSON（含 items[].bucket）")

    paf = sub.add_parser("ai-fill", help="回填平台 AI 的判断：OAR 口径 + BKD 表2 结论")
    paf.add_argument("--file", required=True, help="AI 结果 JSON（exp --ai-out 导出的那份，填完后）")
    paf.add_argument("--xlsx", default="", help="要回填表2 的底稿 xlsx（不给就只写 OAR 口径配置）")
    paf.add_argument("--inplace", action="store_true", help="直接覆盖原底稿（默认另存 _已补结论）")

    pa = sub.add_parser("asset", help="生成固定资产 SWP 底稿（K.01 + Lead + 三张清单）")
    pa.add_argument("--map", default="", help="【已下线】旧版分类对照表，传入会被忽略")
    pa.add_argument("--date", default="", help="资产负债表日，如 2026-09-30")
    pa.add_argument("--customer", default="")
    pa.add_argument("--from", dest="from_", default="", help="【已下线】期间起，传入会被忽略")
    pa.add_argument("--to", dest="to", default="", help="【已下线】期间止，传入会被忽略")

    pf = sub.add_parser("fingerprint", help="查看模板指纹（结构漂移检测）")
    pf.add_argument("--subject", choices=["cash", "exp"], default="cash")

    pb = sub.add_parser("baseline", help="查看回归基准")
    pb.add_argument("--subject", choices=["cash", "exp"], default="cash")

    pp = sub.add_parser("project", help="项目配置：show / export / import 客户配置包")
    pp.add_argument("action", nargs="?", default="show",
                    choices=["show", "export", "import"])
    pp.add_argument("--file", default="", help="导入时的配置包路径")
    pp.add_argument("--out", default="", help="导出时的输出路径")
    pp.add_argument("--customer", default="")
    pp.add_argument("--period", default="")

    ppk = sub.add_parser("importpack", help="导入标准表数据包（kanzhang/v1 zip）")
    ppk.add_argument("--file", required=True, help="数据包 zip 路径")
    ppk.add_argument("--keep", action="store_true",
                     help="追加到现有数据（默认重建表，即替换）")

    pk = sub.add_parser("check", help="跑五组数据检查")
    pk.add_argument("--from", dest="from_", default=None)
    pk.add_argument("--to", dest="to", default=None)
    pk.add_argument("--tol", type=float, default=0.01)
    return p


def main():
    a = build_parser().parse_args()
    if not a.cmd:
        build_parser().print_help()
        return 2
    fn = {"state": cmd_state, "import": cmd_import, "cash": cmd_cash,
          "exp": cmd_exp, "oar-fill": cmd_oar_fill, "check": cmd_check,
          "oarmap-draft": cmd_oarmap_draft, "oarmap-fill": cmd_oarmap_fill,
          "ai-fill": cmd_ai_fill,
          "asset": cmd_asset,
          "fingerprint": cmd_fingerprint, "baseline": cmd_baseline,
          "project": cmd_project, "importpack": cmd_importpack}.get(a.cmd)
    try:
        emit(fn(a), a.pretty)
        return 0
    except CliError as e:
        emit({"ok": False, "code": e.code, "error": e.msg, "next": e.next}, a.pretty)
        return 1
    except Exception as e:
        import traceback
        traceback.print_exc(file=sys.stderr)
        code, nxt = classify(e)
        emit({"ok": False, "code": code, "error": str(e), "next": nxt}, a.pretty)
        return 1


if __name__ == "__main__":
    sys.exit(main())
