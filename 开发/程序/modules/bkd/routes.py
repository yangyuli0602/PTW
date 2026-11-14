# -*- coding: utf-8 -*-
"""底稿生成（货币资金 BKD）板块 · 路由处理

本板块负责：自动读取PTW平台清洗后的科目余额表、接收上年底稿上传、
生成 BKD 底稿与检查表。核心逻辑在 core.py（与统一 server 解耦）。

路由沿用原 BKD 独立服务的接口路径，保证前端交互不变。
"""
import os
import io
import re
import json
import time
import urllib.parse
from datetime import date as _date, datetime as _dt, timedelta

from . import core as BKD

GET_ROUTES = {}
POST_ROUTES = {}


def _get(path):
    def deco(fn):
        GET_ROUTES[path] = fn
        return fn
    return deco


def _post(path):
    def deco(fn):
        POST_ROUTES[path] = fn
        return fn
    return deco


# ---------------------------------------------------------------- 工具
def _parse_multipart(raw, ctype):
    """简易 multipart 解析：提取文件字节 + filename + 文本字段。"""
    m = re.search(r"boundary=([^;]+)", ctype)
    if not m:
        raise ValueError("无 boundary")
    boundary = m.group(1).strip().strip('"').encode()
    parts = raw.split(b"--" + boundary)
    files, fields = {}, {}
    for part in parts:
        if b"\r\n\r\n" not in part:
            continue
        head, _, data = part.partition(b"\r\n\r\n")
        data = data.rstrip(b"\r\n")
        head_s = head.decode("utf-8", "ignore")
        fn = re.search(r'name="([^"]+)"', head_s)
        if not fn:
            continue
        name = fn.group(1)
        ffn = re.search(r'filename="([^"]*)"', head_s)
        if ffn:
            files[name] = {"filename": ffn.group(1), "data": data}
        else:
            fields[name] = data.decode("utf-8", "ignore")
    return {"files": files, "fields": fields}


# ---------------------------------------------------------------- GET
@_get("/api/bkd/health")
def api_health(ctx, q):
    return ctx.json({"ok": True, "db": BKD.KZ_DB, "db_exists": os.path.exists(BKD.KZ_DB)})


@_get("/api/tb")
def api_tb(ctx, q):
    kw = (q.get("kw") or ["银行存款"])[0]
    conn = getattr(ctx.ledger, "conn", None)
    records, meta = BKD.read_tb_from_duckdb(kw, conn=conn)
    if meta.get("error"):
        return ctx.json({"error": meta["error"], "records": []})
    return ctx.json({"records": records, "meta": meta})


@_get("/api/tpl")
def api_tpl(ctx, q):
    return _tpl_info(ctx)


@_get("/api/journal-probe")
def api_journal_probe(ctx, q):
    """序时账自检：返回行数、日期范围、科目分布、货币资金识别情况。"""
    conn = getattr(ctx.ledger, "conn", None)
    try:
        return ctx.json(BKD.probe_journal(conn=conn))
    except Exception as e:
        return ctx.json({"error": "读取序时账失败：" + str(e)})


@_get("/api/tpl/download")
def api_tpl_download(ctx, q):
    return _tpl_download(ctx)


@_get("/api/book/plan")
def api_book_plan(ctx, q):
    """账面信息表：先看能生成多少行、哪些行缺银行名/账号。"""
    kw = (q.get("kw") or ["银行存款"])[0]
    conn = getattr(ctx.ledger, "conn", None)
    try:
        rows, meta = BKD.read_book_rows(conn=conn, keyword=kw)
    except Exception as e:
        return ctx.json({"error": "读取科目余额表失败：" + str(e), "rows": []})
    if meta.get("error"):
        return ctx.json({"error": meta["error"], "rows": []})
    return ctx.json({"rows": rows, "meta": meta,
                     "tpl": os.path.basename(BKD.book_tpl_path() or ""),
                     "tplExists": bool(BKD.book_tpl_path())})


@_get("/api/book/tpl/download")
def api_book_tpl_download(ctx, q):
    tp = BKD.book_tpl_path()
    if not tp:
        return ctx.err("未找到账面信息模板", 404)
    return ctx.send_file(tp,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        os.path.basename(tp))


# ---------------------------------------------------------------- POST
@_post("/api/book/generate")
def api_book_generate(ctx, u, body):
    return _book_generate(ctx, body or {})


@_post("/api/book/tpl/upload")
def api_book_tpl_upload(ctx, u, body):
    return _book_tpl_upload(ctx, ctx.multipart())


@_post("/api/prior/plan")
def api_prior_plan(ctx, u, body):
    body = ctx.multipart()
    return _prior_plan(ctx, body)


@_post("/api/generate")
def api_generate(ctx, u, body):
    body = ctx.multipart()
    return _generate(ctx, body)


@_post("/api/tpl/upload")
def api_tpl_upload(ctx, u, body):
    body = ctx.multipart()
    return _tpl_upload(ctx, body)


# ------------------------------------------------ 上数（无上年底稿也能跑）
@_get("/api/build/c00/plan")
def api_build_c00_plan(ctx, q):
    """银行存款 BKD 预览：能写多少行、有多少格待人工补、缺什么、有哪些主体要填折算率。"""
    conn = getattr(ctx.ledger, "conn", None)
    try:
        _data, info = _run_c00(conn, preview=True)
    except Exception as e:
        return ctx.json({"error": "预览失败：" + str(e)})
    if info.get("error"):
        return ctx.json({"error": info["error"]})
    return ctx.json(info)


def _cutoff_fields_of(body):
    """把前端传来的截止测试设置拍平成 _build_cutoff 认识的字段。

    接受两种写法：嵌套 {"cutoff": {"enabled":true,...}} 与扁平 cutoff_date 等，
    兼容老的 /api/generate 表单口径。
    """
    cf = (body or {}).get("cutoff")
    if isinstance(cf, dict):
        if cf.get("enabled") is False or str(cf.get("enabled")) == "0":
            return None
        d = cf.get("date") or ""
        if not d:
            return None
        return {"cutoff": "1", "cutoff_date": d,
                "cutoff_before": cf.get("before", 1),
                "cutoff_after": cf.get("after", 1),
                "cutoff_threshold": cf.get("threshold", 100000)}
    if isinstance(cf, str) and cf in ("0", ""):
        return None
    if (body or {}).get("cutoff_date"):
        return {"cutoff": str(body.get("cutoff", "1")),
                "cutoff_date": body.get("cutoff_date"),
                "cutoff_before": body.get("cutoff_before", 1),
                "cutoff_after": body.get("cutoff_after", 1),
                "cutoff_threshold": body.get("cutoff_threshold", 100000)}
    return None


@_get("/api/build/baseinfo")
def api_baseinfo_get(ctx, q):
    """共用基础信息：客户名称 / 资产负债表日 / 准则 / 本位币 / TE / SAD。

    全底稿共用一份，任何一份 Lead 生成时都读它写封面。
    """
    try:
        from . import baseinfo as BI
    except ImportError:
        import baseinfo as BI
    return ctx.json({"fields": [{"key": k, "label": l, "hint": h}
                                for k, l, h in BI.FIELDS],
                     "info": BI.load(BKD.KZ_DATA),
                     "cra": {m: BI.cra_of(BI.load(BKD.KZ_DATA), m)
                             for m in BI.CRA_MODULES},
                     "craTypes": list(BI.CRA_TYPES),
                     "craDefaults": {m: BI.default_pcts(m) for m in BI.CRA_MODULES},
                     "assertions": [{"key": k, "label": l} for k, _w, l in BI.ASSERTIONS],
                     # 各模块的认定项不同（费用是 完整/发生/计量/列报），前端按模块切换
                     "assertionsByModule": {
                         m: [{"key": k, "label": l} for k, _w, l in BI.assertions_of(m)]
                         for m in BI.CRA_MODULES},
                     "modules": list(BI.CRA_MODULES),
                     "file": BI.path_of(BKD.KZ_DATA)})


@_get("/api/build/baseinfo/export")
def api_baseinfo_export(ctx, q):
    """导出关键信息（含 CRA 配置）为一段 JSON 文本，前端存成文件。

    为什么要能导出：新项目开张时，客户、准则、本位币、TE、CRA 那一套往往和上个项目
    只差几个字。能导出→改几处→导入，比每次重填六个格 + 五档百分比靠谱。
    """
    try:
        from . import baseinfo as BI
    except ImportError:
        import baseinfo as BI
    info = BI.load(BKD.KZ_DATA)
    payload = {k: info.get(k) for k in ("customer", "balance_date", "standard",
                                        "currency", "te", "sad", "cra")}
    payload["_exportedAt"] = time.strftime("%Y-%m-%d %H:%M:%S")
    payload["_app"] = "PTW"
    return ctx.json({"ok": True, "filename": "PTW关键信息_%s.json" % time.strftime("%Y%m%d"),
                     "text": json.dumps(payload, ensure_ascii=False, indent=1)})


@_post("/api/build/baseinfo/import")
def api_baseinfo_import(ctx, u, body):
    """导入关键信息：接受 {"info": {...}} 或 {"text": "<json 字符串>"}。"""
    try:
        from . import baseinfo as BI
    except ImportError:
        import baseinfo as BI
    body = body or {}
    data = body.get("info")
    if not isinstance(data, dict):
        raw = body.get("text") or ""
        try:
            data = json.loads(raw) if raw else {}
        except Exception as e:
            return ctx.err("导入失败：不是合法的 JSON（%s）" % e)
    if not isinstance(data, dict) or not data:
        return ctx.err("导入失败：文件里没有可识别的关键信息")
    keep = {k: v for k, v in data.items() if k in BI.DEFAULTS}
    keep.pop("updatedAt", None)
    try:
        info = BI.save(BKD.KZ_DATA, keep)
    except Exception as e:
        return ctx.err("导入失败：%s" % e)
    return ctx.json({"ok": True, "info": info,
                     "imported": sorted(k for k in keep if k != "cra")})


@_post("/api/build/baseinfo/save")
def api_baseinfo_save(ctx, u, body):
    try:
        from . import baseinfo as BI
    except ImportError:
        import baseinfo as BI
    try:
        info = BI.save(BKD.KZ_DATA, body or {})
    except Exception as e:
        return ctx.err("保存基础信息失败：%s" % e)
    return ctx.json({"ok": True, "info": info})


@_post("/api/build/c00/run")
def api_build_c00_run(ctx, u, body):
    body = body or {}
    conn = getattr(ctx.ledger, "conn", None)
    prior, pinfo = _load_prior(body.get("prior_fid"), body.get("prior_spec"))
    _pl = _load_prior_lead(body.get("prior_fid"),
                           (body.get("prior_spec") or {}).get("lead"))   # 上年 Lead 就在上年底稿里，同一个文件
    # 选了「上年 Lead」却没读到：必须让用户在结果页看见，否则表1 上期数悄悄变 0。
    # 这个入口没有 warnings 列表，先记下来，等 info 出来再挂进去。
    _pl_err = ("上年 Lead 未使用（表1 上期数、表2 上年数会是 0）：" + _pl["error"]
               if _pl and _pl.get("error") else "")
    try:
        data, info = _run_c00(conn, prior=prior,
                              balance_date=body.get("balance_date", ""),
                              customer=body.get("customer", ""),
                              standard=body.get("standard", ""),
                              currency=body.get("currency", ""),
                              te=_num_or_none(body.get("te")),
                              sad=_num_or_none(body.get("sad")),
                              rates=body.get("rates"),
                              restricted=body.get("restricted"),
                              prior_lead=_pl,
                              cutoff_fields=_cutoff_fields_of(body))
    except Exception as e:
        import traceback
        traceback.print_exc()
        return ctx.err("生成银行存款底稿失败：" + str(e))
    if data is None:
        return ctx.err(info.get("error") or "生成失败")
    name = "货币资金BKD_%s.xlsx" % time.strftime("%Y%m%d_%H%M%S")
    name, latest = _write_out(name, data)
    info["file"] = name
    info["latestFile"] = latest
    if latest and latest != "%s_最新.xlsx" % name.split("_")[0]:
        info.setdefault("warnings", []).append(
            "「%s_最新.xlsx」正被 Excel 打开，已另存为 %s。请看新文件，或先关掉旧文件再生成。"
            % (name.split("_")[0], latest))
    info["url"] = "/download?f=" + urllib.parse.quote(name)
    info["prior"] = pinfo
    if _pl_err:
        info.setdefault("warnings", []).append(_pl_err)
    return ctx.json(info)


def _write_out(name, data):
    """写盘：时间戳文件 + 一份固定名「_最新」副本。返回 (name, 实际写成的最新文件名)。

    为什么要有固定名：bkd_out 里堆着每次生成的文件，用户按文件名分不清哪份是
    刚生成的，一不小心打开几小时前那份，还会以为"改了没生效"。

    **`_最新` 被 Excel 打开着时 Windows 会拒绝写入**（同目录会有 `~$xxx.xlsx`
    锁文件）。这种情况绝不能静默跳过——用户会一直盯着那份旧内容看，然后说
    "你根本没改"。所以这里退到 `_最新(2)`、`_最新(3)`…… 并把实际名字返回，
    由调用方提示用户。
    """
    with open(os.path.join(BKD.OUT, name), "wb") as fh:
        fh.write(data)
    stem = name.split("_")[0]
    cands = ["%s_最新.xlsx" % stem] + ["%s_最新(%d).xlsx" % (stem, i) for i in range(2, 10)]
    for cand in cands:
        try:
            with open(os.path.join(BKD.OUT, cand), "wb") as fh:
                fh.write(data)
            return name, cand
        except OSError:
            continue
    return name, ""


@_get("/api/build/exp/tb")
def api_build_exp_tb(ctx, q):
    """销管费用取数预览：前 10 行 + 合计。让用户先看数，再决定要不要生成。"""
    from . import expense as EXP
    conn = getattr(ctx.ledger, "conn", None)
    pf = (q.get("from") or ["1"])[0]
    pt = (q.get("to") or ["6"])[0]
    try:
        rows, meta = EXP.read_expense(conn, pf, pt)
    except Exception as e:
        return ctx.json({"error": "读取失败：" + str(e)})
    if meta.get("error"):
        return ctx.json({"error": meta["error"]})
    by_l1 = {}
    for r in rows:
        k = r.get("l1") or r.get("name") or "其他"
        by_l1[k] = round(by_l1.get(k, 0) + (r.get("amount") or 0), 2)
    return ctx.json({
        "rows": meta["rows"], "sum": meta["sum"], "books": meta["books"],
        "period": meta["period"], "closeOut": meta.get("close_out"),
        "byL1": by_l1,
        "sample": [{k: r.get(k) for k in ("book", "code", "name", "amount")}
                   for r in rows[:10]],
    })


@_get("/api/build/exp/plan")
def api_build_exp_plan(ctx, q):
    conn = getattr(ctx.ledger, "conn", None)
    pf = (q.get("from") or ["1"])[0]
    pt = (q.get("to") or ["6"])[0]
    try:
        _data, info = _run_exp(conn, pf, pt, preview=True)
    except Exception as e:
        return ctx.json({"error": "预览失败：" + str(e)})
    if info.get("error"):
        return ctx.json({"error": info["error"]})
    return ctx.json(info)


@_post("/api/build/exp/map")
def api_build_exp_map(ctx, u, body):
    """载入「主体 mapping 表」（账簿 → 公司）。

    走 /api/upload 已上传的文件（body 里传 fid），不在这里再解析 multipart——
    上传通道是现成的，少一套边界处理就少一处出错。
    """
    fid = (body or {}).get("fid")
    if not fid:
        return ctx.err("未收到文件 id，请先上传对照表")
    import glob
    hits = glob.glob(os.path.join(BKD.KZ_DATA, "uploads", str(fid) + "_*"))
    if not hits:
        return ctx.err("找不到已上传的对照表文件（fid=%s），请重新上传" % fid)
    path = hits[0]
    from . import expense as EXP
    try:
        with open(path, "rb") as fh:
            pairs = EXP.parse_book_map_xlsx(fh.read())
    except Exception as e:
        return ctx.err("解析对照表失败：" + str(e))
    if not pairs:
        return ctx.err("没解析到「账簿 ↔ 公司」两列。请确认表里有类似「账簿名称 / 公司名称」的表头。")
    n = EXP.save_book_map(pairs)
    sample = list(pairs.items())[:8]
    return ctx.json({"ok": True, "pairs": len(pairs), "total": n,
                     "sample": [{"book": k, "company": v} for k, v in sample]})


@_post("/api/build/exp/prior")
def api_build_exp_prior(ctx, u, body):
    """载入「上年度费用底稿」，回读校验：能认出多少 (账簿, 科目) 键。

    只做校验与预览，不落库——真正回填在生成时由 prior_fid 触发，
    这样同一份底稿既能给费用用，也不会把上年数写死在本年账套里。
    """
    fid = (body or {}).get("fid")
    if not fid:
        return ctx.err("未收到文件 id，请先上传上年费用底稿")
    pm, perr = _load_prior_expense(fid)
    if perr:
        return ctx.err(perr)
    if not pm:
        return ctx.err("上年费用底稿里没读到「上期审定数」列，或该列没有数值")
    with_book = [k for k in pm if k[0]]
    sample = [{"book": k[0], "name": k[1], "amount": pm[k]} for k in with_book[:8]]
    return ctx.json({"ok": True, "keys": len(pm), "withBook": len(with_book),
                     "sample": sample})


@_get("/api/build/exp/map")
def api_build_exp_map_info(ctx, q):
    from . import expense as EXP
    p = EXP._book_map_path()
    if not os.path.exists(p):
        return ctx.json({"exists": False, "hint": "尚未上传账簿↔公司对照表"})
    try:
        with open(p, "r", encoding="utf-8") as fh:
            d = json.load(fh)
    except Exception:
        d = {}
    return ctx.json({"exists": True, "pairs": len(d),
                     "sample": [{"book": k, "company": v} for k, v in list(d.items())[:8]]})


def _num_or_none(v):
    try:
        return float(str(v).replace(",", "")) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


@_post("/api/build/exp/run")
def api_build_exp_run(ctx, u, body):
    body = body or {}
    conn = getattr(ctx.ledger, "conn", None)
    # 前端两个下拉选定的工作表（VC / VD）；不传则按全文件解析，保持旧行为
    _ps = body.get("prior_sheets")
    if isinstance(_ps, str):
        _ps = [x for x in re.split(r"[,，、]", _ps) if x.strip()]
    pm, perr = _load_prior_expense(body.get("prior_fid"), sheets=_ps)
    # 上一期 TOD 底稿：默认就在上面那份上年底稿里（同一册 WP），用 prior_fid 取字节。
    # 只有显式传了 prior_tod_fid 才另读一份，兼容旧页面/命令行。
    _ptod_fid = body.get("prior_tod_fid") or (
        body.get("prior_fid") if body.get("prior_tod_sheet") else None)
    _ptod, _ptod_err = _load_prior_tod(_ptod_fid)
    try:
        data, info = _run_exp(conn, body.get("period_from") or "1", body.get("period_to") or "6",
                              balance_date=body.get("balance_date", ""),
                              customer=body.get("customer", ""),
                              prior_map=pm or None,
                              pre_end=body.get("pre_end"),
                              oar_evidence=os.path.join(BKD.OUT, "OAR_evidence_费用.json"),
                              oar_chg_min=body.get("oar_chg_min"),
                              oar_chg_pct=float(body.get("oar_chg_pct") or 0.10),
                              te=_num_or_none(body.get("te")),
                              sad=_num_or_none(body.get("sad")),
                              standard=body.get("standard") or "",
                              currency=body.get("currency") or "",
                              assurance=body.get("assurance"),
                              prior_fid=body.get("prior_fid"),
                              cra=body.get("cra"),
                              prior_tod=_ptod,
                              prior_tod_sheet=body.get("prior_tod_sheet"))
        if perr:
            info["priorError"] = perr
        if _ptod_err:
            info["priorTodError"] = _ptod_err
    except Exception as e:
        import traceback
        traceback.print_exc()
        return ctx.err("生成费用底稿失败：" + str(e))
    if data is None:
        return ctx.err(info.get("error") or "生成失败")
    name = "费用BKD_%s.xlsx" % time.strftime("%Y%m%d_%H%M%S")
    _write_out(name, data)
    info["file"] = name
    info["url"] = "/download?f=" + urllib.parse.quote(name)
    return ctx.json(info)


def _load_prior_expense(fid, sheets=None):
    """上年费用底稿 -> {(账簿,科目名): 上期审定数}。取不到返回空，不阻断。

    sheets 给定时只读这些工作表（前端两个下拉选定的 VC / VD）；
    口径 mapping 侧仍读全文件，见 `_load_prior_oar`。
    """
    if not fid:
        return {}, ""
    import glob
    hits = sorted(glob.glob(os.path.join(BKD.KZ_DATA, "uploads", str(fid) + "_*")))
    if not hits:
        return {}, "找不到已上传的上年费用底稿（fid=%s）" % fid
    from . import expense as EXP
    try:
        with open(hits[0], "rb") as fh:
            return EXP.parse_prior_expense(fh.read(), only_sheets=sheets)
    except Exception as e:
        return {}, "解析上年费用底稿失败：%s" % e


@_post("/api/build/exp/prior-tod/plan")
def api_build_exp_prior_tod_plan(ctx, u, body):
    """探测「上一期 TOD」文件的工作表，供界面手动指定用哪一张。

    与 BKD 的上年底稿交互一致：拖进来 → 列出所有页 → 标出哪些是 TOD 页、
    属于年审模板（annual）还是预审+剩余模板（split）→ 给推荐项。
    """
    body = ctx.multipart()
    f = (body or {}).get("files", {}).get("prior_tod")
    if not f:
        return ctx.err("未收到上一期 TOD 文件")
    from . import expense_tod as ETOD
    probe = ETOD.probe_prior_tod_sheets(f.get("data"))
    if probe.get("error"):
        return ctx.err(probe["error"])
    fid = None
    try:
        up = ctx.ledger.add_file(f.get("filename") or "上一期TOD.xlsx",
                                 io.BytesIO(f.get("data") or b""),
                                 len(f.get("data") or b""))
        fid = up.get("file_id")
    except Exception:
        fid = None
    return ctx.json({"ok": True, "filename": f.get("filename"), "fid": fid,
                     "sheets": probe["sheets"], "recommended": probe["recommended"]})


@_post("/api/build/exp/prior/plan")
def api_build_exp_prior_plan(ctx, u, body):
    """探测上年费用底稿的工作表结构：BKD 页（VC / VD）+ 上一期 TOD 页，一次给全。

    上一期 TOD 本来就在上一期底稿这同一册 WP 里，界面只该有一个上传口；
    所以这里一次解析同时认两类页，前端不再单独传「上一期 TOD」。
    只探测不落库，最终选哪张由用户定，生成时按 `prior_sheets` /
    `prior_tod_sheet` 再来一次（同一个 fid）。
    """
    body = ctx.multipart()
    files = (body or {}).get("files", {})
    f = files.get("prior") or files.get("prior_tod")
    if not f:
        return ctx.err("未收到上年底稿文件")
    from . import expense as EXP
    probe = EXP.probe_prior_all(f.get("data"))
    if probe.get("error"):
        return ctx.err(probe["error"])
    fid = None
    try:
        up = ctx.ledger.add_file(f.get("filename") or "上年度费用底稿.xlsx",
                                 io.BytesIO(f.get("data") or b""),
                                 len(f.get("data") or b""))
        fid = up.get("file_id")
    except Exception:
        fid = None
    return ctx.json({"ok": True, "filename": f.get("filename"), "fid": fid,
                     "sheets": probe["sheets"],
                     "recVC": probe["recVC"], "recVD": probe["recVD"],
                     "tod": probe.get("tod") or {"sheets": [], "recommended": -1}})


def _load_prior_tod(fid):
    """上一期 TOD 底稿的原始字节。

    这里**不解析**：整页搬运要连格式和颜色一起保留，解析成结构反而丢了样式。
    解析与搬运交给 expense_tod.copy_prior_tod_as_sheet。
    """
    if not fid:
        return None, ""
    import glob
    hits = sorted(glob.glob(os.path.join(BKD.KZ_DATA, "uploads", str(fid) + "_*")))
    if not hits:
        return None, "找不到已上传的上一期 TOD 底稿（fid=%s）" % fid
    try:
        with open(hits[0], "rb") as fh:
            return fh.read(), ""
    except Exception as e:
        return None, "读取上一期 TOD 底稿失败：%s" % e


def _load_prior_oar(fid):
    """上年费用底稿 → ({(主体,科目):(口径,来源)}, {主体:简称}, note)。

    单独读一次文件：`_load_prior_expense` 只取上期审定数，口径 mapping 要的是
    「上年这个科目算哪个口径」，两者用途不同，别硬塞进同一个返回值。
    """
    if not fid:
        return {}, {}, ""
    import glob
    hits = sorted(glob.glob(os.path.join(BKD.KZ_DATA, "uploads", str(fid) + "_*")))
    if not hits:
        return {}, {}, ""
    from . import oar_map as OM
    try:
        with open(hits[0], "rb") as fh:
            return OM.read_prior(fh.read())
    except Exception as e:
        return {}, {}, "上年口径解析失败：%s" % e



def _load_prior_lead(fid, spec=None):
    """按上传文件 id 读上年 Lead。

    spec 与 BKD 那一套同构：{"sheet": "C.00 Lead", "header_row": 10}。
    给了就按指定的表/表头行解析，没给才走"谁读到的数最多"的自动择优。
    """
    if not fid:
        return None
    import glob
    hits = sorted(glob.glob(os.path.join(BKD.KZ_DATA, "uploads", str(fid) + "_*")))
    if not hits:
        return {"error": "找不到已上传的上年 Lead（fid=%s）" % fid}
    try:
        with open(hits[0], "rb") as fh:
            raw = fh.read()
        out = _parse_prior_lead_bytes(raw, spec)
        # 整表复制要用原始字节：解析只取科目/金额，复制要连格式一起搬
        if isinstance(out, dict):
            out["data"] = raw
        return out
    except Exception as e:
        return {"error": "读取上年 Lead 失败：%s" % e}


def _parse_prior_lead_bytes(data, spec=None):
    """上年 Lead 字节流 -> {"t1":[{subject,audited}], "t2":[...], "meta":{...}}。

    只认「科目名称」这一列：表1 的列位在不同版本里挪过，但科目名是人写的，稳定。

    **要择优，不能拿第一张含 Lead 的表就用**：上年底稿里常常同时有「C.00 Lead」
    和一个空的「上年 C.00 Lead」占位表；先取到空表，上年数就整列丢了。
    所以每张候选表都试读一次，**取读到的金额条数最多的那张**。
    """
    try:
        wb = BKD.load_wb_any(data, data_only=True)
    except Exception as e:
        return {"error": "读取上年 Lead 失败：%s" % e}

    def txt(v):
        return str(v or "").replace(" ", "").replace("\u3000", "")

    def try_sheet(sh, hdr=None):
        """返回 (hdr, subj_col, amt_col, rows) 或 None。

        hdr 给了就只认那一行（用户在前端指定的表头行），不再扫前 60 行——
        否则用户改了行号却没生效，界面看着选中、结果用的是别处。
        """
        rows_scan = [hdr] if hdr else list(range(1, min(sh.max_row, 60) + 1))
        for r in rows_scan:
            row = {c: txt(sh.cell(r, c).value)
                   for c in range(1, min(sh.max_column, 20) + 1)}
            sc = next((c for c, t in row.items() if t == "科目名称"), None)
            if not sc:
                continue
            ac = next((c for c, t in row.items() if "期末审定数" in t), None)
            if not ac:
                continue
            rows = []
            for rr in range(r + 1, sh.max_row + 1):
                subj = txt(sh.cell(rr, sc).value)
                if not subj:
                    continue
                if ("合计" in subj or subj.startswith("说明") or "表2" in subj
                        or "审定数" in subj or subj in ("报表科目", "科目", "项目", "Rx")):
                    break
                v = sh.cell(rr, ac).value
                rows.append({"subject": subj,
                             "audited": float(v) if isinstance(v, (int, float)) else None})
            return r, sc, ac, rows
        return None

    # ① 用户指定了工作表：只听用户的，失败也要说清失败在哪
    _spec = spec or {}
    _want = (_spec.get("sheet") or "").strip()
    if _want:
        if _want not in wb.sheetnames:
            return {"error": "上年底稿里没有工作表「%s」（现有：%s）"
                             % (_want, "、".join(wb.sheetnames[:10]))}
        _hdr = None
        try:
            _hdr = int(_spec.get("header_row") or 0) or None
        except (TypeError, ValueError):
            _hdr = None
        got = try_sheet(wb[_want], _hdr)
        if not got:
            return {"error": "工作表「%s」%s 找不到「科目名称 + 期末审定数」表头，"
                             "请核对表头行号后重试"
                             % (_want, "第 %d 行" % _hdr if _hdr else "")}
        _h, _sc, _ac, _rows = got
        if not _rows:
            return {"error": "工作表「%s」第 %d 行下方没有科目行，请确认表头行是否选对"
                             % (_want, _h)}
        t1 = _rows
        name, ws = _want, wb[_want]
        best = (len([x for x in t1 if x["audited"] is not None]), name, ws, got)
    else:
        best = None
    # 不按表名过滤：真实文件常叫 Sheet1（本次样本就是），按名字找会整块丢掉。
    # 表名带 lead 的只加一点点权重，决定权在"读到多少条审定数"。
    for n in ([] if _want else wb.sheetnames):
        got = try_sheet(wb[n])
        if not got:
            continue
        score = len([x for x in got[3] if x["audited"] is not None])
        if "lead" in n.lower():
            score += 0.5
        if not best or score > best[0]:
            best = (score, n, wb[n], got)

    if not best:
        cand = [n for n in wb.sheetnames if "Lead" in n]
        return {"error": "上年文件里%s找不到「科目名称＋期末审定数」表头"
                         % ("（含 Lead 的表：%s）" % "、".join(cand) if cand else "")}

    _, name, ws, (hdr, subj_col, amt_col, t1) = best

    # ---- 表2：找「报表科目」表头
    t2 = []
    for r in range(1, ws.max_row + 1):
        if txt(ws.cell(r, 2).value) != "报表科目":
            continue
        for rr in range(r + 1, ws.max_row + 1):
            item = txt(ws.cell(rr, 2).value)
            if not item:
                continue
            if "合计" in item:
                break
            cur = ws.cell(rr, 3).value
            pri = ws.cell(rr, 4).value
            t2.append({"item": item,
                       "current": float(cur) if isinstance(cur, (int, float)) else None,
                       "prior": float(pri) if isinstance(pri, (int, float)) else None})
        break

    meta = {"sheet": name}
    for r in range(1, 12):
        label = txt(ws.cell(r, 2).value)
        if "资产负债表日" in label and ws.cell(r, 3).value:
            meta["date"] = str(ws.cell(r, 3).value)[:10]
        if "客户名称" in label and ws.cell(r, 3).value:
            meta["customer"] = str(ws.cell(r, 3).value)
    if not t1:
        return {"error": "上年 Lead 表1 没读到科目行（表头在第 %d 行）" % hdr}
    return {"t1": t1, "t2": t2, "meta": meta}


def _load_prior(fid, spec=None):
    """把上传的上年 WP 解析成账户清单，供本年 BKD / OAR 回捞上年数。

    spec 是 AI 确认过的 {sheet, header_row}（「导出结构探测包」那一套）；
    没给就自动挑匹配分最高的表。两种都失败就返回空，不阻断（上年数是可选信息）。
    """
    if not fid:
        return [], {"loaded": False, "reason": "未上传上年底稿"}
    import glob
    hits = sorted(glob.glob(os.path.join(BKD.KZ_DATA, "uploads", str(fid) + "_*")))
    if not hits:
        return [], {"loaded": False, "reason": "找不到已上传的上年 WP"}
    try:
        with open(hits[0], "rb") as fh:
            data = fh.read()
    except Exception as e:
        return [], {"loaded": False, "reason": "读文件失败：%s" % e}
    if spec and spec.get("sheet"):
        cfg = [{"sheet": spec["sheet"], "header_row": int(spec.get("header_row") or 1)}]
        picked = "AI 指定：%s 第 %s 行" % (spec["sheet"], spec.get("header_row"))
    else:
        try:
            wb = BKD.load_wb_any(data, data_only=True)
            pr = BKD.probe_prior_structure(wb)
            try:
                wb.close()
            except Exception:
                pass
            best = pr.get("best") or {}
            if not best.get("name"):
                return [], {"loaded": False, "reason": "没能从上年底稿里认出 BKD 明细表"}
            cfg = [{"sheet": best["name"], "header_row": best["header_row"]}]
            picked = "自动识别：%s 第 %s 行（匹配分 %s）" % (
                best["name"], best["header_row"], best.get("score"))
        except Exception as e:
            return [], {"loaded": False, "reason": "结构探测失败：%s" % e}
    try:
        prior, pinfo = BKD.parse_prior_xlsx(data, cfg)
    except Exception as e:
        return [], {"loaded": False, "reason": "解析失败：%s" % e}
    return prior, {"loaded": bool(prior), "rows": len(prior), "picked": picked,
                   "sheets": pinfo}


def _apply_rates(rows, rates):
    """按主体套折算率：rate 支持 {主体: 率} 或 [{company, rate}]，缺省 1。

    为什么按主体给：一个客户名下可能有境内公司（本位币人民币，率=1）和境外公司
    （本位币港币/美元，要把本币再折成人民币），同一次上数里两种都要对。
    """
    m = {}
    if isinstance(rates, list):
        for it in rates:
            if isinstance(it, dict) and it.get("company"):
                m[str(it["company"]).strip()] = it.get("rate")
    elif isinstance(rates, dict):
        for k, v in rates.items():
            m[str(k).strip()] = v

    def _f(v):
        try:
            f = float(v)
        except (TypeError, ValueError):
            return 1.0
        return f if f > 0 else 1.0

    fx_rows = []
    for r in rows:
        rate = _f(m.get(str(r.get("company") or "").strip(), 1))
        r["rate"] = rate
        r["rmb"] = round((r.get("local") or 0) * rate, 2)
        if abs(rate - 1.0) > 1e-12:
            fx_rows.append(r)
    return fx_rows


def _apply_restricted(rows, restricted):
    """把受限资金金额打到对应账户上。

    键有三种写法，按精确度从高到低：
      「一级科目|账号」或「一级科目:账号」 —— 同一账号跨科目时**必须**用这种
      「账号」                           —— 账号在全部账户里唯一时才可用
      「公司|银行名」                     —— 没有账号时的兜底

    为什么必须区分科目：真实账套里同一个银行账号会同时挂在「银行存款」和
    「其他货币资金」下（一般户 + 保证金户），只按账号匹配会给两行都打上同一个
    金额，受限合计直接翻倍——而且不报错。
    """
    from core.accounts import _acct_norm_py

    def _resolve(r):
        """给一行算出可以拿来匹配的几种键。"""
        acct = _acct_norm_py(r.get("acctNo") or "")
        subj = BKD._norm_subject_l1(r.get("subject")) or BKD.T(r.get("subject"))
        bank = BKD.T(r.get("bank") or r.get("bankName"))
        company = BKD.T(r.get("company"))
        keys = {}
        if acct:
            keys["%s|%s" % (subj, acct)] = ("subj_acct", subj, acct)
            keys[acct] = ("acct", subj, acct)
        if company or bank:
            keys["%s|%s" % (company, bank)] = ("co_bank", subj, "")
        return keys

    # 建索引：键 -> 行号列表
    idx = {}
    for i, r in enumerate(rows):
        for k in _resolve(r):
            idx.setdefault(k, []).append(i)

    # 归一用户给的键（允许全角冒号、空格）
    pairs = []
    if isinstance(restricted, dict):
        src = list(restricted.items())
    elif isinstance(restricted, list):
        src = [("%s|%s" % (it.get("subject") or "", it.get("acctNo") or it.get("acct") or ""),
                it.get("amount")) for it in restricted if isinstance(it, dict)]
    else:
        src = []
    for k, v in src:
        kk = str(k or "").replace("：", "|").replace(":", "|").strip()
        kk = re.sub(r"\s*\|\s*", "|", kk)
        # 只给了账号（没有科目前缀）
        if "|" not in kk:
            kk = _acct_norm_py(kk)
        else:
            subj, _, acct = kk.partition("|")
            acct = _acct_norm_py(acct)
            kk = "%s|%s" % (subj.strip(), acct) if acct else subj.strip()
        try:
            pairs.append((kk, float(v or 0)))
        except (TypeError, ValueError):
            raise ValueError("受限金额不是数字：%r" % (v,))

    hit, ambiguous = 0, []
    for kk, amt in pairs:
        ids = idx.get(kk) or []
        if not ids and "|" in kk:
            # 用户写了「科目|账号」但科目名不匹配（如写"银行存款"而账上是"银行"）
            acct = _acct_norm_py(kk.partition("|")[2])
            ids = idx.get(acct) or []
        if not ids:
            continue
        if len(ids) > 1:
            cand = []
            for i in ids:
                cand.append("%s/%s" % (BKD._norm_subject_l1(rows[i].get("subject")),
                                       rows[i].get("acctNo")))
            ambiguous.append("%s 命中 %d 个账户（%s）" % (kk, len(ids), "、".join(cand[:4])))
            continue
        i = ids[0]
        rows[i]["restricted"] = round(rows[i].get("restricted", 0) + amt, 2)
        hit += 1
    if ambiguous:
        raise ValueError(
            "受限金额的账号不唯一，请加上一级科目再指定（如 其他货币资金|账号=金额）："
            + "；".join(ambiguous))
    return hit


def _company_list(rows):
    """列出主体及是否有非人民币账户，供前端决定给谁填折算率。"""
    m = {}
    for r in rows:
        name = str(r.get("company") or "").strip() or "(未标注主体)"
        d = m.setdefault(name, {"company": name, "rows": 0, "hasFx": False, "currencies": []})
        d["rows"] += 1
        # 币种现在是 ISO 代码（CNY/USD…），判断"是不是外币"必须按代码比，
        # 拿中文字面量比会把所有人民币主体都标成"含外币 ⚠"。
        cur = BKD.cur_code(r.get("currency"))
        if cur and cur not in d["currencies"]:
            d["currencies"].append(cur)
        if cur and cur != "CNY":
            d["hasFx"] = True
    return sorted(m.values(), key=lambda x: x["company"])


def _post_gen(subject, data, info, ctx_extra=None, params=None, check_only=False,
              save_fp=True):
    """生成之后统一做三件事：落指纹、跑不变量、与基准比对。

    放在一处而不是散在各生成器里，是因为后面每加一个科目都要这三件，
    散着写必然会有一个科目漏掉其中一件。

    三件都不阻断生成——它们是"观测"，不是"闸门"。闸门在 `--strict` 那一层。
    """
    from core import fingerprint as FP, invariants as INV, baseline as BL
    out = {}
    if not check_only:
        try:
            fp = FP.of_bytes(data, params=params or {})
            fp_path = FP.path_for(BKD.KZ_DATA, subject)
            old_fp = FP.load(fp_path)
            fp_diffs = FP.diff(old_fp, fp) if old_fp else []
            out["fingerprint"] = {"path": fp_path, "diffs": fp_diffs,
                                  "blocking": len(FP.blocking(fp_diffs)),
                                  "saved": bool(save_fp)}
            # strict 模式不覆盖指纹：否则这次覆盖了，下次就没得比了
            if save_fp:
                FP.save(fp, fp_path)
        except Exception as e:
            out["fingerprint"] = {"error": str(e)[:200]}
    try:
        ctx = {"kind": subject, "info": info}
        ctx.update(ctx_extra or {})
        out["invariants"] = INV.run(ctx, "cash" if subject == "cash" else "exp",
                                    data_dir=BKD.KZ_DATA)
    except Exception as e:
        out["invariants"] = {"error": str(e)[:200]}
    try:
        base = BL.capture("cash" if subject == "cash" else "exp", info)
        bpath = BL.path_for(BKD.KZ_DATA, "cash" if subject == "cash" else "exp")
        old = BL.load(bpath)
        if old:
            diffs = BL.compare(base, old)
            summary = BL.summarize(diffs)
        else:
            # 首次运行没有基准可比。以前这里返回「与基准完全一致」，
            # 语义是反的——会让人以为"验证过了"，其实只是刚建了基准。
            diffs, summary = [], "首次运行，已建立基准"
        out["baseline"] = {"path": bpath, "diffs": diffs, "summary": summary}
        BL.save(base, bpath)
    except Exception as e:
        out["baseline"] = {"error": str(e)[:200]}
    info.update(out)
    return info


def _run_c00(conn, prior=None, balance_date="", customer="", preview=False, rates=None,
             cutoff_fields=None, strict=False, restricted=None,
             standard="", currency="", te=None, sad=None, prior_lead=None):
    """跑一次银行存款 BKD。preview=True 时只算不落盘。

    cutoff_fields 给了就跑 C.03 截止性测试（同 /api/generate 的口径）；
    截止测试失败不阻断主底稿，错误放在 info["cutoffError"] 里交给前端提示。
    """
    records, meta = BKD.read_tb_from_duckdb("货币资金", conn=conn)
    if meta.get("error"):
        return None, {"error": meta["error"]}
    if not records:
        return None, {"error": "科目余额表里没有「库存现金 / 银行存款 / 其他货币资金」账户。"
                               "请先在「导入清洗」里合并科目余额表，并确认一级科目已识别。"}
    res = BKD.merge(records, prior or [])
    # 上年 Lead（可选）：科目清单 + 审定数，用来插行、回填 K 列与表2 上年数
    if prior_lead and not prior_lead.get("error"):
        res["priorLead"] = dict(prior_lead)
    elif prior_lead and prior_lead.get("error"):
        res.setdefault("warnings", []).append("上年 Lead 未使用：" + prior_lead["error"])
    fx_rows = _apply_rates(res["rows"], rates)
    if restricted:
        try:
            _apply_restricted(res["rows"], restricted)
        except ValueError as e:
            return None, {"error": "受限资金没能写入：%s" % e,
                          "code": "RESTRICTED_AMBIGUOUS"}

    cutoff, cutoff_err = None, ""
    # 截止期间一律等于资产负债表日。前端那个「本年 12-31」的默认值是历史遗留，
    # 它和资产负债表日各走各的，曾把整张 C.03 跑成 0 样本还不报错。
    if cutoff_fields and not preview:
        if not str(balance_date or "").strip():
            cutoff_err = ("未填资产负债表日，截止性测试未执行。"
                          "C.03 的截止期间必须等于资产负债表日，不接受单独填一个日期。")
        else:
            cutoff_fields = dict(cutoff_fields)
            cutoff_fields["cutoff_date"] = str(balance_date).strip()[:10]
            try:
                cutoff = _build_cutoff(conn, cutoff_fields)
            except ValueError as e:
                cutoff_err = str(e)
            except Exception as e:                      # 序时账出问题不该拖垮主底稿
                cutoff_err = "截止性测试未执行：%s" % e

    # 注意：这个 dict 要留存引用——build_workbook 会在里面回填 TE/SAD 与来源，
    # 用临时字面量传进去的话，出了函数就拿不到了（info.teSource 会永远是空）。
    # 上期日 = 上一年 12/31（OAR 的 D 列表头要用）
    try:
        _y = int(str(balance_date)[:4])
        prior_date = "%d-12-31" % (_y - 1)
    except (TypeError, ValueError):
        prior_date = ""
    tpl_meta = {"time": time.strftime("%Y-%m-%d %H:%M"), "tbDb": BKD.KZ_DB,
                "balance_date": balance_date, "prior_date": prior_date,
                "customer": customer,
                "standard": standard, "currency": currency,
                "te": te, "sad": sad}
    data, tpl = BKD.build_workbook(res, meta=tpl_meta, cutoff=cutoff)
    if data is None:
        return None, {"error": "未找到内置模板，请先上传 C SWP 货币资金模板。"}
    gaps = res.get("gaps") or []
    info = {
        "target": "c00", "tpl": os.path.basename(tpl or ""),
        "rows": len(res["rows"]),
        "gaps": len(gaps),
        "gapList": gaps[:300],
        "amount": round(sum(r["local"] for r in res["rows"]), 2),
        "rmbAmount": round(sum(r.get("rmb") or 0 for r in res["rows"]), 2),
        "fxCount": len(fx_rows),
        "companies": _company_list(res["rows"]),
        "bySubject": _by_subject(res["rows"]),
        "needsManual": _gap_fields(gaps),
        "priorLoaded": bool(prior),
        "cutoff": _cutoff_payload(cutoff),
        # 上年回捞的匹配情况：matched / newCount / missCount / sumPrior / diff
        "stats": res.get("stats") or {},
        "sample": [{k: r.get(k) for k in ("company", "acctNo", "nature", "confirm",
                                          "priorLocal", "_matched")}
                   for r in (res.get("rows") or [])[:3]],
        "changes": (res.get("changes") or [])[:10],
        # —— 受限资金：BKD 有列、Lead 有落点，两处不重复计数 ——
        "restrictedTotal": res.get("restrictedTotal") or 0,
        "restrictedMode": res.get("restrictedMode") or "",
        "restrictedTarget": res.get("restrictedTarget") or "",
        "restrictedCell": res.get("restrictedCell") or "",
        "restrictedNote": res.get("restrictedNote") or "",
        # —— 派生表执行情况（C.01b / C.00b / C.02 / C.02a）——
        "derived": res.get("derived") or {},
        # 截图状态：找到没、插进去没、为什么没插（以前静默失败，查不出来）
        "c02aShot": res.get("c02aShot") or "",
        "c02aShotEmbedded": bool(res.get("c02aShotEmbedded")),
        "c02aShotError": res.get("c02aShotError") or "",
        # —— 阈值来源：模板没填 TE 时会带 TE_NOT_READ 警告 ——
        "teSource": res.get("teSource") or "",
        "te": res.get("te"), "sad": res.get("sad"),
        "teWarnings": res.get("teWarnings") or [],
        # —— Lead 行号定位结果：写到了哪几行。换模板后第一眼就看这个对不对 ——
        "leadRows": res.get("leadRows") or {},
        # —— 主底稿生成过程中攒下的告警（如 Lead 行号定位失败退回 V7 兜底）——
        "buildWarnings": res.get("warnings") or [],
        # —— 待办任务（截图 / 对账单 / 判断项）：顶层结构化，调用方可机器核对 ——
        "tasks": res.get("tasks") or [],
        # 前 10 行给界面「先看数再生成」用；完整 162 行在 Excel 里，不往 JSON 里塞
        "previewRows": [{k: r.get(k) for k in (
            "company", "subject", "bank", "acctNo", "currency",
            "orig", "local", "rmb", "nature", "confirm", "priorLocal")}
            for r in (res.get("rows") or [])[:10]],
    }
    if cutoff_err:
        info["cutoffError"] = cutoff_err
    if not preview:
        _post_gen("cash", data, info, ctx_extra={"rows": res["rows"], "rates": rates or {}},
                  params={"period": balance_date, "customer": customer,
                          "fxCount": info.get("fxCount")}, save_fp=not strict)
    return (None if preview else data), info


def _run_exp(conn, period_from, period_to, balance_date="", customer="", preview=False,
             prior_map=None, pre_end=None, oar_evidence=None, oar_conclusions=None,
             oar_chg_min=None, oar_chg_pct=0.10, te=None, sad=None,
             standard="", currency="", assurance=None,
             prior_fid=None, prior_oar=None, prior_shorts=None, prior_note="",
             cra=None, prior_tod=None, prior_tod_sheet=None):
    from . import expense as EXP
    tpl_dir = BKD.TPL_DIR          # data/bkd_tpl，与货币资金模板同目录
    if prior_oar is None and prior_fid:
        prior_oar, prior_shorts, prior_note = _load_prior_oar(prior_fid)
    if cra is None:
        # 没显式传就用「销管费用」那套 CRA 配置（默认 50/25/15/10，与模板一致）
        try:
            from . import baseinfo as BI
            cra = BI.cra_of(BI.load(BKD.KZ_DATA), "销管费用")
        except Exception:
            cra = None
    data, info = EXP.build_expense_workbook(
        conn, tpl_dir, period_from=period_from, period_to=period_to,
        balance_date=balance_date, customer=customer, prior_map=prior_map,
        pre_end=pre_end, oar_evidence=oar_evidence, oar_conclusions=oar_conclusions,
        oar_chg_min=oar_chg_min, oar_chg_pct=oar_chg_pct, te=te, sad=sad,
        standard=standard, currency=currency, assurance=assurance,
        prior_oar=prior_oar, prior_shorts=prior_shorts, prior_note=prior_note or "",
        cra=cra, prior_tod=prior_tod, prior_tod_sheet=prior_tod_sheet)
    if data is None:
        return None, info
    info["target"] = "exp"
    if not preview:
        _post_gen("exp", data, info, params={"period": "%s-%s" % (period_from, period_to),
                                             "customer": customer})
    return (None if preview else data), info


def _by_subject(rows):
    """按一级科目汇总。必须归一——底稿里写的是归一后的科目，
    报表这里如果直接用原始科目名，会出现「银行存款_某某支行（账号）」这种
    一条一行的统计，看着像多了几个科目。"""
    out = {}
    for r in rows:
        k = BKD._norm_subject_l1(r.get("subject")) or (r.get("subject") or "其他")
        out[k] = round(out.get(k, 0) + (r.get("local") or 0), 2)
    return out


def _gap_fields(gaps):
    out = {}
    for g in gaps:
        out[g.get("field", "")] = out.get(g.get("field", ""), 0) + 1
    return out


# ---------------------------------------------------------------- 实现
def _tpl_info(ctx):
    tp = BKD._find_tpl()
    if not tp:
        return ctx.json({"exists": False, "version": BKD.TPL_VERSION,
                         "hint": "未找到模板，请上传"})
    return ctx.json({
        "exists": True, "version": BKD.TPL_VERSION,
        "filename": os.path.basename(tp),
        "size": os.path.getsize(tp),
        "note": "安永 Skywind 标准工作底稿：C SWP 货币资金（最新 V7）。可在下方上传替换/下载。",
    })


def _tpl_download(ctx):
    tp = BKD._find_tpl()
    if not tp:
        return ctx.err("未找到模板", 404)
    size = os.path.getsize(tp)
    return ctx.send_file(tp,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        os.path.basename(tp))


def _tpl_upload(ctx, body):
    files = body.get("files", {})
    f = files.get("tpl")
    if not f:
        return ctx.err("未收到模板文件")
    fn = f["filename"] or "模板.xlsx"
    if not fn.lower().endswith((".xlsx", ".xlsm")):
        return ctx.err("模板仅支持 .xlsx / .xlsm；.xls 老格式请先用 Excel 另存为 .xlsx"
                       "（模板的格式、公式、条件格式无法在格式转换中保留）")
    try:
        wb = BKD.load_workbook(io.BytesIO(f["data"]))
    except Exception as e:
        return ctx.err("模板无法解析：" + str(e))
    missing = [s for s in ("C.00 BKD", "C.00 Lead") if s not in wb.sheetnames]
    if missing:
        return ctx.err("该文件不是标准 C SWP 货币资金模板，缺少工作表：" + "、".join(missing))
    with open(BKD.TPL_DEFAULT, "wb") as fh:
        fh.write(f["data"])
    return ctx.json({"ok": True, "filename": os.path.basename(BKD.TPL_DEFAULT)})


def _lead_probe(grid):
    """在已读出的网格里找「科目名称 + 期末审定数」表头。

    返回 (表头行号, 科目行数, 有金额的科目数)；找不到返回 None。
    用它给前端标「（推荐）/（未识别）」并预填表头行，与 BKD 那张表的下拉一一对应。
    """
    if not grid:
        return None
    for r, row in enumerate(grid):
        cells = [BKD.T(v).replace(" ", "").replace("\u3000", "") for v in row]
        if "科目名称" not in cells:
            continue
        ac = next((i for i, t in enumerate(cells) if "期末审定数" in t), None)
        if ac is None:
            continue
        sc = cells.index("科目名称")
        n = nv = 0
        for rr in range(r + 1, len(grid)):
            line = grid[rr]
            subj = BKD.T(line[sc] if sc < len(line) else "").replace(" ", "")
            if not subj:
                continue
            if ("合计" in subj or subj.startswith("说明") or "表2" in subj
                    or "审定数" in subj or subj in ("报表科目", "科目", "项目", "Rx")):
                break
            n += 1
            v = line[ac] if ac < len(line) else None
            if isinstance(v, (int, float)):
                nv += 1
        return r + 1, n, nv
    return None


def _prior_plan(ctx, body):
    files = body.get("files", {})
    f = files.get("prior")
    if not f:
        return ctx.err("未收到上年底稿文件")
    try:
        wb = BKD.load_wb_any(f["data"], data_only=True, read_only=True)
        sheets = []
        for ws in wb.worksheets:
            grid = [[c.value for c in row] for row in ws.iter_rows()]
            hd = BKD.detect_header(grid, BKD.BKD_SPEC)
            headers = [BKD.T(c) for c in (grid[hd["i"]] if hd["i"] < len(grid) else [])]
            mp = BKD.map_columns(headers, BKD.BKD_SPEC)
            usable = bool(mp.get("acctNo") or mp.get("bankName") or mp.get("subject"))
            lp = _lead_probe(grid)
            sheets.append({
                "sheet": ws.title, "header_row": hd["i"] + 1, "score": hd["score"],
                "headers": headers, "map": mp, "usable": usable,
                # 上年 Lead 的候选信息：表头行、科目数、已读到金额的科目数
                "leadHeaderRow": (lp[0] if lp else None),
                "leadRows": (lp[1] if lp else 0),
                "leadAmountCells": (lp[2] if lp else 0),
                "leadUsable": bool(lp and lp[2] > 0),
            })
        wb.close()
        # 同时把文件入库，后续生成 BKD / OAR 时按 fid 重新解析（不用前端再传一遍文件）
        try:
            up = ctx.ledger.add_file(f["filename"], io.BytesIO(f["data"]), len(f["data"]))
            fid = up.get("file_id")
        except Exception:
            fid = None
        return ctx.json({"sheets": sheets, "filename": f["filename"], "fid": fid,
                         "recommended": _recommend_prior_sheet(sheets),
                         "leadRecommended": _recommend_lead_sheet(sheets)})
    except Exception as e:
        return ctx.err("解析上年底稿失败：" + str(e))


def _recommend_lead_sheet(sheets):
    """猜一张「上年 Lead」。

    优先：名字带 Lead 且真的读到审定数；其次：任何读到审定数最多的那张
    （客户可能把 Lead 叫「底稿汇总」，名字靠不住，金额条数才靠得住）。
    """
    best_i, best_n = -1, 0
    for i, s in enumerate(sheets):
        nv = int(s.get("leadAmountCells") or 0)
        if nv <= 0:
            continue
        named = 1 if re.search(r"lead", s.get("sheet", ""), re.I) else 0
        # 名字带 Lead 的加权，避免同分时选到别的表
        score = nv * 10 + named * 5
        if score > best_n:
            best_i, best_n = i, score
    return best_i


def _recommend_prior_sheet(sheets):
    """猜一张「上年 BKD 明细」工作表。

    标准模板里 C.00 Lead 是汇总表、C.01/C.01b/C.01c 是函证程序，都能被识别出账号列，
    但只有 C.00 BKD 是逐户明细。所以按名字优先级挑，而不是简单取第一个可用表。
    """
    def _usable(s):
        return bool(s.get("usable"))

    # 1) 名字直接带 BKD 的（C.00 BKD）
    for i, s in enumerate(sheets):
        if _usable(s) and re.search(r"bkd", s.get("sheet", ""), re.I):
            return i
    # 2) 名字带「货币资金明细 / 银行账户清单」之类
    for i, s in enumerate(sheets):
        n = s.get("sheet", "")
        if _usable(s) and ("货币资金" in n or "银行账户" in n or "账户清单" in n):
            return i
    # 3) 第一个可用的
    for i, s in enumerate(sheets):
        if _usable(s):
            return i
    return 0 if sheets else -1


def _build_cutoff(conn, fields):
    """C.03 截止性测试：只算参数，样本必须取自银行流水。

    **不从序时账推导。** 序时账只有科目、凭证号、金额，看不到收付款双方、账号与到账时间。
    靠「同一凭证号一借一贷」或「跨日金额相等」去凑配对，凑出来的是推测不是证据——
    内部调拨、同一账户的两笔不同业务，都能凑成一对。截止性测试要证的是
    「这笔钱什么时候真正进出银行」，只有对账单口径的流水能证明。

    当前平台还没有接入银行流水，所以本期**样本区留空**：把截止期间、阈值算好写进表，
    在 B11 写明缺什么、取得流水后怎么补。宁可空着，也不拿凑出来的配对充数。
    """
    if str(fields.get("cutoff", "1")) == "0":
        return None
    end = BKD._as_date(fields.get("cutoff_date"))
    if end is None:
        raise ValueError("请填写截止期间（资产负债表日）")
    bd = max(0, int(float(fields.get("cutoff_before") or 1)))
    ad = max(0, int(float(fields.get("cutoff_after") or 1)))
    th = float(fields.get("cutoff_threshold") or 100000)
    date_from = end - timedelta(days=bd)
    date_to = end + timedelta(days=ad)
    stat = {
        "endDate": end.strftime("%Y-%m-%d"),
        "windowStart": date_from.strftime("%Y-%m-%d"),
        "windowEnd": date_to.strftime("%Y-%m-%d"),
        "beforeDays": bd, "afterDays": ad, "threshold": th,
        "journalRows": 0, "transferPairs": 0, "beforeCount": 0, "afterCount": 0,
        "crossCount": 0, "belowThreshold": 0, "singleCount": 0,
        "shownBefore": 0, "shownAfter": 0, "flowRows": 0,
    }
    return {"before": [], "after": [], "single": [], "stat": stat,
            "source": "none", "warn": BKD._cutoff_no_flow_reason(stat)}


def _jsonify(o):
    """把 date / Decimal 转成能进 JSON 的值。

    cutoff 的行里带 vdate（DATE 类型），Decimal 也有。HTTP 那条路因为 ctx.json
    用了 default=str 才没炸，CLI 直接 json.dumps 就会报
    "Object of type date is not JSON serializable"。
    """
    if isinstance(o, dict):
        return {k: _jsonify(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_jsonify(v) for v in o]
    if isinstance(o, _dt) or isinstance(o, _date):
        return o.strftime("%Y-%m-%d")
    try:
        from decimal import Decimal
        if isinstance(o, Decimal):
            return float(o)
    except Exception:
        pass
    return o


def _cutoff_payload(c):
    if not c:
        return {"enabled": False}
    single = c.get("single") or []
    return _jsonify({
        "enabled": True,
        "stat": c.get("stat"),
        "before": c.get("before") or [],
        "after": c.get("after") or [],
        "warn": c.get("warn"),
        "journal": c.get("journal"),
        "single": single[:50],
        "singleTotal": len(single),
    })


def _generate(ctx, body):
    files = body.get("files", {})
    fields = body.get("fields", {})
    f = files.get("prior")
    if not f:
        return ctx.err("未收到上年底稿文件")
    try:
        kw = fields.get("keyword") or "银行存款"
        conn = getattr(ctx.ledger, "conn", None)
        records, meta = BKD.read_tb_from_duckdb(kw, conn=conn)
        if meta.get("error"):
            return ctx.err(meta["error"])
        sheets_cfg = json.loads(fields.get("sheets", "[]"))
        if not sheets_cfg:
            return ctx.err("请先选择哪一个工作表是「上年 BKD 明细」")
        prior, info = BKD.parse_prior_xlsx(f["data"], sheets_cfg)
        if not prior:
            return ctx.err("选中的工作表「%s」没解析到任何账户：请确认它确实是上年 BKD 明细表"
                           "（通常是 C.00 BKD），或调整该表的表头行号后重试。"
                           % sheets_cfg[0].get("sheet", ""))
        rate = float(fields.get("rate", 50)) / 100
        amt = float(fields.get("amt", 100000))
        res = BKD.merge(records, prior, rate_threshold=rate, amt_threshold=amt)

        # ---- C.03 Cutoff 截止性测试：按用户设定的阈值/截止日/前后天数抓序时账
        try:
            cutoff = _build_cutoff(conn, fields)
        except ValueError as e:
            return ctx.err(str(e))

        data, tpl_path = BKD.build_workbook(res,
                              meta={"time": time.strftime("%Y-%m-%d %H:%M"), "tbDb": BKD.KZ_DB},
                              highlight_new=fields.get("highlight", "1") != "0",
                              mark_new=fields.get("mark", "1") != "0",
                              cutoff=cutoff)
        if data is None:
            return ctx.err("未找到内置模板，请先在上方上传模板。")
        stamp = time.strftime("%Y%m%d_%H%M%S")
        name = "货币资金BKD_%s.xlsx" % stamp
        _write_out(name, data)
        return ctx.json({
            "file": name, "url": "/download?f=" + urllib.parse.quote(name),
            "tpl": os.path.basename(tpl_path or ""),
            "stat": res["stats"],
            "news": res["news"], "miss": res["miss"],
            "nonexact": [x for x in res["nonexact"] if x["_matched"]],
            "changes": res["changes"],
            "rows": res["rows"],
            "cutoff": _cutoff_payload(cutoff),
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return ctx.err("生成失败：" + str(e))

# ---------------------------------------------------------------- 账面信息表
def _book_generate(ctx, body):
    """按账套（tb）填「1.银行存款-账面模板」，产出可直接上传的 xlsx。"""
    try:
        kw = body.get("keyword") or "银行存款"
        bd = body.get("balance_date") or ""
        conn = getattr(ctx.ledger, "conn", None)
        rows, meta = BKD.read_book_rows(conn=conn, keyword=kw)
        if meta.get("error"):
            return ctx.err(meta["error"])
        if not rows:
            return ctx.err("科目余额表里没有「银行存款」账户。请确认已合并 tb，"
                           "且清洗时已把「一级科目」识别为「银行存款」。")
        data, tpl = BKD.build_book_xlsx(rows, balance_date=bd)
        if data is None:
            return ctx.err("未找到账面信息模板：data/bkd_tpl 与模块内置 _tpl 下都没有 "
                           + BKD.BOOK_TPL_NAME)
        name = "银行存款账面信息_%s.xlsx" % time.strftime("%Y%m%d_%H%M%S")
        with open(os.path.join(BKD.OUT, name), "wb") as fh:
            fh.write(data)
        return ctx.json({
            "file": name, "url": "/download?f=" + urllib.parse.quote(name),
            "tpl": os.path.basename(tpl or ""), "stat": meta, "rows": rows,
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return ctx.err("生成账面信息表失败：" + str(e))


def _book_tpl_upload(ctx, body):
    """替换账面信息模板（校验表头确实是对的那张表，避免传错文件）。"""
    f = (body.get("files") or {}).get("tpl")
    if not f:
        return ctx.err("未收到模板文件")
    fn = f["filename"] or "模板.xlsx"
    if not fn.lower().endswith((".xlsx", ".xlsm")):
        return ctx.err("模板仅支持 .xlsx / .xlsm")
    try:
        wb = BKD.load_workbook(io.BytesIO(f["data"]))
    except Exception as e:
        return ctx.err("模板无法解析：" + str(e))
    hit = None
    for sn in wb.sheetnames:
        if "账面上传" in sn or "账面信息" in sn:
            hit = sn
            break
    if not hit:
        return ctx.err("这不是「银行存款账面模板」：没找到「账面上传/账面信息」工作表。"
                       "当前工作表：" + "、".join(wb.sheetnames[:6]))
    with open(os.path.join(BKD.TPL_DIR, BKD.BOOK_TPL_NAME), "wb") as fh:
        fh.write(f["data"])
    return ctx.json({"ok": True, "filename": BKD.BOOK_TPL_NAME, "sheet": hit})


# ------------------------------------------------ 上年 WP：结构探测（AI 能力入口）
@_post("/api/build/prior/probe")
def api_build_prior_probe(ctx, u, body):
    """导出上年 WP 的「结构探测包」，交给 AI 判断哪张表是 BKD、表头在哪一行。"""
    fid = (body or {}).get("fid")
    if not fid:
        return ctx.err("未收到文件 id，请先上传上年底稿")
    import glob
    hits = glob.glob(os.path.join(BKD.KZ_DATA, "uploads", str(fid) + "_*"))
    if not hits:
        return ctx.err("找不到已上传的上年底稿（fid=%s），请重新上传" % fid)
    try:
        with open(hits[0], "rb") as fh:
            wb = BKD.load_workbook(io.BytesIO(fh.read()))
    except Exception as e:
        return ctx.err("上年底稿无法解析：" + str(e))
    try:
        out = BKD.probe_prior_structure(wb)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return ctx.json(out)


@_post("/api/build/prior/apply")
def api_build_prior_apply(ctx, u, body):
    """按 AI 给的 {sheet, header_row} 解析上年底稿，返回能回填的记录。"""
    body = body or {}
    fid = body.get("fid")
    spec = body.get("spec") or {}
    if not fid or not spec.get("sheet"):
        return ctx.err("需要上传文件 id 和 AI 返回的 {sheet, header_row}")
    import glob
    hits = glob.glob(os.path.join(BKD.KZ_DATA, "uploads", str(fid) + "_*"))
    if not hits:
        return ctx.err("找不到已上传的上年底稿，请重新上传")
    with open(hits[0], "rb") as fh:
        data = fh.read()
    cfg = [{"sheet": spec["sheet"], "header_row": int(spec.get("header_row") or 1)}]
    try:
        prior, info = BKD.parse_prior_xlsx(data, cfg)
    except Exception as e:
        return ctx.err("按该表头行解析失败：" + str(e))
    if not prior:
        # 空白模板（只有表头没有明细）也会走到这里。不是错误，如实返回让前端提示。
        return ctx.json({"rows": 0, "sum": 0, "sample": [],
                         "warn": "按第 %s 行做表头，没解析到任何账户。"
                                 "若上年底稿是空模板则正常；若有数据，说明 header_row 选错了，"
                                 "请让 AI 复核或换一张工作表。" % spec.get("header_row")})
    return ctx.json({"rows": len(prior), "info": info,
                     "sample": [{k: v for k, v in r.items() if not k.startswith("_")}
                                for r in prior[:5]],
                     "sum": round(sum(r.get("local") or 0 for r in prior), 2)})


# ------------------------------------------------ 一键生成（按科目组合产物）
@_post("/api/build/wp/run")
def api_build_wp_run(ctx, u, body):
    """一键生成：一个按钮把该科目的底稿一次产完。

    货币资金 = C.00 BKD + Lead + Cutoff + 检查表，成功后再追加 OAR；
    销管费用 = VC/VD BKD + 检查表。
    主底稿失败即整体失败（用户必须知道）；附加产物失败只记进 warnings，
    否则 OAR 缺个年度就让已经算好的 162 行 BKD 白跑一遍。
    """
    body = body or {}
    conn = getattr(ctx.ledger, "conn", None)
    subject = (body.get("subject") or "cash").strip()
    warnings = []

    if subject == "exp":
        # 上年底稿：前端两个下拉选定的 VC / VD 工作表；不传则整份读（旧行为）
        _ps = body.get("prior_sheets")
        if isinstance(_ps, str):
            _ps = [x for x in re.split(r"[,，、]", _ps) if x.strip()]
        pm, perr = _load_prior_expense(body.get("prior_fid"), sheets=_ps)
        # 上一期 TOD 就在这份上年底稿里（同一册 WP）：选了 TOD 页就用 prior_fid 取字节，
        # 不再要求用户传第二个文件；显式给了 prior_tod_fid 才另读一份（兼容旧页面）。
        _ptod_fid = body.get("prior_tod_fid") or (
            body.get("prior_fid") if body.get("prior_tod_sheet") else None)
        _ptod, _ptod_err = _load_prior_tod(_ptod_fid)
        try:
            data, info = _run_exp(conn, body.get("period_from") or "1",
                                  body.get("period_to") or "6",
                                  balance_date=body.get("balance_date", ""),
                                  customer=body.get("customer", ""),
                                  prior_map=pm or None,
                                  pre_end=body.get("pre_end"),
                                  oar_evidence=os.path.join(BKD.OUT, "OAR_evidence_费用.json"),
                                  oar_chg_min=body.get("oar_chg_min"),
                                  oar_chg_pct=float(body.get("oar_chg_pct") or 0.10),
                                  # TE/SAD 必须传：不传 Lead C4 就是空 → C22=0 →
                                  # 「进一步调查」恒为"是"，等于整张表都要查一遍
                                  te=_num_or_none(body.get("te")),
                                  sad=_num_or_none(body.get("sad")),
                                  standard=body.get("standard") or "",
                                  currency=body.get("currency") or "",
                                  assurance=body.get("assurance"),
                                  prior_fid=body.get("prior_fid"),
                                  cra=body.get("cra"),
                                  prior_tod=_ptod,
                                  prior_tod_sheet=body.get("prior_tod_sheet"))
            if _ptod_err:
                warnings.append(_ptod_err)
            if perr:
                warnings.append(perr)
                info["priorError"] = perr
            # AI 待办清单的说明（口径待判定 / 表2 / OAR 分析各一条），
            # 让界面与命令行都能看到"还有几件事没做、清单在哪"
            warnings.extend(info.get("warnings") or [])
        except Exception as e:
            import traceback
            traceback.print_exc()
            return ctx.err("生成费用底稿失败：" + str(e))
        if data is None:
            return ctx.err(info.get("error") or "生成失败")
        name = "费用BKD_%s.xlsx" % time.strftime("%Y%m%d_%H%M%S")
        with open(os.path.join(BKD.OUT, name), "wb") as fh:
            fh.write(data)
        info["file"] = name
        info["url"] = "/download?f=" + urllib.parse.quote(name)
        return ctx.json({"subject": "exp", "main": info, "oar": None,
                         "warnings": warnings})

    prior, pinfo = _load_prior(body.get("prior_fid"), body.get("prior_spec"))
    _pl = _load_prior_lead(body.get("prior_fid"),
                           (body.get("prior_spec") or {}).get("lead"))   # 上年 Lead 就在上年底稿里，同一个文件
    # 选了「上年 Lead」却没读到：让结果页直接看到，不然表1 上期数悄悄变 0
    if _pl and _pl.get("error"):
        warnings.append("上年 Lead 未使用（表1 上期数、表2 上年数会是 0）：" + _pl["error"])
    try:
        data, info = _run_c00(conn, prior=prior,
                              balance_date=body.get("balance_date", ""),
                              customer=body.get("customer", ""),
                              standard=body.get("standard", ""),
                              currency=body.get("currency", ""),
                              te=_num_or_none(body.get("te")),
                              sad=_num_or_none(body.get("sad")),
                              rates=body.get("rates"),
                              restricted=body.get("restricted"),
                              prior_lead=_pl,
                              cutoff_fields=_cutoff_fields_of(body))
    except Exception as e:
        import traceback
        traceback.print_exc()
        return ctx.err("生成货币资金底稿失败：" + str(e))
    if data is None:
        return ctx.err(info.get("error") or "生成失败")
    name = "货币资金BKD_%s.xlsx" % time.strftime("%Y%m%d_%H%M%S")
    name, latest = _write_out(name, data)
    info["file"] = name
    info["latestFile"] = latest
    if latest and latest != "%s_最新.xlsx" % name.split("_")[0]:
        info.setdefault("warnings", []).append(
            "「%s_最新.xlsx」正被 Excel 打开，已另存为 %s。请看新文件，或先关掉旧文件再生成。"
            % (name.split("_")[0], latest))
    info["url"] = "/download?f=" + urllib.parse.quote(name)
    info["prior"] = pinfo
    if info.get("cutoffError"):
        warnings.append(info["cutoffError"])
    # 补上派生表的告警：以前这里只回 [上年Lead错误, cutoffError]，C.02a 截图待补这类
    # 提示只躺在 main.buildWarnings 里，走接口的调用方根本看不到（2026-11-08 修正）。
    for _w in (info.get("buildWarnings") or []):
        if _w not in warnings:
            warnings.append(_w)

    return ctx.json({"subject": "cash", "main": info, "tasks": info.get("tasks") or [],
                     "warnings": warnings})


# ================================================================ 判断留痕
# 设计见 03_文档/V1.2.09设计_判断留痕与黄格分诊.md
# 本段只做「读 / 写 / 审」三件事，不改 _fill_bkd 的任何输出——第 1 步要求行为零变化。

try:
    from core import judge as JUDGE
except Exception:                                        # pragma: no cover
    JUDGE = None

#: 允许 AI 填的字段的取值域。不在表里的字段只校验「非空」。
_ENUMS = {"confirm": {"是", "否"}, "nature": {"受限", "非受限"}}

#: 本次 plan 发出去的 jid。apply 时校验回来的是不是我们自己发的点位，
#: 防止 AI 编造位置（无状态服务，只能靠进程内小缓存；1 小时自动失效）。
_PLAN_CACHE = {"scope": "", "ts": 0, "jids": set()}
_PLAN_TTL = 3600


def _period_of(d):
    """2026-10-31 / 202610 / 2026/10 -> '202610'。取不到返回 ''（跨期通用）。"""
    s = re.sub(r"[^0-9]", "", str(d or ""))
    return s[:6] if len(s) >= 6 else ""


def _jctx(ctx):
    conn = getattr(ctx.ledger, "conn", None)
    if JUDGE:
        JUDGE.ensure(conn)
    return conn


def _existing_index(conn, scope):
    """一次拉全量判断建索引。逐格查会变成 162 行 × 4 字段 = 650 次查询。"""
    idx = {}
    if not JUDGE or conn is None:
        return idx
    rank = {"confirmed": 0, "draft": 1, "rejected": 2}
    for r in JUDGE.list_judgments(conn, scope=scope, limit=5000):
        k = (r.get("entity"), r.get("dimKey"), r.get("field"))
        old = idx.get(k)
        if old is None or rank.get(r.get("status"), 9) < rank.get(old.get("status"), 9):
            idx[k] = r
    return idx


@_get("/api/judge/plan")
def api_judge_plan(ctx, q):
    """待判清单：跑一次 C.00 预览，把黄格子分诊，C 类列成可交给 AI 起草的条目。

    这里实时跑 _run_c00(preview=True) 拿 gaps，而不是读缓存——但 preview 不落盘，
    所以对现有生成结果零影响。
    """
    if not JUDGE:
        return ctx.json({"error": "判断留痕模块未加载"})
    conn = _jctx(ctx)
    scope = (q.get("scope") or ["bkd.c00"])[0]
    prior, pinfo = [], {"loaded": False}
    fid = (q.get("prior_fid") or [""])[0]
    if fid:
        try:
            spec = json.loads((q.get("prior_spec") or [""])[0] or "null")
        except Exception:
            spec = None
        prior, pinfo = _load_prior(fid, spec)
    try:
        _data, info = _run_c00(conn, prior=prior,
                               balance_date=(q.get("balance_date") or [""])[0],
                               customer=(q.get("customer") or [""])[0], preview=True)
    except Exception as e:
        return ctx.json({"error": "预览失败：" + str(e)})
    if info.get("error"):
        return ctx.json({"error": info["error"]})

    period = _period_of((q.get("balance_date") or [""])[0])
    idx = _existing_index(conn, scope)
    gaps = []
    for g in (info.get("gapList") or []):
        gg = dict(g)
        gg["field_key"] = JUDGE.field_key(g.get("field"))
        hit = idx.get((g.get("company"), JUDGE._norm_dim("acct", g.get("acct")),
                       gg["field_key"]))
        # 只有 confirmed 才算「已有结论」（A 类）。AI 起草是 draft，必须继续走 C 类：
        # 起草只是给人看的草案，判成 A 类等于让 AI 静默落笔。
        if hit and hit.get("status") == "confirmed":
            gg["resolved"] = True
            gg["source"] = {"human": "人工", "prior_judgment": "往期继承"}.get(
                hit.get("source"), hit.get("source"))
            gg["jid"] = hit.get("jid")
        elif hit and hit.get("status") == "draft":
            gg["hasDraft"] = True
        gaps.append(gg)
    cls_gaps, cnt = JUDGE.classify_all(gaps)

    items, sent = [], set()
    for g in cls_gaps:
        if g.get("class") != "C" or not g.get("field_key"):
            continue
        j = JUDGE.jid(scope, g.get("company"), g.get("subject"), "acct",
                      g.get("acct"), period, g["field_key"])
        if j in sent:
            continue
        sent.add(j)
        hit = idx.get((g.get("company"), JUDGE._norm_dim("acct", g.get("acct")),
                       g["field_key"]))
        items.append({
            "jid": j, "entity": g.get("company"), "subject": g.get("subject"),
            "dimType": "acct", "dimKey": JUDGE._norm_dim("acct", g.get("acct")),
            "period": period, "field": g["field_key"], "fieldLabel": g.get("field"),
            "row": g.get("row"), "col": g.get("col"), "reason": g.get("reason"),
            "hint": g.get("hint"),
            "enum": sorted(_ENUMS.get(g["field_key"]) or []),
            "draft": ({"value": hit.get("value"), "basis": hit.get("basis"),
                       "confidence": hit.get("confidence")} if hit
                      and hit.get("status") == "draft" else None),
        })

    import time as _t
    _PLAN_CACHE.update(scope=scope, ts=_t.time(), jids={i["jid"] for i in items})

    # B 类汇总成「去要什么」（只预览，不落库；落库走 /api/judge/todos/sync）
    todo_items = _todo_items(cls_gaps)
    return ctx.json({
        "scope": scope, "period": period, "rows": info.get("rows"),
        "gaps": len(info.get("gapList") or []), "counts": cnt,
        "items": items, "count": len(items),
        "todos": todo_items, "todoCount": len(todo_items),
        "prior": pinfo,
    })


def _todo_items(cls_gaps):
    """B 类按 (主体, 账户, 字段) 合并——客户回一封邮件能补 8 个账户，清单要 1 条。"""
    m = {}
    for g in cls_gaps or []:
        if g.get("class") != "B":
            continue
        dk = JUDGE._norm_dim("acct", g.get("acct"))
        k = (g.get("company"), dk, g.get("field_key"))
        d = m.get(k)
        if d:
            d["affected"] += 1
            continue
        m[k] = {
            "entity": g.get("company"), "dim_type": "acct", "dim_key": dk,
            "field_key": g.get("field_key"),
            "title": "%s · %s · %s" % (g.get("company") or "(未标注主体)",
                                       g.get("acct") or "无账号", g.get("field") or ""),
            "why": g.get("reason") or "", "ask": JUDGE.ASK_TEXT.get(g.get("field_key"), ""),
            "affected": 1,
        }
    return list(m.values())


@_post("/api/judge/apply")
def api_judge_apply(ctx, u, body):
    """收 AI 起草结果。四条校验，不合格整条打回——不静默丢弃，
    否则 AI 会以为自己写进去了。"""
    if not JUDGE:
        return ctx.err("判断留痕模块未加载")
    conn = _jctx(ctx)
    scope = body.get("scope") or "bkd.c00"
    items = body.get("items") or []
    if not isinstance(items, list) or not items:
        return ctx.err("items 为空")
    import time as _t
    fresh = bool(_PLAN_CACHE["jids"]) and (_t.time() - _PLAN_CACHE["ts"] < _PLAN_TTL) \
        and _PLAN_CACHE["scope"] == scope
    saved, skipped = 0, []
    for i, it in enumerate(items, 1):
        j = str(it.get("jid") or "")
        val = str(it.get("value") or "").strip()
        basis = str(it.get("basis") or "").strip()
        if len(j) != 16 or not re.match(r"^[0-9a-f]{16}$", j):
            skipped.append({"no": i, "jid": j, "reason": "jid 格式不对"})
            continue
        if fresh and j not in _PLAN_CACHE["jids"]:
            skipped.append({"no": i, "jid": j, "reason": "不在本批 plan 内（勿编造位置）"})
            continue
        fk = it.get("field") or ""
        en = _ENUMS.get(fk)
        if en and val not in en:
            skipped.append({"no": i, "jid": j,
                            "reason": "取值「%s」不在允许范围 %s 内" % (val, "/".join(sorted(en)))})
            continue
        if not val:
            skipped.append({"no": i, "jid": j, "reason": "值为空"})
            continue
        if len(basis) < 8:
            skipped.append({"no": i, "jid": j, "reason": "依据太短（须 ≥8 字，写清出处）"})
            continue
        try:
            conf = float(it.get("confidence")) if it.get("confidence") is not None else None
        except (TypeError, ValueError):
            conf = None
        if conf is not None and conf < 0.6:
            skipped.append({"no": i, "jid": j,
                            "reason": "置信度 %.2f 低于 0.6，证据不足不写值" % conf})
            continue
        _j, _st, err = JUDGE.upsert(
            conn, scope, it.get("entity"), it.get("subject"), it.get("dim_type") or "acct",
            it.get("dim_key"), it.get("period") or "", fk, val, source="ai",
            basis=basis, confidence=conf, status="draft",
            run_id=str(body.get("run_id") or ""), created_by=str(body.get("by") or "AI"))
        if err:
            skipped.append({"no": i, "jid": j, "reason": err})
        else:
            saved += 1
    return ctx.json({"saved": saved, "skipped": skipped, "total": len(items),
                     "note": "AI 起草的值须人工确认后才算数；未确认前底稿仍留空（BKD 明细区不涂色）。"})


@_post("/api/judge/save")
def api_judge_save(ctx, u, body):
    """人工直接写一条（填完即确认）。这是唯一的 confirmed 入口之一。"""
    if not JUDGE:
        return ctx.err("判断留痕模块未加载")
    conn = _jctx(ctx)
    j, _st, err = JUDGE.upsert(
        conn, body.get("scope") or "bkd.c00", body.get("entity"), body.get("subject"),
        body.get("dim_type") or "acct", body.get("dim_key"),
        body.get("period") or "", body.get("field") or "",
        body.get("value"), source="human", basis=body.get("basis") or "",
        status="confirmed", created_by=body.get("by") or "")
    if err:
        return ctx.err(err)
    return ctx.json({"ok": True, "jid": j})


@_post("/api/judge/review")
def api_judge_review(ctx, u, body):
    """复核：确认 / 驳回。AI 起草必须过这一步。"""
    if not JUDGE:
        return ctx.err("判断留痕模块未加载")
    conn = _jctx(ctx)
    j = body.get("jid") or ""
    act = body.get("action") or "confirm"
    by = body.get("by") or ""
    if act == "confirm":
        ok, err = JUDGE.confirm(conn, j, by=by, value=body.get("value"),
                                basis=body.get("basis"))
    elif act == "reject":
        ok, err = JUDGE.reject(conn, j, by=by, why=body.get("why") or "")
    else:
        return ctx.err("未知动作：%s" % act)
    return ctx.err(err) if not ok else ctx.json({"ok": True})


@_get("/api/judge/list")
def api_judge_list(ctx, q):
    if not JUDGE:
        return ctx.json({"items": [], "stats": {}})
    conn = _jctx(ctx)
    return ctx.json({
        "items": JUDGE.list_judgments(
            conn, scope=(q.get("scope") or [None])[0],
            status=(q.get("status") or [None])[0],
            entity=(q.get("entity") or [None])[0],
            limit=int((q.get("limit") or ["500"])[0])),
        "stats": JUDGE.stats(conn),
    })


@_get("/api/judge/todos")
def api_judge_todos(ctx, q):
    if not JUDGE:
        return ctx.json({"items": [], "stats": {}})
    conn = _jctx(ctx)
    return ctx.json({"items": JUDGE.list_todos(conn, status=(q.get("status") or [None])[0]),
                     "stats": JUDGE.todo_stats(conn)})


@_post("/api/judge/todos/sync")
def api_judge_todos_sync(ctx, u, body):
    """把 B 类缺料落成待办。由用户点「生成追问清单」触发，预览时不写库。"""
    if not JUDGE:
        return ctx.err("判断留痕模块未加载")
    conn = _jctx(ctx)
    r = JUDGE.todo_sync(conn, body.get("items") or [], scope=body.get("scope") or "bkd.c00")
    return ctx.json(r)


@_post("/api/judge/todo")
def api_judge_todo(ctx, u, body):
    if not JUDGE:
        return ctx.err("判断留痕模块未加载")
    conn = _jctx(ctx)
    ok, err = JUDGE.todo_set(conn, body.get("tid") or "", body.get("status") or "open",
                             note=body.get("note") or "")
    return ctx.err(err) if not ok else ctx.json({"ok": True})


# ================================================================ 长期资产（固定资产 SWP）
# V1.2.98：旧的 K1_100Breakdown 链路（asset.py + 资产台账 + 分类对照表）整条下线，
# 长期资产页直接跑 SWP 引擎：数据从科目余额表 + 序时账来，不再要求客户另给 FA list。
# 引擎在 modules/fa_swp/ 里，这里只做「读预览 / 生成 / 出下载链接」。
FA_SWP_TPL = "K1 SWP 固定资产 202YMMDD XYZ公司.xlsx"


def _fa_swp_tpl():
    return os.path.join(BKD.TPL_DIR, FA_SWP_TPL)


@_get("/api/build/fa_swp/plan")
def api_build_fa_swp_plan(ctx, q):
    """固定资产 SWP 取数预览：类别矩阵、TB 重复导入去重、清单与账的类别互校。

    只看不写，不生成任何文件。点「生成」才真正跑 synth（要造明细，几秒钟）。
    """
    from modules.fa_swp import checks as FC, moves as FM, prep as FP

    conn = getattr(ctx.ledger, "conn", None)
    if conn is None:
        return ctx.json({"error": "还没有打开账套"})
    if not os.path.isfile(_fa_swp_tpl()):
        return ctx.json({
            "error": "找不到模板「%s」。请把它放回 data/bkd_tpl/ 再试。" % FA_SWP_TPL,
            "hasData": False})

    try:
        FP.init(BKD.KZ_DATA)
        p = FP.load(conn)
        m = FM.build(conn, p)
    except Exception as e:                                  # noqa: BLE001
        import traceback
        traceback.print_exc()
        return ctx.json({"error": "读取固定资产数据失败：" + str(e), "hasData": False})

    tot = {k: float(v) for k, v in p["totals"].items()}
    if not any(abs(v) > 0.005 for v in tot.values()):
        return ctx.json({
            "error": "账套里没有固定资产余额（1601/1602/1603 全为 0）。"
                     "先在「一、导入与清洗」导入科目余额表与序时账，再回来。",
            "hasData": False})

    union = FC.check_cat_union(p, m)
    cats = []
    for c in p["cat_order"]:
        mm = p["cats"][c]
        cats.append({
            "cat": c,
            "faBeg": float(mm["fa_beg"]), "faEnd": float(mm["fa_end"]),
            "depEnd": float(mm["dep_end"]), "impEnd": float(mm["imp_end"]),
            "nbvEnd": float(mm["fa_end"] - mm["dep_end"] - mm["imp_end"]),
        })

    # m["totals"] 的键是 (段, 方式) 元组，不是 "fa|购置" 那种字符串（那套写法只在
    # build 的 trace 里）。按元组取，再按 WAY_ORDER 的规范顺序排，页面上的行序才稳定。
    label_of = {"fa": "原值", "dep": "折旧", "imp": "减值"}
    order_idx = {(sec, w): i for sec in FM.WAY_ORDER
                 for i, w in enumerate(FM.WAY_ORDER[sec])}
    ways = [{"sec": k[0], "label": label_of.get(k[0], k[0]), "way": k[1],
             "amt": float(v)}
            for k, v in m["totals"].items()
            if isinstance(k, tuple) and len(k) == 2]
    ways.sort(key=lambda x: (list(label_of).index(x["sec"])
                             if x["sec"] in label_of else 9,
                             order_idx.get((x["sec"], x["way"]), 99)))

    t4 = m["table4"]
    return ctx.json({
        "hasData": True,
        "tpl": {"file": FA_SWP_TPL, "exists": True},
        "rows": {"tb": len(p["rows"]), "journal": len(m["lines"])},
        "dedupe": {
            "n": len(p["trace"]["dropped"]),
            "sample": [{"company": d.get("company"), "acct": d.get("acct"),
                        "srcRow": d.get("src_row"), "dupOf": d.get("dup_of")}
                       for d in (p["trace"]["dropped"] or [])[:8]],
            "adjusts": p["trace"]["adjusts"],
        },
        "catUnion": {"ok": bool(union.get("ok")),
                     "onlyTb": union.get("only_tb") or [],
                     "onlyJnl": union.get("only_jnl") or [],
                     "message": union.get("message") or ""},
        "cats": cats,
        "totals": tot,
        "net": {"beg": float(p["net_beg"]), "end": float(p["net_end"])},
        "ways": ways,
        "table4": {"total": float(t4["total"]), "diff": float(t4["diff"]),
                   "buckets": {k: float(v) for k, v in t4["buckets"].items()}},
    })


@_post("/api/build/fa_swp/run")
def api_build_fa_swp_run(ctx, u, body):
    """生成固定资产 SWP 底稿（K.01 + Lead + 三张清单，17 张 sheet）。"""
    from modules.fa_swp import build as FB

    body = body or {}
    conn = getattr(ctx.ledger, "conn", None)
    if conn is None:
        return ctx.err("还没有打开账套")

    ds = str(body.get("balance_date") or "").strip()
    try:
        bd = _date.fromisoformat(ds)
    except Exception:                                       # noqa: BLE001
        return ctx.err("资产负债表日请填 YYYY-MM-DD（现在是「%s」）" % ds)
    if not os.path.isfile(_fa_swp_tpl()):
        return ctx.err("找不到模板「%s」。请把它放回 data/bkd_tpl/ 再试。" % FA_SWP_TPL)

    info = {
        "customer": body.get("customer") or "",
        "balance_date": ds,
        "te": float(body.get("te") or 0),
        "sad": float(body.get("sad") or 0),
        "standard": body.get("standard") or "中国企业会计准则",
        "currency": body.get("currency") or "人民币元",
    }
    t0 = time.time()
    try:
        res = FB.build(conn, BKD.KZ_DATA, BKD.OUT, info, bd, body.get("seed") or "ptw")
    except FileNotFoundError as e:
        return ctx.err(str(e))
    except Exception as e:                                  # noqa: BLE001
        import traceback
        traceback.print_exc()
        return ctx.err("生成固定资产 SWP 底稿失败：" + str(e))

    tr = res["trace"]
    name = os.path.basename(res["path"])
    latest = os.path.basename(res["latest"])
    out = {
        "ok": True, "target": "fa",
        "file": name, "latestFile": latest,
        "prior": os.path.basename(res["prior"]),
        "url": "/download?f=" + urllib.parse.quote(name),
        "latestUrl": "/download?f=" + urllib.parse.quote(latest),
        "elapsed": round(time.time() - t0, 1),
        "catOrder": tr["cat_order"], "cats": len(tr["cat_order"]),
        "totals": tr["totals"], "net": tr["net"],
        "k01": tr["k01"], "table4": tr["table4"],
        "rowsInserted": tr["rows_inserted"], "rowPlan": tr["row_plan"],
        "blockPlan": tr["block_plan"],
        "tbDup": tr["tb_dup"], "tbRows": tr["tb_rows"],
        "checkItems": tr["check_items"],
        "hidden": tr["hidden"],
    }
    fails = (res.get("checks") or {}).get("fails") or []
    out["checkFails"] = [f.get("name") for f in fails]
    if latest != "固定资产SWP_最新.xlsx":
        out["warnings"] = ["「固定资产SWP_最新.xlsx」正被 Excel 打开，已另存为 %s。"
                           "请看新文件，或先关掉旧文件再生成。" % latest]
    return ctx.json(out)
