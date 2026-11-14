# -*- coding: utf-8 -*-
"""导入与清洗板块 · 路由处理

本板块负责：导入原始数据、数据清洗、字段映射、合并成表、账簿对照表、
清洗方案库、数据检查前的准备等。所有处理函数接收 (ctx) 上下文，ctx 提供
handler 的 _json / _err / ledger / web 等能力。

路由采用「返回 (handled, result)」约定：handled=True 表示已处理并返回，
result 为响应对象；handled=False 表示本模块不处理该路径。
"""
import urllib.parse

# GET 路由：path -> handler(ctx, query_dict)
GET_ROUTES = {}
# POST 路由：path -> handler(ctx, parsed_url, body_dict)
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


# ---------------------------------------------------------------- GET
@_get("/api/state")
def api_state(ctx, q):
    return ctx.json(ctx.ledger.state())


@_get("/api/clean/types")
def api_clean_types(ctx, q):
    return ctx.json(ctx.ledger.clean_types())


@_get("/api/books/list")
def api_books_list(ctx, q):
    return ctx.json(ctx.ledger.books_list(int(q.get("limit", ["500"])[0])))


@_get("/api/recipes")
def api_recipes(ctx, q):
    return ctx.json(ctx.ledger.recipes_list())


# ---------------------------------------------------------------- POST
@_post("/api/upload")
def api_upload(ctx, u, body):
    name = urllib.parse.unquote(ctx.headers.get("X-Filename", "upload.dat"))
    size = int(ctx.headers.get("Content-Length") or 0)
    info = ctx.ledger.add_file(name, ctx.rfile, size)
    return ctx.json(info)


@_post("/api/inspect")
def api_inspect(ctx, u, body):
    return ctx.json(ctx.ledger.inspect(body["fid"], body.get("target")))


@_post("/api/zip/expand")
def api_zip_expand(ctx, u, body):
    """客户 ERP 导出的压缩包（包内是账表）在服务端解开，逐个登记成待导入文件。"""
    return ctx.json(ctx.ledger.expand_zip(body["fid"]))


@_post("/api/import")
def api_import(ctx, u, body):
    return ctx.json(ctx.ledger.import_sheet(
        body["fid"], body["sheet"], body.get("header_row", 1),
        body.get("header_rows", 1), body.get("target"), body.get("mode", "replace")))


@_get("/api/import/progress")
def api_import_progress(ctx, q):
    """导入进度轮询：import 请求进行中由前端定时调用，拿百分比和阶段。"""
    return ctx.json(ctx.ledger.prog_get((q.get("key") or [""])[0]))


@_post("/api/preview")
def api_preview(ctx, u, body):
    return ctx.json(ctx.ledger.raw_preview(body["sid"], body.get("limit", 100),
                                           body.get("stage", "data")))


@_post("/api/mapping")
def api_mapping(ctx, u, body):
    return ctx.json(ctx.ledger.set_mapping(body["sid"], body.get("target"), body.get("mapping")))


@_post("/api/suggest")
def api_suggest(ctx, u, body):
    return ctx.json(ctx.ledger.suggest(body["sid"], body.get("target")))


@_post("/api/aux/detect")
def api_aux_detect(ctx, u, body):
    return ctx.json(ctx.ledger.detect_roles(body["sid"], body.get("mapping"), body.get("col")))


@_post("/api/clean/save")
def api_clean_save(ctx, u, body):
    return ctx.json(ctx.ledger.clean_save(body["sid"], body.get("steps")))


@_post("/api/clean/preview")
def api_clean_preview(ctx, u, body):
    return ctx.json(ctx.ledger.clean_preview(body["sid"], body.get("upto"),
                                             body.get("sample", 30)))


@_post("/api/clean/apply")
def api_clean_apply(ctx, u, body):
    return ctx.json(ctx.ledger.clean_apply(body["sid"]))


@_post("/api/clean/apply-all")
def api_clean_apply_all(ctx, u, body):
    return ctx.json(ctx.ledger.clean_apply_all((body or {}).get("sids")))


@_post("/api/clean/reset")
def api_clean_reset(ctx, u, body):
    return ctx.json(ctx.ledger.clean_reset(body["sid"]))


@_post("/api/clean/preset")
def api_clean_preset(ctx, u, body):
    return ctx.json(ctx.ledger.clean_preset(body["sid"], body.get("key", "generic")))


@_post("/api/clean/dropped")
def api_clean_dropped(ctx, u, body):
    return ctx.json(ctx.ledger.clean_dropped(body["sid"], body.get("limit", 200)))


@_post("/api/clean/recipe/export")
def api_clean_recipe_export(ctx, u, body):
    r = ctx.ledger.clean_recipe_export()
    r["url"] = "/download?f=" + urllib.parse.quote(r["file"])
    return ctx.json(r)


@_post("/api/clean/recipe/parse")
def api_clean_recipe_parse(ctx, u, body):
    return ctx.json(ctx.ledger.clean_recipe_parse(body["fid"]))


@_post("/api/clean/recipe/apply")
def api_clean_recipe_apply(ctx, u, body):
    return ctx.json(ctx.ledger.clean_recipe_apply(
        body["sid"], body.get("steps"), body.get("mapping"), body.get("target")))


@_post("/api/recipes/save")
def api_recipes_save(ctx, u, body):
    return ctx.json(ctx.ledger.recipes_save(body.get("name"), body.get("steps"),
                                            body.get("sid"), body.get("note") or ""))


@_post("/api/recipes/apply")
def api_recipes_apply(ctx, u, body):
    return ctx.json(ctx.ledger.recipes_apply(body["name"], body["sid"]))


@_post("/api/recipes/save-many")
def api_recipes_save_many(ctx, u, body):
    return ctx.json(ctx.ledger.recipes_save_many((body or {}).get("plans")))


@_post("/api/recipes/delete")
def api_recipes_delete(ctx, u, body):
    return ctx.json(ctx.ledger.recipes_delete(body["name"]))


@_post("/api/recipes/rename")
def api_recipes_rename(ctx, u, body):
    return ctx.json(ctx.ledger.recipes_rename(body["name"], body.get("new_name")))


@_post("/api/books/preview")
def api_books_preview(ctx, u, body):
    return ctx.json(ctx.ledger.books_preview(body["fid"]))


@_post("/api/books/import")
def api_books_import(ctx, u, body):
    return ctx.json(ctx.ledger.books_import(
        body["fid"], body["sheet"], body.get("header_row", 1),
        body.get("book_col"), body.get("company_col"), body.get("short_col")))


@_post("/api/books/clear")
def api_books_clear(ctx, u, body):
    return ctx.json(ctx.ledger.books_clear())


@_post("/api/drop")
def api_drop(ctx, u, body):
    """删原始表。body.sid 单张，body.sids 批量（多选删除走这条）。"""
    sids = body.get("sids")
    if isinstance(sids, list) and sids:
        return ctx.json(ctx.ledger.drop_sheets(sids))
    return ctx.json(ctx.ledger.drop_sheet(body["sid"]))


@_post("/api/detect")
def api_detect(ctx, u, body):
    """自动识别表类型：给一组「文件名 / 工作表名 / 表头列」，返回猜的 target 与依据。

    前端在「待导入」分组卡上先填好下拉框，人只负责复核——
    猜不准时 target 返回空串，宁可让人选，也不硬塞一个错的。
    """
    from core import fields as F
    items = body.get("items") or []
    out = []
    for it in items:
        r = F.detect_target(it.get("file_name") or "", it.get("sheet_name") or "",
                            it.get("columns") or [])
        r["key"] = it.get("key") or ""
        out.append(r)
    return ctx.json({"items": out})


@_post("/api/merge")
def api_merge(ctx, u, body):
    """合并成标准表（带清洗达标门禁）。

    清洗不达标时不落库、直接回滚，返回 409 + 不达标清单；
    body.force=true 可越过（不达标项会写进 warnings 留痕）。
    """
    body = body or {}
    from core.db import MergeBlocked
    try:
        return ctx.json(ctx.ledger.merge(body.get("target"), force=bool(body.get("force"))))
    except MergeBlocked as e:
        return ctx.json({
            "ok": False, "blocked": True, "target": e.target, "items": e.items,
            "verdict": "清洗未达标，已阻止合并（%d 项）：%s"
                       % (len(e.items), "；".join(i.get("title") or "" for i in e.items)),
            "hint": "按下面每项的「怎么修」改清洗动作后重新合并；"
                    "确要带病合并，加 force=true。",
        }, 409)


# ------------------------------------------------ 银行流水：账户清单与配对
@_post("/api/flow/accounts")
def api_flow_accounts(ctx, u, body):
    """从科目余额表重建账户清单。

    清单是银行流水配对的依据，补完上年账号或人工补录后可以反复重建；
    重建后调 /api/flow/link 重新配对，不用重导流水。
    """
    from core import accounts as ACC
    conn = getattr(ctx.ledger, "conn", None)
    try:
        stat = ACC.build(conn)
    except Exception as e:
        return ctx.err("重建账户清单失败：%s" % e)
    if stat.get("error"):
        return ctx.err(stat["error"])
    stat["list"] = ACC.list_accounts(conn, limit=500)
    return ctx.json(stat)


@_get("/api/flow/accounts")
def api_flow_accounts_list(ctx, q):
    from core import accounts as ACC
    conn = getattr(ctx.ledger, "conn", None)
    limit = int((q.get("limit") or ["500"])[0])
    return ctx.json({"list": ACC.list_accounts(conn, limit=limit)})


@_post("/api/flow/link")
def api_flow_link(ctx, u, body):
    """重新做一次流水 ↔ 账户清单的配对（全表重算）。"""
    from core import accounts as ACC
    conn = getattr(ctx.ledger, "conn", None)
    try:
        res = ACC.link(conn, only_missing=False)
    except Exception as e:
        return ctx.err("配对失败：%s" % e)
    if res.get("error"):
        return ctx.err(res["error"])
    res["gaps"] = ACC.gap_list(conn, limit=200)
    return ctx.json(res)


@_get("/api/flow/gaps")
def api_flow_gaps(ctx, q):
    from core import accounts as ACC
    conn = getattr(ctx.ledger, "conn", None)
    limit = int((q.get("limit") or ["200"])[0])
    return ctx.json({"list": ACC.gap_list(conn, limit=limit)})


# ------------------------------------------------ 清洗达标自检
@_get("/api/clean/audit")
def api_clean_audit(ctx, q):
    """判定清洗后的数据是否达到「完整、最细颗粒度」。

    只读，不改任何数据。返回逐项结论 + 给 AI 的提示词。
    """
    from core import audit as AUD
    conn = getattr(ctx.ledger, "conn", None)
    try:
        return ctx.json(AUD.audit(conn))
    except Exception as e:
        return ctx.err("自检失败：%s" % e)
