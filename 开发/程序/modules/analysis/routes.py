# -*- coding: utf-8 -*-
"""查账与分析板块 · 路由处理

本板块负责：数据检查（核对）、查账分析（SQL 查询/导出）、分享导出/导入、
规则管理、备份。所有处理函数接收 (ctx) 上下文。
"""
import urllib.parse

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


# ---------------------------------------------------------------- GET
@_get("/api/rules")
def api_rules_get(ctx, q):
    return ctx.json(ctx.ledger.get_rules())


@_get("/api/recon/last")
def api_recon_last(ctx, q):
    return ctx.json(ctx.ledger.recon_last() or {"none": True})


@_get("/api/schema")
def api_schema(ctx, q):
    return ctx.json(ctx.ledger.schema())


@_get("/api/ping")
def api_ping(ctx, q):
    import os
    return ctx.json({"ok": True, "pid": os.getpid()})


# ---------------------------------------------------------------- POST
@_post("/api/recon")
def api_recon(ctx, u, body):
    return ctx.json(ctx.ledger.recon(
        body.get("period_from") or None,
        body.get("period_to") or None, body.get("tol", 0.01),
        body.get("tb_base") or "ytd"))


@_post("/api/recon/export")
def api_recon_export(ctx, u, body):
    r = ctx.ledger.recon_export(body.get("name"))
    r["url"] = "/download?f=" + urllib.parse.quote(r["file"])
    return ctx.json(r)


@_get("/api/books/matches")
def api_book_matches(ctx, q):
    """账簿名推断留痕清单（只列对照表没覆盖、靠推断补上的）。"""
    only = str((q or {}).get("only_confirm") or "") in ("1", "true", "yes")
    return ctx.json(ctx.ledger.book_matches(only_confirm=only))


@_post("/api/books/matches/export")
def api_book_matches_export(ctx, u, body):
    r = ctx.ledger.book_matches_export((body or {}).get("name"))
    r["url"] = "/download?f=" + urllib.parse.quote(r["file"])
    return ctx.json(r)


@_post("/api/diagnose")
def api_diagnose(ctx, u, body):
    return ctx.json(ctx.ledger.diagnose())


@_post("/api/delete-rows")
def api_delete_rows(ctx, u, body):
    """按列值删行。conditions 里多条之间是 AND；dry_run 只试算不删。"""
    body = body or {}
    return ctx.json(ctx.ledger.delete_rows(
        body.get("table"),
        body.get("col"),
        body.get("mode", "等于"),
        body.get("value", ""),
        conditions=body.get("conditions"),
        dry_run=bool(body.get("dry_run"))))


@_get("/api/delete-rows/modes")
def api_delete_modes(ctx, q):
    """删除条件可用的模式 + 每张标准表的可选列（中文名）。"""
    from core.fields import table_columns
    return ctx.json({
        "modes": list(ctx.ledger.DEL_MODES),
        "tables": {t: [{"key": k, "label": lb} for k, lb in table_columns(t)]
                   for t in ("journal", "tb")},
    })


@_post("/api/query")
def api_query(ctx, u, body):
    return ctx.json(ctx.ledger.query(body["sql"], body.get("limit", 2000), body.get("params")))


@_post("/api/interrupt")
def api_interrupt(ctx, u, body):
    return ctx.json(ctx.ledger.interrupt())


@_post("/api/count")
def api_count(ctx, u, body):
    return ctx.json(ctx.ledger.count(body["sql"], body.get("params")))


@_post("/api/explain")
def api_explain(ctx, u, body):
    return ctx.json({"plan": ctx.ledger.explain(body["sql"], body.get("params"))})


@_post("/api/export")
def api_export(ctx, u, body):
    r = ctx.ledger.export(body["sql"], body.get("fmt", "csv"), body.get("name"), body.get("params"))
    r["url"] = "/download?f=" + urllib.parse.quote(r["file"])
    return ctx.json(r)


@_post("/api/rules")
def api_rules_save(ctx, u, body):
    return ctx.json(ctx.ledger.save_rules(body))


@_post("/api/rules/export")
def api_rules_export(ctx, u, body):
    r = ctx.ledger.export_rules_xlsx()
    r["url"] = "/download?f=" + urllib.parse.quote(r["file"])
    return ctx.json(r)


@_post("/api/rules/import")
def api_rules_import(ctx, u, body):
    return ctx.json(ctx.ledger.import_rules_xlsx(body["fid"], body.get("mode", "append")))


@_post("/api/share/export")
def api_share_export(ctx, u, body):
    r = ctx.ledger.share_export(note=body.get("note", ""))
    r["url"] = "/download?f=" + urllib.parse.quote(r["file"])
    return ctx.json(r)


@_post("/api/share/import")
def api_share_import(ctx, u, body):
    return ctx.json(ctx.ledger.share_import(body["fid"], body.get("mode", "replace")))


@_post("/api/backup")
def api_backup(ctx, u, body):
    r = ctx.ledger.backup()
    r["url"] = "/download?f=" + urllib.parse.quote(r["file"])
    return ctx.json(r)
