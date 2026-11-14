# -*- coding: utf-8 -*-
"""固定资产 SWP · 编排。

一条链路：

    duckdb ──prep──> 目标矩阵 ──moves──> K.01 行项目
                 │                        │
                 ├──prior──> 模拟上年 K.01 ┘
                 └──synth──> 三张清单 ──> tables 写入
                                        └──> k01 四张表 ──> lead

产物落在 ``<data>/bkd_out/``，同时留一份 ``_最新.xlsx``。
"""

from __future__ import annotations

import datetime as dt
import json
import os
import shutil
import tempfile

import openpyxl

from openpyxl.utils import get_column_letter

from . import (checks, embed, fmt, k01, lead, moves, prep, prior, synth, tables,
               xlcols, xlrows, xlshrink)

TPL_NAME = "K1 SWP 固定资产 202YMMDD XYZ公司.xlsx"


def tpl_path(data_dir: str) -> str:
    """模板只在 data/bkd_tpl 里找；找不到就报错列路径，不静默换模板。"""
    p = os.path.join(data_dir, "bkd_tpl", TPL_NAME)
    if os.path.isfile(p):
        return p
    cands = []
    d = os.path.join(data_dir, "bkd_tpl")
    if os.path.isdir(d):
        cands = [f for f in os.listdir(d) if f.lower().endswith(".xlsx")]
    raise FileNotFoundError("找不到固定资产 SWP 模板：%s；bkd_tpl 里有 %s"
                            % (p, cands or "（空）"))


def _sheet(wb, name, like=None):
    if name in wb.sheetnames:
        return wb[name]
    for s in wb.sheetnames:
        if like and like in s:
            return wb[s]
    raise KeyError("工作簿里没有 sheet %r，现有 %s" % (name, wb.sheetnames))


def build(conn, data_dir: str, out_dir: str, info: dict,
          balance_date: dt.date, seed: str = "ruiyun") -> dict:
    """跑完整流程，返回结果包（含产物路径、校验数据、留痕）。"""
    # 0) 配置：类别优先序、伪类别名单、减值准备归哪类、造明细的兜底参数
    prep.init(data_dir)

    # 1) 数据
    p = prep.load(conn)
    m = moves.build(conn, p)

    # 1.5) 类别并集互校：TB 侧与序时账侧的类别集合必须一致。
    # K.01 的年初来自 TB、变动来自序时账，两边分叉时期末必然对不上，
    # 但表现只是「差异若干元」，看不出根因 —— 这里提前拦住。
    u = checks.check_cat_union(p, m)
    if not u["ok"]:
        raise ValueError("类别口径不一致，不能出底稿：%s。TB 侧证据：%s；序时账侧证据：%s"
                         % (u["message"], u["rows_tb"], u["rows_jnl"]))

    # 2) 上年 K.01（模拟）
    # 文件名里的客户取项目信息里的客户名，不写死某一家（原来硬编码「瑞云」，换客户会带错名）。
    prior_end = dt.date(balance_date.year - 1, 12, 31)
    _cust = (info.get("customer") or "").strip() or "客户"
    prior_path = os.path.join(out_dir, "上年K.01_%s_%s.xlsx" % (_cust, prior_end.strftime("%Y%m%d")))
    prior.build(p["cats"], info.get("customer", ""), str(prior_end), prior_path,
                order=p["cat_order"])
    prior_cats = prior.read(prior_path, p["cat_order"])

    # 期初必须与 TB 期初逐格一致，否则整张 K.01 都是错的
    for c in p["cat_order"]:
        for f, key in (("fa", "fa_beg"), ("dep", "dep_beg"), ("imp", "imp_beg")):
            if abs(prior_cats[c][f] - float(p["cats"][c][key])) > 0.005:
                raise ValueError("上年 K.01 的 %s %s 与本年 TB 期初不一致：%.2f vs %.2f"
                                 % (c, f, prior_cats[c][f], float(p["cats"][c][key])))

    # 3) 三张清单
    # 客户台账已导入（fa/fa_beg 非空）就照搬真卡片，否则按序时账反推明细。
    cards = synth.load_cards(conn)
    syn = synth.build(p, m, balance_date, seed, cards)

    # 3.5) 数据层校验：不过就不生成文件（宁可不出，也不出一份看着平的错底稿）
    chk = checks.check(p, m, syn)
    chk_prior = checks.check_prior(p, prior_cats)
    if not chk["ok"] or not chk_prior["ok"]:
        raise ValueError("数据层校验未通过：\n%s\n%s"
                         % (checks.summary(chk), "；".join(chk_prior["fails"])))

    # 4) 写模板
    tpl = tpl_path(data_dir)
    wb = openpyxl.load_workbook(tpl)

    # 4.1) ★ 先按数据里出现的方式插行 —— 必须在 clone 之前。
    # 插完再复制，上年页与本页天然同构（copy_worksheet 连样式 / 条件格式 / 行高一起复制）。
    # 反过来做：本年 9 行、上年 7 行，check_sheets 的「两页几何一致」整片错位。
    #
    # 列计划也必须**先算**（用插行前的 layout 也行，行号不影响列号），
    # 但真正落地要排在插行之后 —— 见下面的解释。
    ws_k = _sheet(wb, k01.SHEET_K01, "K.01")
    lay0 = k01.layout(ws_k)
    blocks = k01.plan_blocks(lay0, p["cat_order"])
    row_plan = k01.plan_inserts(lay0, m)
    row_trace = xlrows.apply_plan(wb, ws_k, row_plan) if row_plan else {"inserts": [], "xrefs": []}
    lay = k01.layout(ws_k) if row_plan else lay0
    if row_plan:
        k01.rewrite_year_ends(ws_k, lay)

    # 4.2) ★ 再动列。**必须先插行后动列**：
    # `xlrows.write_skeleton_row` 把合计列写死在第 26/27/28 列（Z/AA/AB），
    # 先动列的话那三列已经不是合计块了，骨架公式会写进错的列 —— 而合计照样平。
    col_trace = {"mode": "none", "inserts": 0, "formula_shift": 0, "xrefs": 0,
                 "clone": {}, "across": 0, "cf": {}}
    roles = {"book": [b["money"] for b in blocks["blocks"]],
             "adj": [b["name"] for b in blocks["blocks"]],
             "aud": [b["aud"] for b in blocks["blocks"]]}

    if blocks["n_new"]:
        col_trace["mode"] = "insert"
        col_trace["inserts"] = 1
        ins = xlcols.insert_cols_keep(ws_k, blocks["insert_at"], blocks["insert_n"])
        col_trace["insert"] = ins
        # 先平移公式，再克隆母本块 —— 顺序反了，克隆出来的公式会被二次 +3
        col_trace["formula_shift"] = xlcols.shift_sheet_formulas_cols(
            ws_k, blocks["insert_at"], blocks["insert_n"])
        hits = xlcols.shift_workbook_col_refs(
            wb, ws_k.title, blocks["insert_at"], blocks["insert_n"])
        col_trace["xrefs"] = len(hits)
        col_trace["xref_hits"] = [(a, b) for a, b, _c, _d in hits]

        src = blocks["blocks"][blocks["n_have"] - 1]["money"]
        # 新块宽度取「各类别块同位置的最大值」：沿用相邻那块会把 2.66 亿
        # 塞进 11.8 列宽，整列显示 ####
        wmax = tuple(
            max((ws_k.column_dimensions[get_column_letter(blocks["blocks"][i][role])].width or 0)
                for i in range(blocks["n_need"]))
            for role in ("money", "name", "aud")
        )
        col_trace["clone"] = xlcols.clone_block_cols(
            ws_k, src, blocks["insert_at"], range(1, ws_k.max_row + 1),
            value_rows_max=lay["cat_row"] + 1, widths=wmax)
        xlcols.shift_cf_cols(ws_k, blocks["insert_at"], blocks["insert_n"])

    elif blocks["delete_n"]:
        # 类别数少于模板预留块数：把「最后一个实际类别」之后、合计块之前的空块
        # **整段删掉**，否则 K.01 中间挂着几组 `【…】`，看着像还没填完。
        col_trace["mode"] = "delete"
        col_trace["delete_cols"] = xlshrink.delete_cols_keep(
            ws_k, blocks["delete_at"], blocks["delete_n"])
        # 横向合计必须**先**按新块列表重生成：删列后 `Q12` 里还留着 `Q12` 自己
        # （原来的 7 项写法），通用平移会因为它「引用落在删除区」直接抛错。
        lay_tmp = k01.layout(ws_k)
        col_trace["across"] = xlshrink.rewrite_total_sums(
            ws_k, k01.total_cols(lay_tmp), roles)

    # 4.3) 表2 / 表3 的行数对齐（多则插、少则删）
    # 放在动列之后：行里要写 `=E10` 这类块表头引用，必须按最终列号生成。
    # 传的是**块列表**不是计划字典 —— 传错时 len() 出来是字典的键数（7），
    # 正好等于模板预留行数，于是什么都不做，而后面填表时会静默错位。
    exp = k01.expand_table23(ws_k, wb, k01.layout(ws_k), blocks["blocks"])
    if exp["extra"]:
        col_trace["table23_extra"] = exp["extra"]
    if exp["deleted"]:
        col_trace["table23_deleted"] = exp["deleted"]

    # 4.3.5) 删列模式的收尾：表2/表3 的空行已经删掉，那些 `=R10/=U10/=X10`
    # 指向被删列的公式也一并消失，这时才做列方向的通用平移。
    # 顺序反了就会在「一行马上要被删掉的公式」上停下来报错。
    if blocks["delete_n"]:
        col_trace["formula_shift"] = len(
            xlshrink.shift_sheet_formulas_del(ws_k, "col",
                                              blocks["delete_at"], blocks["delete_n"]))
        hits = xlshrink.shift_workbook_refs_del(
            wb, ws_k.title, "col", blocks["delete_at"], blocks["delete_n"])
        col_trace["xrefs"] = len(hits)
        col_trace["xref_hits"] = [(a, b) for a, b, _c, _d in hits]
        col_trace["cf"] = xlshrink.shift_cf_del(
            ws_k, "col", blocks["delete_at"], blocks["delete_n"])
        xlshrink.shift_dv_del(ws_k, "col", blocks["delete_at"], blocks["delete_n"])

    # 4.4) ★ 手术之后必须重跑 layout —— 否则 abs_cat 指向错的表头格，
    # SUMIFS 会把所有类别归到同一列：金额看着都有，分类全错。
    lay = k01.layout(ws_k)

    # 4.5) 横向合计重写（插列模式）：插列后 Z/AA/AB 仍是 7 项，新块进不去。
    # 这是全流程最危险的失效 —— CHECK 比的是「审定合计 − Lead」，两边一起少，
    # 恒为 0，文件能开、没有错误值、只有数是错的。
    # 删列模式已经在 4.2 里用 `rewrite_total_sums` 做过（判据不同，见那个函数的说明）。
    if blocks["n_new"]:
        col_trace["across"] = xlcols.rewrite_across_sums(ws_k, roles)

    # 4.6) 条件格式：那条挂了 N 块 × 3 段的「空值涂黄」必须整条重写，
    # 平移只会把老块挪到新列号，新类别的空格不会标黄（人工待办入口就没了）。
    segs = [(lay["sec"][k]["beg"], lay["sec"][k]["blank"]) for k in ("fa", "dep", "imp")]
    col_trace["cf"] = xlcols.rebuild_block_cf(
        ws_k, lay["cat_row"], lay["cat_cols"], segs,
        first_cell="%s%d" % (get_column_letter(lay["cat_cols"][0] - 1),
                             lay["rows"]["fa"]["beg"]))
    if blocks["n_new"]:
        new_money = [blocks["blocks"][i]["money"] for i in range(blocks["n_have"], blocks["n_need"])]
        xlcols.add_tb_cell_cf(ws_k, new_money,
                              [lay["tb"][k] for k in ("fa", "dep", "imp", "nbv")],
                              ref_money=blocks["blocks"][0]["money"])

    # 4.7) ★ 先复制上年两页 —— 必须在任何 fill 之前，且必须在上面的手术之后。
    # 反过来做会把本年数复制成上年数，页面上有数、看着像，一分钱错误都不报。
    anchor = "%s%d" % (get_column_letter(lay["cat_cols"][0] - 1), lay["rows"]["fa"]["beg"])
    embed.clone_sheet(wb, k01.SHEET_K01, embed.SHEET_PRIOR_K01,
                      after=k01.SHEET_K01, expect_empty_cell=anchor)
    embed.clone_sheet(wb, lead.SHEET, embed.SHEET_PRIOR_LEAD,
                      after=lead.SHEET, expect_empty_cell="J47")

    # 4.8) 上年页数值化：快照公式 → 清空 → 写上年数 → 回解同表单格引用
    wk_p = wb[embed.SHEET_PRIOR_K01]
    wl_p = wb[embed.SHEET_PRIOR_LEAD]
    saved_k = embed.clear_formulas(wk_p)
    saved_l = embed.clear_formulas(wl_p)
    lay_p = k01.layout(wk_p)
    cats_used = list(p["cat_order"])
    embed.fill_prior_k01(wk_p, lay_p, cats_used, prior_cats,
                         "上年明细资料未获取：本页仅有上年末余额（= 本年期初）。"
                         "年初余额与各变动行留空，浅黄底＝未获取，不是为零。")
    embed.resolve_simple_refs(wk_p, saved_k)
    lay_l_p = lead.locate_table1(wl_p)
    embed.fill_prior_lead(wl_p, lay_l_p["rows"], prior_cats, info, str(prior_end))
    embed.resolve_simple_refs(wl_p, saved_l)
    embed.assert_no_formula(wk_p, embed.SHEET_PRIOR_K01)
    embed.assert_no_formula(wl_p, embed.SHEET_PRIOR_LEAD)

    # 4.9) TB 页（原始 44 行；要拿 first/last 给 K.01 拼 SUMIFS 范围）
    tb = embed.write_tb_sheet(wb, prep.load_tb(conn), p["trace"]["dropped"], p["cats"])

    ws_fa = _sheet(wb, "FA list")
    ws_mv = _sheet(wb, "新增清单", "新增清单")
    ws_dp = _sheet(wb, "处置清单", "处置清单")
    cat_map = {
        "fa": tables.write_fa(ws_fa, syn["fa"]),
        "mv": tables.write_moves(ws_mv, syn["moves"]),
        "dp": tables.write_disposals(ws_dp, syn["disposals"]),
    }

    k_info = k01.fill_table1(ws_k, lay, p, m, prior_cats, cat_map,
                             os.path.basename(prior_path),
                             prior_sheet=embed.SHEET_PRIOR_K01, tb=tb,
                             blocks=blocks["blocks"])
    k01.fill_table23(ws_k, lay, cat_map, blocks["blocks"])
    k01.fill_table4(ws_k, lay, m["table4"], customer=_cust)

    ws_l = _sheet(wb, lead.SHEET, "Lead")
    l_head = lead.fill_header(ws_l, info)
    l_tab = lead.fill_table1(ws_l, prior_cats, os.path.basename(prior_path),
                             prior_sheet=embed.SHEET_PRIOR_LEAD)
    # 审计调整列接上底部「调整汇总表」。**必须在 clone 上年页（上面第 4.0 步）
    # 与 assert_no_formula 之后**：早一步写，clone 前的公式快照会把 H 列当成
    # 「模板自带公式」清掉，上年页的审定数就变成 G+空。
    l_adj = lead.fill_adj_summary(ws_l, l_tab)

    # 5) 留痕（写在 K.01 的 Notes 区）
    # 明细来源分两种，说法必须跟着变 —— 客户台账是客户给的真卡片，
    # 序时账重建的那份「资产编码/部门/年限」才是程序补的，不能混为一谈。
    if syn["trace"].get("source") == "cards":
        notes = [
            "【数据说明】本底稿固定资产明细取自客户资产台账（期初/期末两张表），金额取自客户真实账套；",
            "台账期初、期末与科目余额表逐类对齐，差异已调整到该类最大一张卡片上（见程序留痕）。",
        ]
    else:
        notes = [
            "【数据说明】本底稿固定资产明细为按科目辅助核算 + 序时账重建，金额取自客户真实账套；",
            "资产编码、使用部门、存放地点、折旧年限为程序补全，不可作为审计证据。",
        ]
    notes.append(
        "【TB 去重】TB 中 1601/1602/1603 被整块重复导入一次，已按「主体+科目+辅助核算+四个金额」去重，"
        "删除 %d 行；去重后与序时账净额精确勾稽。" % len(p["trace"]["dropped"]))
    notes.append(
        "【来源说明】年初余额、TB 核对块、审计调整均为同簿公式引用（公式里带页名与单元格），"
        "来源见公式本身；按用户要求不再写单元格批注。")
    for a in p["trace"]["adjusts"]:
        notes.append("【期初修正】%s %s：TB 期初 %s → 反推 %s（%s）"
                     % (a["company"], a["acct"], a["tb_open"], a["expect_open"], a["reason"]))
    for t in p["trace"]["cat_rules"]:
        if t["kind"] == "cat_fallback":
            notes.append("【类别归集】%s %s 无类别辅助核算，归入「%s」（金额 %s）"
                         % (t["company"], t["acct"], t["to"], t["amount"]))

    # 5.1) 写盘前再校一遍「写进 Excel 的东西」：上年页数值化、TB 页口径、公式形态
    chk2 = checks.check_sheets(p, {
        "k01_ws": ws_k, "k01_lay": lay, "lead_ws": ws_l,
        "lead_lay": lead.locate_table1(ws_l), "lead_adj": l_adj,
        "prior_k01": wk_p, "prior_lead": wl_p, "tb": tb,
        "prior_cats": prior_cats, "prep_trace": p["trace"],
        "prior_lay": lay_p, "blocks": blocks["blocks"],
    })
    if not chk2["ok"]:
        raise ValueError("sheet 层校验未通过：%s"
                         % "；".join("%s(%s)" % (f["name"], f["extra"]) for f in chk2["fails"]))

    # 5.2) 隐藏三张辅助页（放保存前最后一步，便于 trace 回报）
    hidden = embed.hide(wb, [embed.SHEET_PRIOR_K01, embed.SHEET_PRIOR_LEAD, embed.SHEET_TB])

    # 5.3) 数字格式归一：模板自带的 ¥ / $ / € 与代码写的 [Red] 格式并存，
    # 同一页会出现两种零、两种负号。只改 number_format，不动值/字体/边框。
    fmt_tr = fmt.normalize_money(wb)

    # 6) 保存
    os.makedirs(out_dir, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(out_dir, "固定资产SWP_%s.xlsx" % stamp)
    wb.save(out)
    latest = os.path.join(out_dir, "固定资产SWP_最新.xlsx")
    try:
        shutil.copyfile(out, latest)
    except PermissionError:
        latest = os.path.join(out_dir, "固定资产SWP_最新(2).xlsx")
        shutil.copyfile(out, latest)

    trace = {
        "check_ok": chk["ok"],
        "sheet_check_ok": chk2["ok"],
        "row_plan": row_plan,
        "rows_inserted": sum(len(x["labels"]) for x in row_plan),
        "block_plan": {"n_have": blocks["n_have"], "n_need": blocks["n_need"],
                       "n_new": blocks["n_new"], "insert_at": blocks["insert_at"]},
        "col_trace": col_trace,
        "xrefs": [len(x["other"]) for x in row_trace.get("xrefs", [])],
        "hidden": hidden,
        "tb_rows": tb["rows"],
        "tb_dup": tb["dup"],
        "check_items": len(chk["items"]),
        "customer": info.get("customer"),
        "balance_date": str(balance_date),
        "totals": {k: str(v) for k, v in p["totals"].items()},
        "net": {"beg": str(p["net_beg"]), "end": str(p["net_end"])},
        "k01": {"%s|%s" % k: str(v) for k, v in m["totals"].items()},
        "table4": {k: str(v) for k, v in m["table4"]["buckets"].items()},
        "prep": p["trace"],
        "synth": syn["trace"],
        "cats": k_info["cats"],
        "lead_adj": l_adj,
        "fmt": fmt_tr,
        "cat_order": p["cat_order"],
        "out": out,
        "latest": latest,
        "prior": prior_path,
    }
    return {"path": out, "latest": latest, "prior": prior_path,
            "trace": trace, "prep": p, "moves": m, "synth": syn,
            "k01_layout": lay, "lead": l_tab, "checks": chk}
