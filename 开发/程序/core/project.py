# -*- coding: utf-8 -*-
"""项目配置与客户配置包（.ptwpack）。

对应 Skill 附录 B 的 `project.yaml`，以及"下一年直接导入复用"这件事。

**为什么拆两轨**（这是设计上最重要的一条）
  结构类（期间分段、缩放、sheet 匹配、阈值来源）→ 导入即可用
  判断类（性质、是否函证、受限、特殊拆分）→ 导入后一律 `draft`，必须人复核

混在一起一键套用，人会以为"都对了"，但实际上去年的判断今年可能已经变了
（比如去年受限的账户今年解禁）。审计上这是不可接受的。
"""
import json
import os
import zipfile
from datetime import datetime

PACK_VERSION = 1


# ---------------------------------------------------------------- 项目配置
def default_config():
    """新项目的默认配置。字段语义对齐 Skill 附录 B。"""
    return {
        "version": PACK_VERSION,
        "customer": "",
        "period": "",
        "period_split": {"pre": [1, 6], "final": [7, 12]},
        "sheets": {
            "cash_lead": ["C.00 *Lead*"],
            "cash_bkd": ["C.00 BKD"],
            "cash_sec": ["C.01b *"],
            "cash_class": ["C.00b *"],
            "cash_recon": ["C.02 *"],
            "cash_fx": ["C.02a *"],
            "exp_lead": ["Uexp.00 Lead"],
            "exp_sell": ["VC.00 *销售费用BKD*"],
            "exp_admin": ["VD.00 *管理费用BKD*"],
            "tod_pool": ["样本池-TOD"],
            "tod_neg": ["样本池-负值"],
            "tod_cut": ["样本池-CUT*"],
        },
        "rules": {
            "fx": {"source": "cell", "dateRequired": True},
            "restricted": {"column": "受限资金", "reclassToLead": True},
            "te": {"source": "template", "fallback": 100000},
        },
        "tolerances": {"amount": 0.01, "ratio": 1e-6},
    }


def config_path(data_dir):
    return os.path.join(data_dir, "project.json")


def load_config(data_dir):
    p = config_path(data_dir)
    cfg = default_config()
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as fh:
                d = json.load(fh)
            if isinstance(d, dict):
                # 浅合并：文件里改了几个字段，其余沿用默认，避免老配置文件少一个键就崩
                for k, v in d.items():
                    if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                        cfg[k].update(v)
                    else:
                        cfg[k] = v
        except Exception:
            pass
    return cfg


def save_config(data_dir, cfg):
    p = config_path(data_dir)
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=1)
    return p


# ---------------------------------------------------------------- 导出
def _collect_judgments(ledger, scope=None):
    """只导 confirmed。draft 是 AI 起草未复核的，带出去会污染下一个项目。"""
    out = []
    try:
        rows = ledger.conn.execute(
            "SELECT scope,entity,subject,dim_type,dim_key,period,field,value,"
            "source,basis,confidence,created_by,created_at FROM _judgments "
            "WHERE status='confirmed'" + (" AND scope=?" if scope else ""),
            ([scope] if scope else [])).fetchall()
    except Exception:
        return out
    for r in rows:
        out.append({"scope": r[0], "entity": r[1], "subject": r[2],
                    "dim_type": r[3], "dim_key": r[4], "period": r[5],
                    "field": r[6], "value": r[7], "source": r[8],
                    "basis": r[9], "confidence": r[10],
                    "by": r[11], "at": str(r[12] or "")})
    return out


def _collect_recipes(ledger):
    out = []
    try:
        rows = ledger.conn.execute("SELECT name, steps, note FROM _recipes").fetchall()
    except Exception:
        return out
    for name, steps, note in rows:
        try:
            st = json.loads(steps or "[]")
        except Exception:
            st = []
        out.append({"name": name, "steps": st, "note": note or ""})
    return out


def export_pack(ledger, data_dir, out_path, customer="", period="", scope=None):
    """把当前项目的配置、配方、已确认判断打成一个 zip。"""
    cfg = load_config(data_dir)
    if customer:
        cfg["customer"] = customer
    if period:
        cfg["period"] = period
    fp = None
    fp_dir = os.path.join(data_dir, "fp")
    if os.path.isdir(fp_dir):
        cands = sorted((os.path.join(fp_dir, f) for f in os.listdir(fp_dir)
                        if f.endswith(".json")), key=os.path.getmtime)
        if cands:
            try:
                with open(cands[-1], encoding="utf-8") as fh:
                    fp = json.load(fh)
            except Exception:
                fp = None
    # 导出即视为认可当前配置，落一份到本地：否则下次进来还是默认值，
    # 用户会以为"我明明导出过"
    save_config(data_dir, cfg)
    meta = {"packVersion": PACK_VERSION,
            "customer": cfg.get("customer") or "",
            "period": cfg.get("period") or "",
            "createdAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "judgmentCount": 0}
    judg = _collect_judgments(ledger, scope=scope)
    meta["judgmentCount"] = len(judg)
    recipes = _collect_recipes(ledger)
    meta["recipeCount"] = len(recipes)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        def put(name, obj):
            z.writestr(name, json.dumps(obj, ensure_ascii=False, indent=1))
        put("meta.json", meta)
        put("config.json", cfg)
        if fp:
            put("fingerprint.json", fp)
        put("judgments.json", judg)
        put("recipes.json", recipes)
    return {"path": out_path, "size": os.path.getsize(out_path), "meta": meta}


# ---------------------------------------------------------------- 导入
def import_pack(ledger, data_dir, path, apply_structure=True, apply_judgments=True):
    """导入配置包。

    返回 {config: 动作, judgments: 导入条数, recipes: 条数, fingerprint: 比对结果, warnings}

    判断一律写成 draft（等人工复核），不直接 confirmed——即便导出时它们是 confirmed。
    """
    out = {"config": "skipped", "judgments": 0, "recipes": 0,
           "fingerprint": None, "warnings": []}
    if not os.path.isfile(path):
        raise FileNotFoundError("配置包不存在：%s" % path)
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())

        def get(name):
            try:
                return json.loads(z.read(name).decode("utf-8"))
            except Exception:
                return None

        meta = get("meta.json") or {}
        if int(meta.get("packVersion") or 0) > PACK_VERSION:
            out["warnings"].append("配置包版本 %s 高于本程序 %s，可能有字段读不懂"
                                   % (meta.get("packVersion"), PACK_VERSION))

        # 1) 结构轨：配置直接落盘
        cfg = get("config.json")
        if cfg and apply_structure:
            cur = load_config(data_dir)
            cur.update(cfg if isinstance(cfg, dict) else {})
            out["config"] = "imported → " + save_config(data_dir, cur)

        # 2) 指纹轨：只比对，不覆盖本地（本地指纹代表当前模板的真实状态）
        fp_new = get("fingerprint.json")
        if fp_new:
            from core import fingerprint as FP
            fp_cur = FP.load(FP.path_for(data_dir, "pack"))
            if fp_cur:
                diffs = FP.diff(fp_cur, fp_new)
                out["fingerprint"] = {"diffs": diffs, "blocking": len(FP.blocking(diffs))}
            else:
                out["fingerprint"] = {"diffs": [], "blocking": 0,
                                      "note": "本地无指纹，无法比对（导入的是包里的历史指纹）"}

        # 3) 配方轨：同名覆盖
        recs = get("recipes.json") or []
        for r in recs:
            try:
                ledger.recipes_save(r.get("name") or "", r.get("steps") or [],
                                    note=r.get("note") or "")
                out["recipes"] += 1
            except Exception as e:
                out["warnings"].append("配方 %s 导入失败：%s" % (r.get("name"), e))

        # 4) 判断轨：一律 draft
        judg = get("judgments.json") or []
        if judg and apply_judgments:
            from core import judge as J
            J.ensure(ledger.conn)
            for j in judg:
                try:
                    _jid, st, err = J.upsert(
                        ledger.conn, j.get("scope") or "", j.get("entity"),
                        j.get("subject"), j.get("dim_type") or "acct",
                        j.get("dim_key"), j.get("period") or "",
                        j.get("field") or "", j.get("value"),
                        source="import", basis=(j.get("basis") or "") + "（导入自配置包）",
                        confidence=j.get("confidence"), status="draft",
                        created_by=j.get("by") or "")
                    if err:
                        out["warnings"].append("判断 %s/%s 导入失败：%s"
                                               % (j.get("entity"), j.get("field"), err))
                    else:
                        out["judgments"] += 1
                except Exception as e:
                    out["warnings"].append("判断导入异常：%s" % e)
        out["meta"] = {"customer": meta.get("customer"), "period": meta.get("period"),
                       "createdAt": meta.get("createdAt")}
        out["next"] = ("结构配置已应用；%d 条判断已导入为「待复核」，"
                       "在判断留痕页确认后才算数" % out["judgments"])
    return out


def list_packs(folder):
    if not os.path.isdir(folder):
        return []
    out = []
    for f in sorted(os.listdir(folder)):
        if f.lower().endswith(".ptwpack"):
            p = os.path.join(folder, f)
            out.append({"file": f, "path": p, "size": os.path.getsize(p),
                        "mtime": datetime.fromtimestamp(os.path.getmtime(p))
                        .strftime("%Y-%m-%d %H:%M")})
    return out
