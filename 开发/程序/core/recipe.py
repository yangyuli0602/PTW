# -*- coding: utf-8 -*-
"""清洗方案导出 / 导入 xlsx。

为什么用 xlsx 而不是 json
    审计组换版本、换机器时，需要一份人能看懂、能手改、能放进底稿的文件。
    一种清洗模式一个 sheet，改了 sheet 名就等于改了方案名。

文件结构
    说明         格式说明（导入时跳过）
    _方案清单     方案 → 原来用在哪些工作表的表头上
    _映射        字段映射备份，按「表头签名」关联（不是按文件名）
    方案1/方案2…  每行一个动作：序号 | 动作 | 类型 | 启用 | 参数(JSON) | 说明

一个动作的参数是变长的（有的要选列、有的要填关键词），所以参数统一压成
一列 JSON。手改也简单：cols 是数组，其余是字符串/数字。
"""
import hashlib
import json
import re

HEAD = ["序号", "动作", "类型", "启用", "参数(JSON)", "说明"]
JSON_COL = "参数(JSON)"


# ---------------------------------------------------------------- 表头签名
def _norm_name(s):
    """列名规范化：全角转半角、去空白、去括号与常见标点、转小写。

    目的是让「期初借方(本币)」「期初借方（本币）」「 期初借方(本币) 」算同一个名字。
    """
    s = str(s or "")
    out = []
    for ch in s:
        code = ord(ch)
        if code == 0x3000:
            ch = " "
        elif 0xFF01 <= code <= 0xFF5E:
            ch = chr(code - 0xFEE0)
        out.append(ch)
    s = "".join(out).lower()
    s = re.sub(r"[\s_\-()\[\]{}<>:：,，.。、/\\'\"]", "", s)
    return s


def head_key(cols):
    """表头指纹的原文（前端分组用的是同一个串，便于前后端对齐）。

    排过序：同一批列名、顺序不同也算同一套表头——映射是按列名走的，与列序无关。
    """
    names = sorted(_norm_name(c) for c in (cols or []) if str(c or "").strip())
    return "|".join(names)


def sig_of(cols):
    """表头签名：head_key 的 md5 前 12 位。写进 xlsx，人看不出来但机器认得。"""
    return hashlib.md5(head_key(cols).encode("utf-8")).hexdigest()[:12]


def head_brief(cols, n=4):
    """给人看的表头摘要：前 n 个列名 + 总列数。"""
    cs = [str(c or "") for c in (cols or []) if str(c or "").strip()]
    head = "、".join(cs[:n])
    return (head + ("…" if len(cs) > n else "")) + ("（%d 列）" % len(cs) if cs else "")


def safe_sheet(name, used=None):
    s = re.sub(r"[\\/*?:\[\]]", "_", str(name or "")).strip() or "方案"
    s = s[:31]
    if used is None:
        return s
    base, n = s, 1
    while s in used:
        n += 1
        tail = "(%d)" % n
        s = base[:31 - len(tail)] + tail
    return s


def steps_key(steps):
    """步骤签名：动作类型 + 启用 + 参数完全一样才算同一套方案。"""
    return json.dumps([{k: s.get(k) for k in ("type", "enabled", "params")}
                       for s in (steps or [])], sort_keys=True, ensure_ascii=False)


def group(sheets):
    """把「工作表 + 清洗步骤」归并成方案：步骤完全相同的合成一个方案。

    sheets: [{"file_name","sheet_name","target","steps","mapping","columns"}]
    返回 [{name, key, steps, uses:[{file,sheet,target,n_steps,sig,head}]}]
    """
    plans, index = [], {}
    for sh in sheets:
        steps = sh.get("steps") or []
        key = steps_key(steps)
        if key not in index:
            p = {"name": "方案%d" % (len(plans) + 1), "key": key,
                 "steps": steps, "uses": []}
            index[key] = p
            plans.append(p)
        cols = sh.get("columns") or []
        index[key]["uses"].append({
            "file": sh.get("file_name") or "",
            "sheet": sh.get("sheet_name") or "",
            "target": sh.get("target") or "",
            "n_steps": len(steps),
            "sig": sh.get("sig") or sig_of(cols),
            "head": head_brief(cols),
        })
    for n, p in enumerate(plans, 1):
        u = p["uses"][0]
        # 方案名带上第一个使用它的工作表名，比「方案1」好认，也便于导入时对号入座
        tail = re.sub(r"\.(xlsx|xlsm|csv|txt)$", "", u["file"] or "", flags=re.I)
        label = (u["sheet"] or tail or "").strip()
        p["auto_name"] = "方案%d" % n
        if label:
            p["name"] = ("%s·%s" % (p["auto_name"], label))[:28]
    return plans


def export(plans, maps, path, meta=None):
    """写 xlsx。maps: {工作表名: {"target":.., "mapping":{field:col}}}"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="EFF3F7")

    ws = wb.active
    ws.title = "说明"
    lines = [
        ("PTW平台 · 清洗方案导出", 1),
        ("", 0),
        ("用途：把这次调好的清洗动作存成文件，换版本 / 换机器后重新导入，不用再调一遍。", 0),
        ("", 0),
        ("表结构", 1),
        ("一种清洗模式一个 sheet。sheet 名就是方案名，可以随手改成「余额表-通用」这类好认的名字。", 0),
        ("每个 sheet 的第一行是表头，往下每行一个动作，按「序号」从上到下执行。", 0),
        ("  序号    执行顺序，可以留空，留空按行序", 0),
        ("  动作    给人看的名字，导入时可以不管，系统认「类型」", 0),
        ("  类型    动作类型（drop_empty / drop_keyword / trim_text …），必须能对上系统内置动作", 0),
        ("  启用    1=启用，0=停用。填 0 的动作保留但不执行", 0),
        ("  参数(JSON)  该动作的参数，例如 {\"cols\": [\"摘要\"], \"mode\": \"all\"}", 0),
        ("  说明    备注，导入时忽略", 0),
        ("", 0),
        ("_方案清单  记录每个方案原来用在哪些工作表的表头上，导入时用来对号入座。", 0),
        ("_映射      字段映射备份。**按表头签名关联，不是按文件名**——文件名每年都变，表头不变。", 0),
        ("           同一套表头（列名相同、顺序可以不同）在任何文件里都能自动配上这份映射。", 0),
        ("           表头签名 = 列名规范化排序后的 md5，换个表格文件不用重配。", 0),
        ("", 0),
        ("手工改注意", 1),
        ("1. 参数里的列名必须和实际工作表的列名一字不差（含全角括号、空格）。", 0),
        ("2. 不要改「参数(JSON)」这一列的表头，改了系统就找不到参数。", 0),
        ("3. 复制一个 sheet 改名，就等于多存了一个方案。", 0),
    ]
    for i, (t, b) in enumerate(lines, 1):
        ws.cell(row=i, column=1, value=t)
        if b:
            ws.cell(row=i, column=1).font = Font(bold=True, size=13 if i == 1 else 11)
    ws.column_dimensions["A"].width = 110

    ws2 = wb.create_sheet("_方案清单")
    ws2.append(["方案", "文件", "工作表", "目标表", "动作数", "表头签名", "表头摘要"])
    for c in ws2[1]:
        c.font = bold
        c.fill = fill
    for p in plans:
        for u in p["uses"]:
            ws2.append([p["name"], u["file"], u["sheet"], u["target"], u["n_steps"],
                        u.get("sig") or "", u.get("head") or ""])
    for col, w in zip("ABCDEFG", (28, 34, 26, 12, 10, 16, 44)):
        ws2.column_dimensions[col].width = w

    # 字段映射按「表头签名」存：换一张同表头的表照样认得，与文件名无关
    ws3 = wb.create_sheet("_映射")
    ws3.append(["表头签名", "列数", "表头摘要", "表头(规范化)", "目标表", "标准字段", "原始列"])
    for c in ws3[1]:
        c.font = bold
        c.fill = fill
    for sig, m in (maps or {}).items():
        mp = (m or {}).get("mapping") or {}
        cols = (m or {}).get("cols") or []
        for k, v in mp.items():
            if str(k).startswith("_"):
                continue
            ws3.append([sig, len(cols), head_brief(cols, 8), head_key(cols),
                        (m or {}).get("target") or "", k, str(v)])
    for col, w in zip("ABCDEFG", (16, 8, 46, 60, 12, 20, 30)):
        ws3.column_dimensions[col].width = w

    used = {"说明", "_方案清单", "_映射"}
    for p in plans:
        title = safe_sheet(p["name"], used)
        used.add(title)
        s = wb.create_sheet(title)
        s.append(HEAD)
        for c in s[1]:
            c.font = bold
            c.fill = fill
        for i, st in enumerate(p["steps"], 1):
            s.append([i, st.get("_name") or st.get("name") or st.get("type"),
                      st.get("type"),
                      0 if st.get("enabled") is False else 1,
                      json.dumps(st.get("params") or {}, ensure_ascii=False),
                      st.get("_desc") or ""])
        for col, w in zip("ABCDEF", (6, 18, 18, 6, 60, 42)):
            s.column_dimensions[col].width = w
        for row in s.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(vertical="top", wrap_text=True)
    if meta:
        wb.properties.title = "PTW平台清洗方案"
        wb.properties.description = json.dumps(meta, ensure_ascii=False)
    wb.save(path)
    return path


def parse(path):
    """读回 xlsx，返回 {plans:[{name, steps}], maps:{sheet: {target, mapping}}}"""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    plans, maps, warns = [], {}, []
    uses = {}
    for ws in wb.worksheets:
        title = str(ws.title or "")
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        if title.startswith("_方案清单"):
            head = [str(c or "").strip() for c in rows[0]]
            try:
                i_p = head.index("方案")
            except ValueError:
                continue
            i_f = head.index("文件") if "文件" in head else -1
            i_s = head.index("工作表") if "工作表" in head else -1
            i_t = head.index("目标表") if "目标表" in head else -1
            for r in rows[1:]:
                if not r or not r[i_p]:
                    continue
                uses.setdefault(str(r[i_p]).strip(), []).append({
                    "file": str(r[i_f] or "").strip() if i_f >= 0 else "",
                    "sheet": str(r[i_s] or "").strip() if i_s >= 0 else "",
                    "target": str(r[i_t] or "").strip() if i_t >= 0 else "",
                })
            continue
        if title.startswith("_映射") or title.lower().startswith("说明"):
            if title.startswith("_映射"):
                head = [str(c or "").strip() for c in rows[0]]
                try:
                    i_field = head.index("标准字段")
                    i_col = head.index("原始列")
                except ValueError:
                    continue
                # 新格式按「表头签名」，老文件（v1.1.0 及以前）按「工作表」名，读进来标 legacy
                i_sig = head.index("表头签名") if "表头签名" in head else -1
                i_key = head.index("表头(规范化)") if "表头(规范化)" in head else -1
                i_sheet = head.index("工作表") if "工作表" in head else -1
                i_tgt = head.index("目标表") if "目标表" in head else -1
                i_n = head.index("列数") if "列数" in head else -1
                for r in rows[1:]:
                    if not r or len(r) <= max(i_field, i_col):
                        continue
                    hkey = ""
                    if i_key >= 0 and i_key < len(r):
                        hkey = str(r[i_key] or "").strip()
                    if i_sig >= 0:
                        key = str(r[i_sig] or "").strip() if i_sig < len(r) else ""
                        legacy = False
                    elif i_sheet >= 0:
                        key = str(r[i_sheet] or "").strip() if i_sheet < len(r) else ""
                        legacy = True
                    else:
                        continue
                    f, c = str(r[i_field] or "").strip(), str(r[i_col] or "").strip()
                    if not key or not f or not c:
                        continue
                    slot = maps.setdefault(key, {"target": "", "mapping": {},
                                                 "cols": [], "legacy": legacy,
                                                 "hkey": hkey})
                    if hkey:
                        slot["hkey"] = hkey
                    if i_tgt >= 0 and i_tgt < len(r) and r[i_tgt]:
                        slot["target"] = str(r[i_tgt]).strip()
                    if i_n >= 0 and i_n < len(r) and r[i_n]:
                        try:
                            slot["n_cols"] = int(r[i_n])
                        except Exception:
                            pass
                    slot["mapping"][f] = c
            continue
        # 找表头行：同时有「类型」和「参数」的那一行
        hi = None
        for i, r in enumerate(rows[:10]):
            cells = [str(c or "").strip() for c in r]
            if "类型" in cells and any("参数" in c for c in cells):
                hi = i
                break
        if hi is None:
            warns.append("工作表「%s」找不到表头（需要「类型」和「参数(JSON)」两列），已跳过" % title)
            continue
        head = [str(c or "").strip() for c in rows[hi]]
        def col(*names):
            for n in names:
                if n in head:
                    return head.index(n)
            return -1
        i_no, i_ty = col("序号"), col("类型")
        i_en = col("启用")
        i_pj = col(JSON_COL, "参数", "参数（JSON）", "参数(Json)")
        i_nm = col("动作", "名称")
        steps = []
        for r in rows[hi + 1:]:
            if not r or not any(v not in (None, "") for v in r):
                continue
            ty = str(r[i_ty] or "").strip() if i_ty >= 0 else ""
            if not ty:
                continue
            en = True
            if i_en >= 0:
                v = str(r[i_en] if r[i_en] is not None else "").strip().lower()
                en = v not in ("0", "false", "否", "no", "n", "停用", "关闭")
            params = {}
            if i_pj >= 0 and r[i_pj] not in (None, ""):
                try:
                    params = json.loads(str(r[i_pj]))
                    if not isinstance(params, dict):
                        params = {}
                except Exception:
                    warns.append("工作表「%s」第 %s 行动作 %s 的参数不是合法 JSON，已按空参数处理"
                                 % (title, r[i_no] or "?", ty))
            if i_no >= 0 and r[i_no] not in (None, ""):
                try:
                    params["_no"] = int(r[i_no])
                except Exception:
                    pass
            steps.append({"type": ty, "enabled": en, "params": params,
                          "name": str(r[i_nm] or ty) if i_nm >= 0 else ty})
        if steps:
            steps.sort(key=lambda s: s.get("params", {}).get("_no") or 0)
            for s in steps:
                s.get("params", {}).pop("_no", None)
            plans.append({"name": title, "steps": steps,
                          "uses": uses.get(title, [])})
        else:
            warns.append("工作表「%s」没有可用的动作行" % title)
    try:
        wb.close()
    except Exception:
        pass
    return {"plans": plans, "maps": maps, "warnings": warns}


def normalize_steps(steps, valid_types, existing_cols=None):
    """把读回来的步骤清洗成系统能跑的样子。

    - 丢掉系统不认识的类型（宁可少跑一步，也不要整个清洗炸掉）
    - 参数字段按动作定义的 key 过滤，多余键（如 _no）去掉
    - cols 里的列名如果当前表里没有，自动丢掉并记录
    """
    out, warns = [], []
    for i, st in enumerate(steps, 1):
        ty = str(st.get("type") or "").strip()
        spec = valid_types.get(ty)
        if not spec:
            warns.append("第 %d 步「%s」：系统没有这个动作类型，已跳过" % (i, ty))
            continue
        raw = dict(st.get("params") or {})
        keep = {p["key"] for p in (spec.get("params") or [])}
        params = {}
        for k in keep:
            if k in raw:
                params[k] = raw[k]
        if existing_cols:
            if isinstance(params.get("cols"), list):
                miss = [c for c in params["cols"] if c not in existing_cols]
                if miss:
                    warns.append("第 %d 步「%s」：列 %s 在当前工作表里不存在，已去掉"
                                 % (i, spec["name"], "、".join(map(str, miss[:5]))))
                params["cols"] = [c for c in params["cols"] if c in existing_cols]
            c1 = params.get("col")
            if c1 and c1 not in existing_cols:
                warns.append("第 %d 步「%s」：列「%s」在当前工作表里不存在，这一步会跳过"
                             % (i, spec["name"], c1))
        out.append({"id": "r%d" % i, "type": ty,
                    "enabled": st.get("enabled", True) is not False,
                    "params": params})
    return out, warns
