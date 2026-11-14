# -*- coding: utf-8 -*-
"""一条命令：按归档配方把一套客户账表从「导入」跑到「六组检查」。

    python 01_测试/自动回归/run_clean.py \
        --data-root "<还原目录>" \
        --recipe "04_归档/01_PTW项目级配置文件/2026海柔PTW配置文件/海柔清洗配方_20261112.json" \
        --std-accounts "<同目录>/海柔_客户科目表_std_accounts.json" \
        --work-dir tmp/海柔复跑_20261113 --port 8899

脚本自己起服务、自己停，数据目录一律隔离（KZ_DATA_DIR=work_dir/data），绝不碰平台 data/。

设计上与早前那种"按表头行提示逐条 import 试错"的脚本有三点不同：
  1. 先识别、后导入：在 preview 上把「表头行 × 规则」一次匹配出来，每张表只 import 一次
     （100 万行的凭证表按旧写法会被导 4 次，等于 4 倍耗时）
  2. mapping 放在 clean/apply 之后：apply 会把 mapping 里"当前不存在的列"过滤掉，
     而「一级科目 / 唯一凭证号」是清洗动作生成的
  3. 409（清洗门禁）不吞：解析 blocked.items 打出来，绝不自动 force

退出码：0 全绿 / 2 前置错误 / 3 服务起不来 / 4 有表未命中或导入失败 / 5 merge 被拦 / 6 六组未全绿
"""
import argparse
import hashlib
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = HERE
while not os.path.isdir(os.path.join(ROOT, "程序")):
    parent = os.path.dirname(ROOT)
    if parent == ROOT:
        break
    ROOT = parent

DEFAULT_RECIPE = os.path.join(ROOT, "04_归档", "01_PTW项目级配置文件",
                              "2026海柔PTW配置文件", "海柔清洗配方_20261112.json")
DEFAULT_STD = os.path.join(ROOT, "04_归档", "01_PTW项目级配置文件",
                           "2026海柔PTW配置文件", "海柔_客户科目表_std_accounts.json")


# ----------------------------------------------------------------- 小工具
def jload(p):
    with open(p, encoding="utf-8") as fh:
        return json.load(fh)


def md5_of(p):
    h = hashlib.md5()
    with open(p, "rb") as fh:
        for chunk in iter(lambda: fh.read(4 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def free_port(start, tries=10):
    for i in range(tries):
        p = start + i
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return None


def norm(s):
    s = "" if s is None else str(s).strip()
    return (s.replace("（", "(").replace("）", ")")
             .replace("－", "-").replace("—", "-"))


def header_names(preview, hr, hrows):
    """复刻平台 _header_names：多行表头对上行做左填充，用 _ 连接，重名加序号。"""
    block = preview[hr - 1: hr - 1 + max(1, hrows)]
    if not block:
        return []
    width = max(len(r) for r in block)
    rows = []
    for i, r in enumerate(block):
        r = [norm(v) for v in list(r) + [""] * (width - len(r))]
        if i < len(block) - 1:
            last = ""
            for j in range(width):
                if r[j]:
                    last = r[j]
                else:
                    r[j] = last
        rows.append(r)
    out, seen = [], {}
    for j in range(width):
        parts = [rows[i][j] for i in range(len(rows)) if rows[i][j] and rows[i][j] not in ("-", "/")]
        nm = "_".join(dict.fromkeys(parts)) or ("列%d" % (j + 1))
        seen[nm] = seen.get(nm, 0) + 1
        out.append(nm if seen[nm] == 1 else "%s_%d" % (nm, seen[nm]))
    return out


def detect_ok(cols, d):
    cols = set(cols or [])
    if any(c not in cols for c in d.get("must", [])):
        return False
    one = d.get("must_one_of") or d.get("any") or []
    if one and not (set(one) & cols):
        return False
    anyk = d.get("must_any") or (d.get("any") if not one else [])
    if anyk and not (set(anyk) & cols):
        return False
    if set(d.get("none", [])) & cols:
        return False
    return True


def detect_score(cols, d):
    cols = set(cols or [])
    return len(set(d.get("must", [])) & cols) * 10 + len(
        (set(d.get("must_any", [])) | set(d.get("must_one_of", []))) & cols)


# ------------------------------------------------ 配方里的 {{公司名…}} 占位符
# 平台不认 `{{...}}`：它只是配方里给「运行者」的占位约定（谁跑谁负责填）。
# 不解析的后果很隐蔽——占位符字符串会当成公司名落库，两家不同主体被并成一家，
# 「按公司分组」的末级科目检查随即误报父子并存（2026-11-13 瑞云实测 3 个假报）。
PLACEHOLDER = re.compile(r"\{\{\s*公司名[^{}]*\}\}")


def company_from_preview(preview, max_rows=40):
    """表内标题行找「公司名称：xxx」（单主体余额表第 2 行就是）。"""
    for row in (preview or [])[:max_rows]:
        for v in row or []:
            s = norm(v)
            m = re.match(r"^(?:公司|单位|企业)名称\s*[:：]\s*(.+)$", s)
            if m:
                return m.group(1).strip()
    return None


def company_from_name(file_name):
    """文件名兜底：`2026年9期 深圳市瑞影科技有限公司_序时账.xlsx` → 公司全称。"""
    base = os.path.splitext(os.path.basename(file_name))[0]
    m = re.match(r"^\d{4}年[\d\-]+[期月]\s*(.+?)(?:_|$)", base)
    if not m:
        return None
    cand = m.group(1).strip(" -_")
    if len(cand) < 4 or ("公司" not in cand and "集团" not in cand):
        return None                       # 看着不像公司名就不猜
    return cand


def resolve_company(preview, file_name, known):
    """返回 (公司名, 来源) 或 (None, 失败原因)。绝不猜：对不上对照表就停。"""
    for src, val in (("表内标题行", company_from_preview(preview)),
                     ("文件名", company_from_name(file_name))):
        if not val:
            continue
        if known and val not in known:
            near = [k for k in known if kmatch(val, k)]
            return None, ("%s取到「%s」，对照表里没有（相近：%s）"
                          % (src, val, "、".join(near[:3]) or "无"))
        return val, src
    return None, "表内没有「公司名称：」标题行，文件名也取不到公司全称"


def kmatch(a, b):
    return a in b or b in a


def fill_placeholders(steps, mapping, company):
    def sub(v):
        return PLACEHOLDER.sub(company, v) if isinstance(v, str) else v
    new_steps = []
    for s in steps or []:
        s2 = json.loads(json.dumps(s))
        p = s2.get("params")
        if isinstance(p, dict):
            for k in list(p):
                p[k] = sub(p[k])
        new_steps.append(s2)
    return new_steps, {k: sub(v) for k, v in (mapping or {}).items()}


def has_placeholder(obj):
    return bool(PLACEHOLDER.search(json.dumps(obj or [], ensure_ascii=False)))


# ----------------------------------------------------------------- HTTP
class Api:
    def __init__(self, base, timeout=3600):
        self.base = base.rstrip("/")
        self.timeout = timeout
        self.op = urllib.request.build_opener(urllib.request.ProxyHandler({}))

    def get(self, path, timeout=None):
        req = urllib.request.Request(self.base + path)
        return json.loads(self.op.open(req, timeout=timeout or self.timeout).read())

    def post(self, path, obj=None, timeout=None):
        req = urllib.request.Request(
            self.base + path, data=json.dumps(obj or {}).encode(),
            method="POST", headers={"Content-Type": "application/json"})
        try:
            return json.loads(self.op.open(req, timeout=timeout or self.timeout).read())
        except urllib.error.HTTPError as e:
            raw = e.read().decode("utf-8", "replace")
            try:
                body = json.loads(raw)
            except Exception:
                body = {"error": raw[:300]}
            body["_status"] = e.code
            return body

    def upload(self, path, timeout=1800):
        name = os.path.basename(path)
        size = os.path.getsize(path)
        req = urllib.request.Request(
            self.base + "/api/upload", data=open(path, "rb"), method="POST",
            headers={"X-Filename": urllib.parse.quote(name), "Content-Length": str(size)})
        try:
            return json.loads(self.op.open(req, timeout=timeout).read())
        except urllib.error.HTTPError as e:
            return {"_status": e.code, "error": e.read().decode("utf-8", "replace")[:300]}


# ----------------------------------------------------------------- 服务
class Server:
    def __init__(self, py, script, data_dir, port, log_path):
        self.py, self.script = py, script
        self.data_dir, self.port, self.log_path = data_dir, port, log_path
        self.proc = None
        self.log = None

    def start(self, wait=180):
        port = free_port(self.port)
        if not port:
            raise RuntimeError("端口 %d~%d 全被占用" % (self.port, self.port + 9))
        self.port = port
        env = dict(os.environ)
        env.update({"KZ_DATA_DIR": self.data_dir, "KZ_PORT": str(port),
                    "KZ_NO_BROWSER": "1", "no_proxy": "127.0.0.1", "NO_PROXY": "127.0.0.1",
                    "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"})
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)
        self.log = open(self.log_path, "w", encoding="utf-8", errors="replace")
        flags = subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
        self.proc = subprocess.Popen([self.py, self.script], cwd=ROOT, env=env,
                                     stdout=self.log, stderr=subprocess.STDOUT,
                                     creationflags=flags)
        api = Api("http://127.0.0.1:%d" % port, timeout=5)
        t0 = time.time()
        while time.time() - t0 < wait:
            if self.proc.poll() is not None:
                raise RuntimeError("服务进程退出（code=%s），见 %s" % (self.proc.returncode, self.log_path))
            try:
                if api.get("/api/ping", timeout=2).get("ok"):
                    time.sleep(0.6)          # 等 ledger 初始化完成
                    return
            except Exception:
                pass
            time.sleep(0.5)
        raise RuntimeError("服务 %d 秒内未就绪，见 %s" % (wait, self.log_path))

    def stop(self, wait=25):
        if not self.proc or self.proc.poll() is not None:
            return
        try:
            self.proc.terminate()
        except Exception:
            pass
        t0 = time.time()
        while time.time() - t0 < wait and self.proc.poll() is None:
            time.sleep(0.4)
        if self.proc.poll() is None:
            try:
                self.proc.kill()
            except Exception:
                pass
        if self.log:
            try:
                self.log.close()
            except Exception:
                pass


# ----------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(description="按归档配方跑完整清洗与六组检查")
    ap.add_argument("--data-root", required=True, help="账表所在目录（递归找 xlsx）")
    ap.add_argument("--recipe", default=DEFAULT_RECIPE)
    ap.add_argument("--std-accounts", default=DEFAULT_STD)
    ap.add_argument("--period-from", default=None)
    ap.add_argument("--period-to", default=None)
    ap.add_argument("--port", type=int, default=8899)
    ap.add_argument("--work-dir", default=None)
    ap.add_argument("--out-dir", default=None)
    ap.add_argument("--python", default=os.path.join(ROOT, "程序", "runtime", "python", "python.exe"))
    ap.add_argument("--server", default=os.path.join(ROOT, "程序", "server.py"))
    ap.add_argument("--only", default="", help="逗号分隔，只跑文件名含这些字样的表")
    ap.add_argument("--force-merge", action="store_true", help="排障用，默认不带病合并")
    ap.add_argument("--reuse", action="store_true", help="允许复用已有数据目录")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    ts = time.strftime("%Y%m%d_%H%M%S")
    work = os.path.abspath(args.work_dir or os.path.join(ROOT, "tmp", "复跑_" + ts))
    out = os.path.abspath(args.out_dir or os.path.join(work, "out"))
    data = os.path.join(work, "data")
    os.makedirs(out, exist_ok=True)

    log_lines = []

    def say(msg=""):
        print(msg, flush=True)
        log_lines.append(str(msg))

    def die(code, msg):
        say("!! " + msg)
        with open(os.path.join(out, "console.log"), "w", encoding="utf-8") as fh:
            fh.write("\n".join(log_lines))
        sys.exit(code)

    # ---------- S0 前置 ----------
    if not os.path.isdir(args.data_root):
        die(2, "数据目录不存在：%s" % args.data_root)
    if not os.path.isfile(args.recipe):
        die(2, "配方不存在：%s" % args.recipe)
    if not os.path.isfile(args.std_accounts):
        die(2, "客户科目表不存在：%s" % args.std_accounts)
    recipe = jload(args.recipe)

    # ---------- S1 工作目录 + 科目表（必须在任何 clean/apply 之前）----------
    if args.reuse:
        os.makedirs(data, exist_ok=True)
    else:
        if os.path.exists(os.path.join(data, "ledger.duckdb")):
            die(2, "数据目录里已有账套：%s（要复用加 --reuse，要重跑换个 --work-dir）" % data)
        os.makedirs(data, exist_ok=True)
    shutil.copyfile(args.std_accounts, os.path.join(data, "std_accounts.json"))

    files = []
    for base, _, fs in os.walk(args.data_root):
        for f in fs:
            if f.lower().endswith((".xlsx", ".xlsm", ".xls")) and not f.startswith("~$"):
                files.append(os.path.join(base, f))
    files.sort()
    only = [x.strip() for x in args.only.split(",") if x.strip()]
    if only:
        files = [f for f in files if any(k in os.path.basename(f) for k in only)]
    if not files:
        die(2, "数据目录里没有可导入的表格：%s" % args.data_root)
    books_files = [f for f in files if "mapping" in os.path.basename(f).lower()]
    ledger_files = [f for f in files if f not in books_files]

    period = recipe.get("recon", {}).get("_this_period") or {}
    pfrom = args.period_from or period.get("period_from")
    pto = args.period_to or period.get("period_to")
    if not (pfrom and pto):
        die(2, "配方里没有 _this_period，请用 --period-from/--period-to 指定")

    say("=" * 78)
    say("数据目录 : %s（%d 张表 + %d 个对照表文件）" % (args.data_root, len(ledger_files), len(books_files)))
    say("配方     : %s（md5 %s）" % (os.path.basename(args.recipe), md5_of(args.recipe)[:12]))
    say("科目表   : %s（%d 个客户编码，已在 clean 之前就位）"
        % (os.path.basename(args.std_accounts), len(jload(args.std_accounts).get("accounts") or {})))
    say("期间     : %s ~ %s（tb_base=ytd）" % (pfrom, pto))
    say("=" * 78)

    # ---------- S3 起服务 ----------
    srv = Server(args.python, args.server, data, args.port, os.path.join(out, "server.log"))
    report = {"run": {"started_at": ts, "data_root": args.data_root, "work_dir": work,
                      "data_dir": data, "period": {"from": pfrom, "to": pto},
                      "recipe": {"path": args.recipe, "md5": md5_of(args.recipe),
                                 "revision": recipe.get("revision")},
                      "std_accounts": {"path": args.std_accounts, "md5": md5_of(args.std_accounts)}},
              "files": [], "failures": [], "artifacts": {}}
    exit_code = 0
    try:
        srv.start()
        api = Api("http://127.0.0.1:%d" % srv.port)
        report["run"]["port"] = srv.port
        try:
            report["run"]["platform_version"] = api.get("/api/modules").get("version")
        except Exception:
            pass

        known_companies = set()
        # ---------- S4 对照表（先于逐表导入：占位符公司名要与它对账）----------
        if books_files:
            bpath = books_files[0]
            bk = recipe.get("books") or {}
            det = bk.get("detect") or {}
            up = api.upload(bpath)
            ins = api.post("/api/inspect", {"fid": up["file_id"]}, timeout=600)
            sh = (ins.get("sheets") or [{}])[0]
            prev = sh.get("preview") or []
            hr = bcol = ccol = None
            for i, r in enumerate(prev[:40], 1):
                cells = [norm(c) for c in r if norm(c)]
                books_like = [c for c in cells if (det.get("book_col_contains") or "账簿") in c]
                comp_like = [c for c in cells if c in (det.get("company_col_any_of") or ["公司名称", "全称", "公司"])]
                if books_like and comp_like:
                    hr = i
                    bcol = max(books_like, key=len)
                    ccol = comp_like[0]
                    break
            if not hr:
                report["failures"].append({"stage": "books", "file": os.path.basename(bpath),
                                           "detail": "扫不到表头（账簿列 + 公司列）"})
                say("!! 对照表扫不到表头：%s" % os.path.basename(bpath))
            else:
                api.post("/api/books/clear", {})
                br = api.post("/api/books/import", {"fid": up["file_id"], "sheet": sh["sheet"],
                                                    "header_row": hr, "book_col": bcol,
                                                    "company_col": ccol}, timeout=900)
                report["books"] = {"file": os.path.basename(bpath), "sheet": sh["sheet"],
                                   "header_row": hr, "book_col": bcol, "company_col": ccol,
                                   "rows": br.get("rows"), "total": br.get("total"),
                                   "conflicts": len(br.get("conflicts") or [])}
                say("")
                say("对照表   %s → 第 %d 行 / %s / %s → %s 条（冲突 %d）"
                    % (os.path.basename(bpath), hr, bcol, ccol, br.get("rows"),
                       len(br.get("conflicts") or [])))
                bl = api.get("/api/books/list")
                known_companies = {r[1] for r in (bl.get("rows") or []) if r[1]}
                say("对照表   已知主体 %d 个（占位符公司名要对得上这里）" % len(known_companies))


        # ---------- S5 逐表 ----------
        say("")
        say("%-6s %-44s %-16s %14s %10s" % ("阶段", "文件", "规则", "行数(前→后)", "耗时"))
        for path in ledger_files:
            name = os.path.basename(path)
            t0 = time.time()
            up = api.upload(path)
            if "_status" in up:
                report["failures"].append({"stage": "upload", "file": name, "detail": up})
                say("上传失败 %-40s %s" % (name[:40], up.get("error")))
                continue
            fid = up["file_id"]
            ins = api.post("/api/inspect", {"fid": fid, "target": "tb"}, timeout=600)
            if ins.get("_status"):
                report["failures"].append({"stage": "inspect", "file": name, "detail": ins})
                say("识别失败 %-40s %s" % (name[:40], ins.get("error")))
                continue

            picked = None
            for sh in ins.get("sheets") or []:
                prev = sh.get("preview") or []
                if not prev:
                    continue
                cands = []
                for hr in range(1, min(11, len(prev) + 1)):
                    for hrows in (1, 2):
                        cols = header_names(prev, hr, hrows)
                        if not cols:
                            continue
                        for rule in recipe.get("rules") or []:
                            if detect_ok(cols, rule.get("detect") or {}):
                                cands.append((detect_score(cols, rule.get("detect") or {}), hr, hrows, rule))
                cands.sort(key=lambda x: -x[0])
                if not cands:
                    report["failures"].append({"stage": "match", "file": name, "sheet": sh.get("sheet"),
                                               "detail": "未命中任何规则", "preview": prev[:3]})
                    say("未命中规则 %-38s %s" % (name[:38], sh.get("sheet")))
                    continue
                if len(cands) > 1 and cands[0][0] == cands[1][0] and cands[0][3]["id"] != cands[1][3]["id"]:
                    report["failures"].append({"stage": "match", "file": name, "sheet": sh.get("sheet"),
                                               "detail": "规则歧义：%s / %s" % (cands[0][3]["id"], cands[1][3]["id"])})
                    say("规则歧义 %-38s %s" % (name[:38], sh.get("sheet")))
                    continue
                if picked is None:
                    picked = (sh, cands[0])
            if not picked:
                continue
            sh, (score, hr, hrows, rule) = picked

            steps_use = rule.get("steps") or []
            mapping_use = rule.get("mapping") or {}
            if has_placeholder(steps_use) or has_placeholder(mapping_use):
                co, src = resolve_company(sh.get("preview"), name, known_companies)
                if not co:
                    report["failures"].append({"stage": "company", "file": name,
                                               "detail": "公司名占位符未解析：%s" % src})
                    say("公司名待定 %-38s %s" % (name[:38], src))
                    continue
                steps_use, mapping_use = fill_placeholders(steps_use, mapping_use, co)
                say("公司名   %-42s %s ← %s" % (name[:42], co, src))

            imp = api.post("/api/import", {"fid": fid, "sheet": sh["sheet"], "header_row": hr,
                                           "header_rows": hrows, "target": rule["target"]})
            if imp.get("_status") or not imp.get("sid"):
                report["failures"].append({"stage": "import", "file": name, "detail": imp})
                say("导入失败 %-40s %s" % (name[:40], str(imp.get("error"))[:60]))
                continue
            sid = imp["sid"]
            api.post("/api/clean/save", {"sid": sid, "steps": steps_use}, timeout=600)
            pv = api.post("/api/clean/preview", {"sid": sid, "sample": 0}, timeout=1800)
            ap = api.post("/api/clean/apply", {"sid": sid}, timeout=3600)
            mp = api.post("/api/mapping", {"sid": sid, "target": rule["target"],
                                           "mapping": mapping_use}, timeout=600)

            step_stats = []
            for s in (pv.get("steps") or []):
                step_stats.append({"type": s.get("type") or s.get("name"), "before": s.get("before"),
                                   "after": s.get("after"), "dropped": s.get("dropped"),
                                   "note": s.get("note"), "error": s.get("error")})
            after = None
            for s in reversed(step_stats):
                if s.get("after") is not None:
                    after = s["after"]
                    break
            rec = recipe.get("selfcheck", {}).get("_this_period_reference") or {}
            exp = None
            for ex in recipe.get("_examples_this_period") or []:
                if os.path.basename(ex.get("file_name") or "") == name:
                    exp = ex.get("rows_after_clean")
            row = {"file": name, "mb": round(os.path.getsize(path) / 1048576.0, 1), "fid": fid,
                   "sheet": sh["sheet"], "rule_id": rule["id"], "target": rule["target"],
                   "header_row": hr, "header_rows": hrows, "match_score": score,
                   "rows_before": pv.get("total_before"), "rows_after": after,
                   "dropped": pv.get("total_before") - after if (after is not None and pv.get("total_before") is not None) else None,
                   "steps": step_stats,
                   "mapping_keys": len(mapping_use),
                   "mapping_resp": {k: v for k, v in (mp or {}).items() if k in ("ok", "applied", "dropped")},
                   "expected_rows": exp, "expected_zero": exp == 0,
                   "elapsed_s": round(time.time() - t0, 1)}
            report["files"].append(row)
            say("%-6s %-44s %-16s %14s %9.1fs"
                % ("导入", name[:44], rule["id"],
                   "%s→%s" % (pv.get("total_before"), after), row["elapsed_s"]))

        # ---------- S6 合并 ----------
        merge_blocked = None
        for tgt in ("tb", "journal"):
            t0 = time.time()
            m = api.post("/api/merge", {"target": tgt, "force": bool(args.force_merge)}, timeout=3600)
            if m.get("_status") == 409 or m.get("blocked"):
                merge_blocked = {"target": tgt, "items": m.get("items") or [],
                                 "verdict": m.get("verdict"), "hint": m.get("hint")}
                say("")
                say("!! 合并被清洗门禁拦下（%s）：%s" % (tgt, m.get("verdict")))
                for it in (m.get("items") or [])[:6]:
                    say("   - %s | %s | 怎么修：%s" % (it.get("title"), it.get("detail"), it.get("fix")))
                break
            report.setdefault("merge", {})[tgt] = {
                "rows": m.get("rows"), "sheets": m.get("sheets"),
                "elapsed_s": round(time.time() - t0, 1),
                "warnings": len(m.get("warnings") or []),
                "missing_required": m.get("missing_required") or [],
                "forced": bool(m.get("forced"))}
            say("合并     %-8s %s 行 / %s 张表  %.1fs"
                % (tgt, m.get("rows"), m.get("sheets"), report["merge"][tgt]["elapsed_s"]))
        if merge_blocked:
            report["merge_blocked"] = merge_blocked
            die(5, "合并被拦，未做六组检查")

        # ---------- S7 六组 ----------
        t0 = time.time()
        rec = api.post("/api/recon", {"period_from": pfrom, "period_to": pto,
                                      "tol": 0.01, "tb_base": "ytd"}, timeout=3600)
        if rec.get("_status"):
            die(6, "检查接口报错：%s" % str(rec.get("error"))[:200])
        sm = rec.get("summary") or {}
        items = sm.get("items") or []
        six_green = (len(items) == 6 and not sm.get("failed") and not sm.get("bad_items"))
        report["recon"] = {"period": rec.get("period"), "tb_base": rec.get("tb_base"),
                           "elapsed_s": round(time.time() - t0, 1),
                           "summary": sm,
                           "checks": [{"no": c.get("no"), "title": c.get("title"),
                                       "ok": c.get("ok"), "total": c.get("total"),
                                       "bad": c.get("bad"), "error": c.get("error"),
                                       "notes": c.get("notes") or [],
                                       "parts": [{"title": p.get("title"), "total": p.get("total"),
                                                  "bad": p.get("bad")} for p in (c.get("parts") or [])]}
                                      for c in (rec.get("checks") or [])]}
        say("")
        say("%-4s %-34s %8s %6s  %s" % ("组", "名称", "检查量", "问题", "结论"))
        for it in items:
            say("%-4s %-34s %8s %6s  %s" % (it.get("no"), (it.get("title") or "")[:34],
                                            it.get("total"), it.get("bad"),
                                            "通过" if it.get("ok") else "不通过"))

        # ---------- S8 自查与导出 ----------
        def scalar(sql):
            r = api.post("/api/query", {"sql": sql}, timeout=1800)
            rows = r.get("rows") or []
            return rows[0][0] if rows and rows[0] else None

        checks = {}
        checks["tb_rows"] = scalar("select count(*) from tb")
        checks["tb_ytd_debit"] = float(scalar("select round(sum(coalesce(ytd_debit,0)),2) from tb") or 0)
        checks["journal_rows_total"] = scalar("select count(*) from journal")
        # period 有两种写法：海柔 '01'（只存月）、瑞云 '2026-01'（存年月）。
        # 写死 pfrom[-2:] 会让其中一类客户永远查不到行（2026-11-13 瑞云实测期间内 0）。
        pcond = ("(period between '%s' and '%s' or period between '%s' and '%s')"
                 % (pfrom[-2:], pto[-2:], pfrom, pto))
        checks["journal_rows_in_period"] = scalar(
            "select count(*) from journal where " + pcond)
        checks["journal_debit_in_period"] = float(scalar(
            "select round(sum(coalesce(debit,0)),2) from journal where " + pcond) or 0)
        checks["companies"] = scalar("select count(distinct company_name) from tb")
        checks["bad_code_acct_l1"] = scalar(
            "select count(distinct acct_l1) from journal where acct_l1 ~ '^[0-9]+$'")

        ref = (recipe.get("selfcheck") or {}).get("_this_period_reference") or {}
        cmp_rows = []
        for k in ("tb_rows", "tb_ytd_debit", "journal_rows_total", "journal_rows_in_period",
                  "journal_debit_in_period", "companies"):
            if k not in ref:
                continue
            a, b = checks.get(k), ref.get(k)
            same = (abs(float(a) - float(b)) < 0.01) if isinstance(a, (int, float)) else a == b
            cmp_rows.append({"name": k, "actual": a, "reference": b, "same": bool(same)})
        report["verify"] = {"six_green": bool(six_green), "checks": checks,
                            "ref_compare": cmp_rows,
                            "bad_code_acct_l1": checks.get("bad_code_acct_l1")}

        try:
            ex = api.post("/api/recon/export", {"name": "复跑_数据检查_%s_%s" % (pfrom, pto)}, timeout=1800)
            src = ex.get("path")
            if src and os.path.isfile(src):
                dst = os.path.join(out, os.path.basename(src))
                shutil.copyfile(src, dst)
                report["artifacts"]["recon_xlsx"] = dst
        except Exception as e:
            report["artifacts"]["recon_xlsx"] = None
            report["artifacts"]["export_error"] = str(e)[:200]

        say("")
        say("自查：tb %s 行 / 本年借方 %s | journal 全量 %s / 期间内 %s | 主体 %s | 一级科目纯数字 %s"
            % (checks["tb_rows"], checks["tb_ytd_debit"], checks["journal_rows_total"],
               checks["journal_rows_in_period"], checks["companies"], checks["bad_code_acct_l1"]))
        for c in cmp_rows:
            say("比对：%-24s 本次 %-18s 参考 %-18s %s"
                % (c["name"], c["actual"], c["reference"], "一致" if c["same"] else "≠"))
        say("")
        say("六组：%s" % ("6/6 通过，bad 合计 0 ✅" if six_green else "未全绿 ❌"))

        if report["failures"]:
            exit_code = 4
        elif not six_green:
            exit_code = 6

    except RuntimeError as e:
        report["fatal"] = str(e)
        say("!! " + str(e))
        exit_code = 3
    finally:
        srv.stop()
        lock = os.path.join(data, "server.lock")
        if os.path.exists(lock):
            try:
                os.remove(lock)
            except OSError:
                pass
        report["run"]["finished_at"] = time.strftime("%Y%m%d_%H%M%S")
        report["exit_code"] = exit_code
        report["artifacts"]["console_log"] = os.path.join(out, "console.log")
        report["artifacts"]["server_log"] = os.path.join(out, "server.log")
        with open(os.path.join(out, "report.json"), "w", encoding="utf-8") as fh:
            json.dump(report, fh, ensure_ascii=False, indent=2)
        report["artifacts"]["report_json"] = os.path.join(out, "report.json")
        with open(os.path.join(out, "console.log"), "w", encoding="utf-8") as fh:
            fh.write("\n".join(log_lines))
        say("")
        say("报告：%s" % report["artifacts"]["report_json"])
        say("数据目录（跑完可删）：%s" % data)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
