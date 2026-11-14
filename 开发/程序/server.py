# -*- coding: utf-8 -*-
"""PTW 平台 · 本地服务入口

只用标准库 + duckdb，不依赖 pip 安装。启动后自动打开浏览器。

本文件只做「外壳」：扫描 modules/*/manifest.json 自动装配各板块，
新增功能只需在 modules/ 下加一个目录，不必改这里。
"""
import importlib
import json
import mimetypes
import os
import re
import socket
import sys
import threading
import time
import traceback
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

from core import paths          # noqa: E402  统一路径（兼容 exe 打包）
from core import logutil        # noqa: E402  日志：落文件 + 界面可下载
from core import uiwindow       # noqa: E402  界面以独立窗口打开（Edge/Chrome --app）
from core.db import Ledger  # noqa: E402

try:
    from app_version import APP_VERSION as VERSION
except Exception:               # 源码运行时没有该模块，回退
    VERSION = "1.2.122"
APP_NAME = "PTW平台"

#: 所有 /api/ 请求共用同一个 DuckDB 连接（Ledger.conn），而 ThreadingHTTPServer
#: 每个请求开一个线程。两个请求撞在一起时 DuckDB 直接抛
#: "connection is busy with another call"，表现是随机某个按钮报错。
#: 本地单用户工具没有并发吞吐的需求，串行化是最省事也最稳的做法。
_API_LOCK = threading.RLock()
#: 最近一次收到请求的时间。更新重启后用它判断「还有没有旧窗口连回来」——
#: 有就不必再开一个新窗口，没有（用户把窗口关了）才补开，避免出现两个窗口。
_LAST_REQ = [0.0]

MODULES_DIR = os.path.join(APP_DIR, "modules")


def load_modules():
    """扫描 modules/*/manifest.json，自动装配板块。

    manifest.json 字段（除 key 外都可省）：
        key      板块标识，同时是静态资源前缀 /<key>/，默认取目录名
        label    左侧导航分组名
        order    排序，越小越前
        type     native（前端在 app.js 里有对应视图）| embed（iframe 独立页）
        entry    type=embed 时的入口文件，默认 index.html
        static   静态资源子目录（如 web），相对板块目录
        children type=native 时的子项 [{k, no, label}]
        icon     导航图标字符

    routes.py 若存在则导入，读取其 GET_ROUTES / POST_ROUTES（路径自带 /api 前缀，
    这里不做任何改写，保证老接口地址不变）。
    """
    mods, get_routes, post_routes, static_map = [], {}, {}, {}
    if not os.path.isdir(MODULES_DIR):
        return mods, get_routes, post_routes, static_map

    for name in sorted(os.listdir(MODULES_DIR)):
        mdir = os.path.join(MODULES_DIR, name)
        mf = os.path.join(mdir, "manifest.json")
        if not (os.path.isdir(mdir) and os.path.isfile(mf)):
            continue
        try:
            with open(mf, encoding="utf-8") as f:
                man = json.load(f)
        except Exception as e:
            sys.stderr.write("[板块] %s/manifest.json 解析失败，已跳过：%s\n" % (name, e))
            continue

        man["key"] = man.get("key") or name
        man.setdefault("label", man["key"])
        man.setdefault("order", 100)
        man.setdefault("type", "native")
        man.setdefault("children", [])

        rp = os.path.join(mdir, "routes.py")
        if os.path.isfile(rp):
            try:
                mod = importlib.import_module("modules.%s.routes" % name)
                get_routes.update(getattr(mod, "GET_ROUTES", None) or {})
                post_routes.update(getattr(mod, "POST_ROUTES", None) or {})
            except Exception as e:
                traceback.print_exc()
                man["error"] = "路由加载失败：%s" % e

        sub = man.get("static")
        if sub:
            sd = os.path.join(mdir, sub.replace("/", os.sep))
            if os.path.isdir(sd):
                static_map[man["key"]] = [sd, man.get("entry") or "index.html"]

        mods.append(man)

    mods.sort(key=lambda m: (m.get("order", 100), m["key"]))
    return mods, get_routes, post_routes, static_map


def _api_modules(ctx, q):
    return ctx.json({"app": APP_NAME, "version": VERSION, "modules": MODULES})


MODULES, GET_ROUTES, POST_ROUTES, MODULE_STATIC = load_modules()

# 平台自身接口（不属于任何板块）：前端据此渲染导航
GET_ROUTES.setdefault("/api/modules", _api_modules)


# ---------------------------------------------------------------- 日志接口
# 打包成 exe 后没有控制台，用户遇到问题无法提供现场。这组接口让界面能
# 直接看日志、下载日志包（含环境信息），发回来就能定位。
def _api_logs_tail(ctx, q):
    n = int((q.get("n") or ["300"])[0] or 300)
    lv = (q.get("level") or [""])[0] or None
    return ctx.json({"items": logutil.tail(n, lv), "files": logutil.files(),
                     "dir": logutil.log_dir(), "paths": paths.describe()})


def _api_logs_download(ctx, q):
    """下载日志 zip（当前 + 历史 + env.txt + 脱敏后的 settings.json）。"""
    data = logutil.bundle(include_settings=True)
    name = "PTW日志_%s.zip" % time.strftime("%Y%m%d_%H%M%S")
    ctx.h.send_response(200)
    ctx.h.send_header("Content-Type", "application/zip")
    ctx.h.send_header("Content-Length", str(len(data)))
    ctx.h.send_header("Content-Disposition",
                      "attachment; filename*=UTF-8''" + urllib.parse.quote(name))
    ctx.h.send_header("Cache-Control", "no-store")
    ctx.h.end_headers()
    ctx.h.wfile.write(data)
    logutil.info("用户下载了日志包（%d 字节）" % len(data), "ui")


def _api_logs_clear(ctx, q):
    """清空内存缓冲（文件不动，留作证据）。"""
    logutil._RING.clear()
    logutil.info("用户清空了界面日志缓冲", "ui")
    return ctx.json({"ok": True})


GET_ROUTES.setdefault("/api/logs", _api_logs_tail)
GET_ROUTES.setdefault("/api/logs/download", _api_logs_download)
POST_ROUTES.setdefault("/api/logs/clear", _api_logs_clear)


def _api_shutdown(ctx, u, body):
    """界面上的「退出」按钮：优雅停机（走 finally 的 CHECKPOINT，不硬杀进程）。"""
    srv = getattr(Handler, "srv", None)
    if srv is None:
        return ctx.json({"ok": False, "error": "当前不是由本进程托管的服务，无法从这里退出"})
    logutil.info("用户从界面点了「退出」", "ui")

    def _stop():
        try:
            srv.shutdown()          # 让 serve_forever 返回，主线程继续走 finally
        except Exception:
            os._exit(0)

    threading.Timer(0.3, _stop).start()
    return ctx.json({"ok": True, "note": "正在退出，窗口可以关闭了"})


POST_ROUTES.setdefault("/api/app/quit", _api_shutdown)


# ---------------------------------------------------------------- AI 接口（内置 B2）
# 与外部 AI 旁路完全等价：结果写同一批 *_result.json，两边可互换。
def _api_ai_status(ctx, q):
    from core import ptw_ai
    return ctx.json(ptw_ai.status())


def _api_ai_save(ctx, u, body):
    from core import ptw_ai
    return ctx.json(ptw_ai.save_settings(body or {}))


def _api_ai_test(ctx, u, body):
    from core import ptw_ai
    return ctx.json(ptw_ai.test())


def _api_ai_run(ctx, u, body):
    from core import ptw_ai
    task = (body or {}).get("task") or ""
    limit = (body or {}).get("limit") or 60
    try:
        return ctx.json(ptw_ai.run_task(task, limit))
    except ptw_ai.AIError as e:
        return ctx.json({"ok": False, "kind": e.kind, "error": e.msg,
                         "detail": e.detail[:300]}, 400)


GET_ROUTES.setdefault("/api/ai/status", _api_ai_status)
POST_ROUTES.setdefault("/api/ai/save", _api_ai_save)
POST_ROUTES.setdefault("/api/ai/test", _api_ai_test)
POST_ROUTES.setdefault("/api/ai/run", _api_ai_run)


# ---------------------------------------------------------------- 更新接口
def _api_update_check(ctx, q):
    from core import updater
    try:
        info, attempts = updater.check()
        return ctx.json({"ok": True, "current": VERSION, "info": info,
                         "attempts": attempts, "root": updater.root_dir()})
    except updater.UpdateError as e:
        return ctx.json({"ok": False, "current": VERSION, "error": e.msg,
                         "attempts": e.attempts})


#: 更新任务的进度状态。下载 35MB 在慢网上要几分钟，绝不能同步阻塞请求
#: （API 是串行加锁的，一旦阻塞，整个界面都会卡死）。
UPDS = {"phase": "idle", "got": 0, "total": 0, "version": "", "error": "", "note": ""}
_UPDS_LOCK = threading.Lock()
UPD_BUSY = ("checking", "downloading", "staging", "restarting")


def _api_update_progress(ctx, q):
    return ctx.json(dict(UPDS))


def _do_update():
    from core import updater
    try:
        info, _ = updater.check()
        if not info.get("is_new"):
            UPDS.update({"phase": "idle", "error": "",
                         "note": "当前已是最新版本（%s）" % VERSION})
            return
        UPDS.update({"phase": "downloading", "version": info["version"],
                     "got": 0, "total": info.get("size") or 0, "note": ""})

        def prog(got, total):
            UPDS["got"] = got
            if total:
                UPDS["total"] = total

        stage, zp = updater.prepare(info, prog)
        UPDS.update({"phase": "staging",
                     "note": "已下载，正在解压到 %s" % os.path.basename(stage)})
        logutil.info("更新已下载并解压到 %s（%s → %s）"
                     % (stage, VERSION, info["version"]), "update")
        UPDS.update({"phase": "restarting", "version": info["version"],
                     "note": "正在替换程序并重启"})
        mode = updater.launch_apply()
        # launcher 正在等我退出：立刻退；否则是刚拉起的 launcher，给它 3 秒启动
        delay = 0.3 if mode == "launcher-waiting" else 3.0
        threading.Timer(delay, lambda: os._exit(0)).start()
    except updater.UpdateError as e:
        UPDS.update({"phase": "failed", "error": e.msg})
        logutil.write("ERROR", "更新失败：%s" % e.msg, "update")
    except Exception as e:
        UPDS.update({"phase": "failed", "error": str(e)[:200]})
        logutil.write("ERROR", "更新失败：%s\n%s" % (e, traceback.format_exc()), "update")


def _api_update_apply(ctx, u, body):
    """立刻返回，后台线程干活；前端轮询 /api/update/progress 看进度。"""
    with _UPDS_LOCK:
        if UPDS.get("phase") in UPD_BUSY:
            return ctx.json({"ok": True, "phase": UPDS["phase"],
                             "note": "正在更新，请勿关闭窗口"})
        UPDS.update({"phase": "checking", "got": 0, "total": 0,
                     "version": "", "error": "", "note": "检查新版本…"})
    threading.Thread(target=_do_update, daemon=True).start()
    return ctx.json({"ok": True, "phase": "checking", "note": "开始更新"})


GET_ROUTES.setdefault("/api/update/check", _api_update_check)
POST_ROUTES.setdefault("/api/update/apply", _api_update_apply)
GET_ROUTES.setdefault("/api/update/progress", _api_update_progress)



def free_port(start=8787):
    for p in range(start, start + 40):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start


class _Ctx:
    """板块路由的上下文：把 Handler 的能力以简洁接口暴露给各模块。"""

    def __init__(self, handler):
        self.h = handler

    @property
    def ledger(self):
        return self.h.ledger

    @property
    def web_dir(self):
        return self.h.web_dir

    @property
    def headers(self):
        return self.h.headers

    @property
    def rfile(self):
        return self.h.rfile

    @property
    def query_string(self):
        return self.h.query_string

    @property
    def content_type(self):
        return self.h.headers.get("Content-Type", "")

    def json(self, obj, code=200):
        return self.h._json(obj, code)

    def err(self, msg, code=400):
        return self.h._err(msg, code)

    def send_file(self, path, ctype, fname=None):
        return self.h._send_file(path, ctype, fname)

    def body_bytes(self):
        return self.h._body()

    def multipart(self):
        """解析 multipart/form-data 请求体，返回 {'files': {name:{filename,data}}, 'fields': {name:str}}。"""
        return self.h._multipart()


class Handler(BaseHTTPRequestHandler):
    server_version = "PTW/" + VERSION
    protocol_version = "HTTP/1.1"
    ledger = None
    web_dir = None
    module_static = {}          # {板块 key: [目录, 入口文件]}

    # ---------------------------------------------------------- 基础
    def log_message(self, fmt, *args):
        if os.environ.get("KZ_VERBOSE"):
            sys.stderr.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _err(self, msg, code=400):
        self._json({"error": str(msg)}, code)

    def _send_file(self, path, ctype, fname=None):
        """发送本地文件（板块模块用，如 BKD 模板下载）。"""
        if not os.path.exists(path):
            return self._err("文件不存在", 404)
        size = os.path.getsize(path)
        fname = fname or os.path.basename(path)
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(size))
        self.send_header("Content-Disposition",
                         "attachment; filename*=UTF-8''" + urllib.parse.quote(fname))
        self.end_headers()
        with open(path, "rb") as f:
            while True:
                b = f.read(1 << 20)
                if not b:
                    break
                self.wfile.write(b)

    def _ctx(self):
        """构造传给板块路由的上下文对象。"""
        return _Ctx(self)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0:
            return b""
        buf = []
        while n > 0:
            b = self.rfile.read(min(n, 4 << 20))
            if not b:
                break
            buf.append(b)
            n -= len(b)
        return b"".join(buf)

    def _payload(self):
        raw = self._body()
        if not raw:
            return {}
        try:
            return json.loads(raw.decode("utf-8"))
        except Exception:
            raise ValueError("请求体不是合法 JSON")

    def _multipart(self):
        """解析 multipart/form-data 请求体（BKD 底稿生成板块用）。"""
        raw = self._body()
        ctype = self.headers.get("Content-Type", "")
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

    # ---------------------------------------------------------- 路由
    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        p = u.path
        q = urllib.parse.parse_qs(u.query)
        t0 = time.time()
        _LAST_REQ[0] = t0
        try:
            if p.startswith("/api/"):
                with _API_LOCK:
                    r = self._api_get(p, q)
                if p != "/api/logs":          # 日志轮询不刷屏
                    logutil.info("GET %s %.0fms" % (p, (time.time() - t0) * 1000), "req")
                return r
            if p == "/download":
                return self._download(q.get("f", [""])[0])
            return self._static(p)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as e:
            logutil.write("ERROR", "GET %s 失败：%s\n%s"
                          % (p, e, traceback.format_exc()), "req")
            traceback.print_exc()
            self._err(e, 500)

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        t0 = time.time()
        _LAST_REQ[0] = t0
        try:
            with _API_LOCK:
                r = self._api_post(u.path, u)
            logutil.info("POST %s %.0fms" % (u.path, (time.time() - t0) * 1000), "req")
            return r
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as e:
            logutil.write("ERROR", "POST %s 失败：%s\n%s"
                          % (u.path, e, traceback.format_exc()), "req")
            traceback.print_exc()
            self._err(e, 400)

    # ---------------------------------------------------------- GET
    def _api_get(self, p, q):
        fn = GET_ROUTES.get(p)
        if fn is None:
            return self._err("未知接口 " + p, 404)
        return fn(self._ctx(), q)

    # ---------------------------------------------------------- POST
    def _api_post(self, p, u):
        fn = POST_ROUTES.get(p)
        if fn is None:
            return self._err("未知接口 " + p, 404)
        ctx = self._ctx()
        # 上传接口：直接读原始字节流（模块里用 ctx.rfile / ctx.body_bytes 处理）
        if p == "/api/upload":
            return fn(ctx, u, {})
        # BKD 底稿生成的 multipart 接口：body 由模块自己用 ctx.multipart() 解析，
        # 这里传一个空占位，避免预先消费请求体。
        # 按 content-type 判断而不是列路径名单——否则每加一个上传接口都要回来改 server.py。
        if "multipart/form-data" in (self.headers.get("Content-Type") or ""):
            return fn(ctx, u, None)
        body = self._payload()
        return fn(ctx, u, body)


    # ---------------------------------------------------------- 文件
    def _download(self, fname):
        fname = os.path.basename(fname)
        # 依次在 exports（查账导出/备份/数据包）和 bkd_out（底稿生成结果）里找
        path = os.path.join(self.ledger.out_dir, fname)
        if not os.path.exists(path):
            bkd_out = os.path.join(self.ledger.data_dir, "bkd_out")
            alt = os.path.join(bkd_out, fname)
            if os.path.exists(alt):
                path = alt
            else:
                return self._err("文件不存在", 404)
        size = os.path.getsize(path)
        self.send_response(200)
        ctype = mimetypes.guess_type(fname)[0] or "application/octet-stream"
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(size))
        self.send_header("Content-Disposition",
                         "attachment; filename*=UTF-8''" + urllib.parse.quote(fname))
        self.end_headers()
        with open(path, "rb") as f:
            while True:
                b = f.read(1 << 20)
                if not b:
                    break
                self.wfile.write(b)

    def _static(self, p):
        if p in ("/", ""):
            p = "/index.html"
        rel = urllib.parse.unquote(p).lstrip("/")
        if ".." in rel:
            return self._err("bad path", 403)
        # 板块静态资源：/<key>/xxx → 该板块 manifest 声明的目录（BKD 走这条）
        base = self.web_dir
        seg = rel.split("/", 1)[0]
        if seg in self.module_static:
            base, entry = self.module_static[seg]
            rel = rel[len(seg):].lstrip("/") or entry
        path = os.path.join(base, rel.replace("/", os.sep))
        if not os.path.exists(path) or os.path.isdir(path):
            return self._err("not found: " + rel, 404)
        ctype = mimetypes.guess_type(path)[0] or "application/octet-stream"
        if path.endswith(".js"):
            ctype = "application/javascript; charset=utf-8"
        if path.endswith((".html", ".css", ".json")):
            ctype += "; charset=utf-8"
        data = open(path, "rb").read()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)


def _open_ui_later(url):
    """打开界面窗口。

    - KZ_NO_BROWSER=1：完全不开（自动化测试、纯服务端部署）
    - PTW_NO_UI=1：更新后重启的场景。旧窗口并不会随进程消失，它会自己轮询并刷新回来，
      所以先等 20 秒；真没人连回来（用户把窗口关了）才补开一个，避免出现两个窗口。
    """
    if os.environ.get("KZ_NO_BROWSER") == "1":
        return
    log = lambda m: print("  " + m)
    if os.environ.get("PTW_NO_UI") != "1":
        threading.Timer(0.6, lambda: uiwindow.open_ui(url, log=log)).start()
        return

    def _wait_then_maybe_open():
        deadline = time.time() + 20
        while time.time() < deadline:
            time.sleep(1)
            if _LAST_REQ[0] > 0:          # 旧窗口已自动重连
                print("  旧窗口已重新连上，不再新开窗口")
                return
        log("旧窗口没有连回来，补开一个")
        uiwindow.open_ui(url, log=log)

    threading.Thread(target=_wait_then_maybe_open, daemon=True).start()


def main():
    # 日志最先装：后面任何一步失败（端口、数据库、板块加载）都要留痕
    logutil.install(os.environ.get("KZ_LOG_LEVEL") or "INFO")
    try:
        logutil.purge_old()
    except Exception:
        pass
    try:
        _seeded = paths.seed_templates()
        if _seeded:
            logutil.info("首次释放内置模板 %d 个：%s" % (len(_seeded), "、".join(_seeded)), "boot")
    except Exception as e:
        logutil.warn("模板释放失败：%s" % e, "boot")
    port = int(os.environ.get("KZ_PORT") or 0) or free_port(int(os.environ.get("KZ_PORT_START", 8787)))
    data_dir = paths.DATA_DIR
    lock_path = os.path.join(data_dir, "server.lock")

    # ---- 单实例保护：已有实例在跑时，读它的端口直接开浏览器，不再起第二个进程
    def _pid_alive(pid):
        try:
            if not pid:
                return False
            import ctypes
            h = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))
            if not h:
                return False
            ctypes.windll.kernel32.CloseHandle(h)
            return True
        except Exception:
            # 查不了就当"没在跑"：宁可多起一次、让端口占用来报错，
            # 也不要因为一次查询失败就拒绝对用户可见的启动。
            return False

    def _port_open(p):
        """端口能不能连上。锁文件是在 bind 成功后才写的，
        所以"PID 活着但端口连不上"就等于锁是上次崩溃留下的。"""
        import socket
        try:
            with socket.create_connection(("127.0.0.1", int(p or 0)), timeout=0.6):
                return True
        except Exception:
            return False

    def _read_lock():
        try:
            with open(lock_path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    prev = _read_lock()
    if prev and _pid_alive(prev.get("pid")) and _port_open(prev.get("port") or port):
        url = prev.get("url") or ("http://127.0.0.1:%d/" % prev.get("port", port))
        print("=" * 62)
        print("  %s 已有实例在运行" % APP_NAME)
        print("  地址：%s（PID %s）" % (url, prev.get("pid")))
        print("  直接打开浏览器，不再启动第二个进程。")
        print("  如确认没有实例在跑，删除 %s 后重试。" % lock_path)
        print("=" * 62)
        if os.environ.get("KZ_NO_BROWSER") != "1":
            threading.Timer(0.4, lambda: uiwindow.open_ui(url, log=lambda m: print("  " + m))
                              ).start()
        return

    if prev:
        # 锁还在、但进程已退出或端口没监听：上次是关窗口、断电或任务管理器结束的。
        # 以前这里直接判"已有实例"就退出，用户只能手动删锁 —— 双击 bat 起不来就是这么来的。
        print("  检测到上次启动残留的锁（PID %s 未在监听），已清理，继续启动。"
              % prev.get("pid"))
        try:
            os.remove(lock_path)
        except Exception:
            pass

    print("=" * 62)
    print("  %s v%s · 正在启动" % (APP_NAME, VERSION))
    print("  数据目录：%s" % data_dir)
    print("  日志目录：%s" % logutil.log_dir())
    if MODULES:
        print("  已装板块：" + " / ".join(
            "%s%s" % (m.get("label", m["key"]), ("（异常：%s）" % m["error"]) if m.get("error") else "")
            for m in MODULES))
    t0 = time.time()
    try:
        ledger = Ledger(APP_DIR)
    except Exception as e:
        msg = str(e)
        if "already open" in msg or "另一个程序正在使用" in msg or "Cannot open file" in msg:
            print("  ⚠ 数据目录正被另一个实例占用（数据库文件已打开）。")
            print("    请先关掉已有的 PTW 窗口，或结束残留的 python 进程后重试。")
            print("    数据目录：%s" % data_dir)
            print("    如确认没有实例在跑，删除 %s 后重试。" % lock_path)
        else:
            print("  数据库打开失败：%s" % msg[:300])
        return
    print("  数据库就绪（%.1fs）" % (time.time() - t0))
    Handler.ledger = ledger
    Handler.web_dir = os.path.join(APP_DIR, "web")
    Handler.module_static = MODULE_STATIC
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    srv.daemon_threads = True
    url = "http://127.0.0.1:%d/" % port
    # 写锁：给下一个实例读端口用
    try:
        os.makedirs(data_dir, exist_ok=True)
        with open(lock_path, "w", encoding="utf-8") as f:
            json.dump({"pid": os.getpid(), "port": port, "url": url}, f)
    except Exception:
        pass
    print("  服务地址：%s" % url)
    print("  关闭本窗口即退出（数据会保存在 data 目录）")
    print("=" * 62)
    Handler.srv = srv          # 供「退出」接口优雅停机（要走到 finally 里的 CHECKPOINT）
    _open_ui_later(url)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出")
    finally:
        try:
            os.remove(lock_path)
        except Exception:
            pass
        try:
            ledger.conn.execute("CHECKPOINT")
            ledger.conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
