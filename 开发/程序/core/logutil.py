# -*- coding: utf-8 -*-
"""日志：一份写文件、一份留在内存供界面查看/下载。

为什么不用裸 print：打包成 exe（无控制台）后，print 直接丢失——出问题时
用户除了「没反应」什么也给不出来。这里做三件事：
  1. 接管 stdout/stderr（业务代码里几十处 print 不用改，全部落到日志文件）
  2. 捕获未处理异常与线程异常（崩溃也有记录）
  3. 内存里留最近 N 条，前端「日志」页可直接看 + 一键下载 zip

文件策略：按天一个 app-YYYYMMDD.log，保留最近 14 天；单文件超 20MB 自动换新。
"""
import io
import os
import sys
import time
import zipfile
import threading
import traceback
from collections import deque

from core import paths

MAX_MEM = 3000           # 内存里保留的条数
KEEP_DAYS = 14
MAX_FILE = 20 * 1024 * 1024

_LOCK = threading.RLock()
_RING = deque(maxlen=MAX_MEM)
_STATE = {"installed": False, "file": None, "day": None, "size": 0, "level": "INFO"}
_LEVELS = {"DEBUG": 10, "INFO": 20, "WARN": 30, "ERROR": 40}


def log_dir():
    d = paths.LOG_DIR
    os.makedirs(d, exist_ok=True)
    return d


def _today():
    return time.strftime("%Y%m%d")


def _current_file():
    """当天日志文件；跨天或超限时自动换新。"""
    day = _today()
    if _STATE["file"] and _STATE["day"] == day and _STATE["size"] < MAX_FILE:
        return _STATE["file"]
    d = log_dir()
    base = os.path.join(d, "app-%s.log" % day)
    path, i = base, 1
    while os.path.exists(path) and os.path.getsize(path) >= MAX_FILE:
        path = os.path.join(d, "app-%s-%d.log" % (day, i))
        i += 1
    _STATE["file"], _STATE["day"] = path, day
    try:
        _STATE["size"] = os.path.getsize(path)
    except OSError:
        _STATE["size"] = 0
    return path


def write(level, msg, tag=""):
    """写一条日志。level: DEBUG/INFO/WARN/ERROR。"""
    level = str(level or "INFO").upper()
    if _LEVELS.get(level, 20) < _LEVELS.get(_STATE["level"], 20):
        return
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = "%s [%-5s]%s %s" % (ts, level, (" " + tag) if tag else "", str(msg))
    with _LOCK:
        _RING.append({"ts": ts, "level": level, "tag": tag, "msg": str(msg)})
        try:
            p = _current_file()
            with io.open(p, "a", encoding="utf-8") as f:
                f.write(line + "\n")
            _STATE["size"] += len(line.encode("utf-8")) + 1
        except Exception:
            pass          # 日志写不进去不能影响业务
    return line


def info(msg, tag=""):
    return write("INFO", msg, tag)


def warn(msg, tag=""):
    return write("WARN", msg, tag)


def error(msg, tag=""):
    return write("ERROR", msg, tag)


def exception(msg, tag=""):
    return write("ERROR", "%s\n%s" % (msg, traceback.format_exc()), tag)


class _Tee:
    """替代 sys.stdout/stderr：既进日志，也在有控制台时照常打印。"""

    def __init__(self, stream, level, tag):
        self._s = stream
        self._level = level
        self._tag = tag

    def write(self, s):
        if not s:
            return 0
        t = s.rstrip("\n")
        if t:
            write(self._level, t, self._tag)
        if self._s is not None:
            try:
                self._s.write(s)
            except Exception:
                pass
        return len(s)

    def flush(self):
        if self._s is not None:
            try:
                self._s.flush()
            except Exception:
                pass

    def isatty(self):
        return False

    @property
    def encoding(self):
        return "utf-8"


def install(level="INFO"):
    """进程启动时调用一次。"""
    if _STATE["installed"]:
        return
    _STATE["installed"] = True
    _STATE["level"] = str(level or "INFO").upper()

    # 打包窗口模式（console=False）下 stdout 可能是 None，Tee 已处理
    try:
        sys.stdout = _Tee(sys.stdout, "INFO", "out")
        sys.stderr = _Tee(sys.stderr, "ERROR", "err")
    except Exception:
        pass

    def _hook(exc_type, exc, tb):
        write("ERROR", "未捕获异常：%s" % "".join(traceback.format_exception(exc_type, exc, tb)),
              "crash")

    sys.excepthook = _hook
    try:
        threading.excepthook = lambda a: write(
            "ERROR", "线程异常 %s：%s" % (getattr(a, "thread", None),
                                      "".join(traceback.format_exception(
                                          a.exc_type, a.exc_value, a.exc_traceback))), "crash")
    except Exception:
        pass

    info("=" * 60, "boot")
    info(paths.describe(), "boot")
    info("Python %s | PID %d" % (sys.version.split()[0], os.getpid()), "boot")


def tail(n=300, level=None):
    """最近 n 条（新的在后），可按级别过滤。"""
    with _LOCK:
        items = list(_RING)
    if level:
        need = _LEVELS.get(str(level).upper(), 0)
        items = [x for x in items if _LEVELS.get(x["level"], 20) >= need]
    return items[-int(n or 300):]


def files():
    """日志文件清单（新→旧）。"""
    d = log_dir()
    out = []
    for f in sorted(os.listdir(d), reverse=True):
        if not f.endswith(".log"):
            continue
        p = os.path.join(d, f)
        try:
            st = os.stat(p)
            out.append({"name": f, "size": st.st_size,
                        "mtime": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_mtime))})
        except OSError:
            pass
    return out


def bundle(include_settings=False):
    """把当前 + 历史日志（可选设置）打成一个 zip 字节流，供界面下载。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        d = log_dir()
        for f in sorted(os.listdir(d)):
            if f.endswith(".log"):
                try:
                    z.write(os.path.join(d, f), "logs/" + f)
                except OSError:
                    pass
        # 附一份运行环境，排查时省一轮问答
        env = ["%s=%s" % (k, os.environ.get(k, ""))
               for k in ("KZ_PORT", "KZ_DATA_DIR", "OS", "PROCESSOR_ARCHITECTURE")]
        env.append(paths.describe())
        env.append("time=%s" % time.strftime("%Y-%m-%d %H:%M:%S"))
        z.writestr("env.txt", "\n".join(env))
        if include_settings:
            import json
            sp = os.path.join(paths.DATA_DIR, "settings.json")
            if os.path.exists(sp):
                try:
                    cfg = json.load(io.open(sp, encoding="utf-8"))
                    # 脱敏：任何 key/secret 一律打码，避免用户把 key 传给第三方
                    for k in list(cfg.keys()):
                        if any(s in k.lower() for s in ("key", "token", "secret", "password")):
                            cfg[k] = "***"
                    z.writestr("settings.json", json.dumps(cfg, ensure_ascii=False, indent=2))
                except Exception:
                    pass
    return buf.getvalue()


def purge_old(days=KEEP_DAYS):
    """删掉超期日志（默认 14 天）。"""
    cutoff = time.time() - days * 86400
    d = log_dir()
    n = 0
    for f in os.listdir(d):
        if not f.endswith(".log"):
            continue
        p = os.path.join(d, f)
        try:
            if os.path.getmtime(p) < cutoff:
                os.remove(p)
                n += 1
        except OSError:
            pass
    return n
