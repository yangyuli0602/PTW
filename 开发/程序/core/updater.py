# -*- coding: utf-8 -*-
"""自动更新：检查 → 下载 → 校验 → 解压暂存 → 交给 launcher 替换并重启。

与 MDMQ 的差别：MDMQ 只提示、不下载（用户手动覆盖）。PTW 要求「点『是』后自动完成」，
所以这里多了下载、校验、暂存三步；**真正的文件替换交给 launcher**（PTW.exe），
因为 Windows 不能覆盖正在运行的程序。

运行时目录约定（打包态）：
    平台根/
      PTW.exe              ← launcher，永不更新，负责检查/替换/拉起
      version.json         ← 本机版本指针（launcher 读）
      runtime/             ← 主程序目录（整体替换：PTWApp.exe + _internal + modules + web）
      data/                ← 用户数据，不参与更新
      runtime_new/         ← 下载解压后的暂存目录
      update_pending.json  ← 待应用标记（launcher 看到就替换）
"""
import hashlib
import io
import json
import os
import re
import shutil
import socket
import ssl
import subprocess
import sys
import time
import urllib.error
import urllib.request
import zipfile

from . import paths

try:
    from app_version import APP_VERSION, UPDATE_REPO, GITHUB_PAGE
except Exception:
    APP_VERSION, UPDATE_REPO, GITHUB_PAGE = "0.0.0", "", ""

TIMEOUT = 5
BUDGET = 15
MAX_JSON = 64 * 1024
CHUNK = 1 << 20
NOTES_MAX = 4000

ENDPOINTS = [
    ("raw", "https://raw.githubusercontent.com/{repo}/main/version.json"),
    ("jsdelivr", "https://cdn.jsdelivr.net/gh/{repo}@main/version.json"),
    ("fastly", "https://fastly.jsdelivr.net/gh/{repo}@main/version.json"),
    ("github-api", "https://api.github.com/repos/{repo}/releases/latest"),
]

_PRE = re.compile(r"(?:^|[-_.])(alpha|beta|rc|dev|pre|snapshot)", re.I)


class UpdateError(Exception):
    def __init__(self, msg, attempts=None):
        super().__init__(msg)
        self.msg = msg
        self.attempts = attempts or []


#: 平台根 = runtime 的上一级（打包态）；源码态下等于 data 的上一级
def root_dir():
    if paths.FROZEN:
        return os.path.dirname(paths.EXE_DIR)      # runtime/.. = 平台根
    return os.path.abspath(os.path.join(paths.DATA_DIR, ".."))


def staging_dir():
    return os.path.join(root_dir(), "runtime_new")


def pending_file():
    return os.path.join(root_dir(), "update_pending.json")


def launcher_path():
    return os.path.join(root_dir(), "PTW.exe")


# ---------------------------------------------------------------- 版本
def ver_key(s):
    s = str(s or "").strip().lstrip("vV")
    base = re.split(r"[-+]", s, 1)[0]
    nums = re.findall(r"\d+", base)[:4]
    return tuple(int(n) for n in nums) or (0,)


def ver_cmp(a, b):
    ka, kb = ver_key(a), ver_key(b)
    n = max(len(ka), len(kb))
    ka, kb = ka + (0,) * (n - len(ka)), kb + (0,) * (n - len(kb))
    if ka != kb:
        return 1 if ka > kb else -1
    pa, pb = bool(_PRE.search(str(a))), bool(_PRE.search(str(b)))
    return (not pa) - (not pb)


# ---------------------------------------------------------------- 检查
def _ssl_ctx():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


_CTX = _ssl_ctx()


def _get(url, timeout=TIMEOUT, limit=MAX_JSON):
    req = urllib.request.Request(url, headers={
        "User-Agent": "PTW/%s" % APP_VERSION,
        "Accept": "application/json, text/plain, */*", "Cache-Control": "no-cache"})
    with urllib.request.urlopen(req, timeout=timeout, context=_CTX) as r:
        return r.read(limit)


def _get_any(url):
    """本地路径 / file: / http 回环 —— 内网与自测用（不走 SSL 上下文）。"""
    if url.lower().startswith("file:") or os.path.exists(url):
        p = url[7:] if url.lower().startswith("file:") else url
        with open(p, "rb") as f:
            return f.read(MAX_JSON)
    with urllib.request.urlopen(url, timeout=TIMEOUT) as r:
        return r.read(MAX_JSON)


def _allowed_url(u):
    """只放行 https 与本地回环 http（内网更新服务器）；其他一律拒绝。"""
    s = str(u or "").lower()
    return (s.startswith("https://") or s.startswith("http://127.0.0.1")
            or s.startswith("http://localhost") or s.startswith("file:"))


def _reason(e):
    if isinstance(e, urllib.error.HTTPError):
        if e.code == 404:
            return "HTTP 404（版本文件不存在 / 私有仓库）"
        if e.code == 403:
            try:
                if (e.headers or {}).get("X-RateLimit-Remaining") == "0":
                    return "HTTP 403（限流）"
            except Exception:
                pass
            return "HTTP 403（被拒绝）"
        return "HTTP %d" % e.code
    if isinstance(e, urllib.error.URLError):
        r = e.reason
        if isinstance(r, (socket.timeout, TimeoutError)):
            return "超时"
        if isinstance(r, ssl.SSLError):
            return "SSL 证书校验失败"
        return "无法连接（%s）" % (r,)
    if isinstance(e, (socket.timeout, TimeoutError)):
        return "超时"
    if isinstance(e, json.JSONDecodeError):
        return "返回不是 JSON（疑似被劫持）"
    return ("%s: %s" % (type(e).__name__, e))[:60]


def _parse(j, source):
    if not isinstance(j, dict):
        return None
    if "tag_name" in j:                     # GitHub Release API
        urls = []
        for a in (j.get("assets") or []):
            if isinstance(a, dict) and a.get("browser_download_url"):
                urls.append(a["browser_download_url"])
        return {"version": str(j.get("tag_name") or "").lstrip("vV"),
                "notes": (j.get("body") or "")[:NOTES_MAX],
                "page_url": j.get("html_url") or GITHUB_PAGE,
                "date": str(j.get("published_at") or "")[:10],
                "urls": urls, "sha256": "", "size": None, "source": source}
    ver = str(j.get("version") or "").lstrip("vV")
    if not ver:
        return None
    urls = j.get("urls") or j.get("url") or []
    if isinstance(urls, str):
        urls = [urls]
    return {"version": ver,
            "notes": (str(j.get("notes") or ""))[:NOTES_MAX],
            "page_url": j.get("page_url") or GITHUB_PAGE,
            "date": str(j.get("date") or "")[:10],
            "urls": [u for u in urls if isinstance(u, str)],
            "sha256": str(j.get("sha256") or "").lower(),
            "size": j.get("size"), "source": source}


def check(current=None):
    """返回 (info, attempts)。所有源失败抛 UpdateError。

    支持 PTW_UPDATE_MANIFEST 环境变量覆盖清单地址（本地路径 / http 回环 / https）：
    内网部署时把 version.json 放自己的服务器上，不必依赖 GitHub。
    """
    cur = current or APP_VERSION
    override = (os.environ.get("PTW_UPDATE_MANIFEST") or "").strip()
    if override:
        try:
            if override.lower().startswith(("http://127.0.0.1", "http://localhost", "file:")):
                raw = _get_any(override)
            else:
                raw = _get(override)
            info = _parse(json.loads(raw.decode("utf-8-sig", "replace")), "自定义源")
            if not info or not info["version"]:
                raise UpdateError("清单缺少 version 字段")
            info["is_new"] = ver_cmp(info["version"], cur) > 0
            return info, [("自定义源", True, "OK（%s）" % info["version"])]
        except UpdateError:
            raise
        except Exception as e:
            raise UpdateError("自定义更新源读取失败", [("自定义源", False, _reason(e))])
    if not UPDATE_REPO:
        raise UpdateError("未配置更新源（app_version.UPDATE_REPO 为空）")
    attempts, t0 = [], time.time()
    for name, tpl in ENDPOINTS:
        if time.time() - t0 > BUDGET:
            attempts.append((name, False, "跳过（预算用尽）"))
            continue
        try:
            j = json.loads(_get(tpl.format(repo=UPDATE_REPO)).decode("utf-8-sig", "replace"))
            info = _parse(j, name)
            if not info or not info["version"]:
                attempts.append((name, False, "缺少 version"))
                continue
            info["is_new"] = ver_cmp(info["version"], cur) > 0
            attempts.append((name, True, "OK（%s）" % info["version"]))
            return info, attempts
        except Exception as e:
            attempts.append((name, False, _reason(e)))
    raise UpdateError("所有更新源均失败", attempts)


# ---------------------------------------------------------------- 下载
def _open_url(url):
    if str(url).lower().startswith("file:"):
        return open(url[7:], "rb")
    h = {"User-Agent": "PTW/%s" % APP_VERSION}
    # api.github.com 的 Release 资产地址必须显式要二进制，否则返回一段 JSON 元数据
    if "api.github.com/repos/" in str(url) and "/releases/assets/" in str(url):
        h["Accept"] = "application/octet-stream"
    req = urllib.request.Request(url, headers=h)
    ctx = _CTX if str(url).lower().startswith("https://") else None
    return urllib.request.urlopen(req, timeout=30, context=ctx)


def download(info, progress=None):
    """按 urls 顺序下载；校验 sha256（若 version.json 给了）。返回本地 zip 路径。"""
    urls = info.get("urls") or []
    if not urls:
        raise UpdateError("version.json 没有给下载地址（urls 为空）")
    d = os.path.join(paths.DATA_DIR, "updates")
    os.makedirs(d, exist_ok=True)
    dst = os.path.join(d, "PTW_%s.zip" % info["version"])
    errs = []
    for u in urls:
        if not _allowed_url(u):
            errs.append((u, "非 https / 非内网回环，拒绝"))
            continue
        try:
            tmp = dst + ".part"
            got, total = 0, info.get("size") or 0
            with _open_url(u) as r, open(tmp, "wb") as f:
                total = total or int(r.headers.get("Content-Length") or 0)
                while True:
                    b = r.read(CHUNK)
                    if not b:
                        break
                    f.write(b)
                    got += len(b)
                    if progress:
                        progress(got, total)
            os.replace(tmp, dst)
            want = (info.get("sha256") or "").lower()
            if want:
                h = hashlib.sha256()
                with open(dst, "rb") as f:
                    for b in iter(lambda: f.read(CHUNK), b""):
                        h.update(b)
                if h.hexdigest().lower() != want:
                    os.remove(dst)
                    raise UpdateError("sha256 不匹配（下载被篡改或版本文件过期）")
            return dst
        except Exception as e:
            errs.append((u, str(e)[:120]))
            try:
                os.remove(dst + ".part")
            except OSError:
                pass
    raise UpdateError("所有下载地址均失败", errs)


def extract(zip_path, progress=None):
    """解压到 runtime_new/（先清空）。要求 zip 内顶层就是 runtime 的内容。"""
    stage = staging_dir()
    if os.path.isdir(stage):
        shutil.rmtree(stage, ignore_errors=True)
    os.makedirs(stage, exist_ok=True)
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        # 兼容「包了一层目录」的情况：探测共同的顶层前缀
        tops = {n.split("/")[0] for n in names if "/" in n}
        prefix = ""
        if len(tops) == 1 and not any(n == list(tops)[0] for n in names):
            prefix = list(tops)[0] + "/"
        z.extractall(stage)
        if prefix:
            inner = os.path.join(stage, prefix.rstrip("/"))
            for f in os.listdir(inner):
                shutil.move(os.path.join(inner, f), os.path.join(stage, f))
            shutil.rmtree(inner, ignore_errors=True)
    return stage


def prepare(info, progress=None):
    """下载 + 解压 + 写待应用标记。返回 (staging, zip)。"""
    zp = download(info, progress)
    stage = extract(zp, progress)
    doc = {"version": info["version"], "staged": stage, "zip": zp,
           "from": APP_VERSION, "at": time.strftime("%Y-%m-%d %H:%M:%S")}
    with io.open(pending_file(), "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    return stage, zp


def launch_apply():
    """触发替换。两种路径：

    · 正常情况下主程序由 launcher 拉起（环境变量 PTW_LAUNCHED=1），launcher 正在
      wait() 等它退出——这里只要 return，调用方立刻 os._exit(0)，launcher 会接手替换。
    · 用户直接双击 PTWApp.exe 时没有 launcher 在场，这里补启动一个
      `PTW.exe --apply-update`（它会等本进程退出再替换）。
    """
    if os.environ.get("PTW_LAUNCHED") == "1":
        return "launcher-waiting"
    lp = launcher_path()
    if not os.path.exists(lp):
        raise UpdateError("找不到 launcher（%s）——请用 PTW.exe 启动后再更新" % lp)
    subprocess.Popen([lp, "--apply-update"], close_fds=True,
                     creationflags=0x00000008)   # DETACHED_PROCESS
    return "spawned-launcher"


def cleanup_zip():
    try:
        for f in os.listdir(os.path.join(paths.DATA_DIR, "updates")):
            if f.endswith(".zip"):
                os.remove(os.path.join(paths.DATA_DIR, "updates", f))
    except OSError:
        pass
