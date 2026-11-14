# -*- coding: utf-8 -*-
"""PTW 一键发版（GitHub）。与 MDMQ 同一套做法，token 放本地文件，不进仓库。

用法（在 开发/构建 目录执行，用 tmp/buildenv 的 python）：
    publish.py --check                只验证 token 与仓库，不改任何东西
    publish.py --setup                首次：建仓库（已存在则跳过）+ 写 UPDATE_REPO + 推首版代码
    publish.py                        发版：打包 → 建 Release → 传 zip → 更新 version.json
    publish.py --notes "修复 xxx"      自定义更新说明（会显示在用户的更新提示里）
    publish.py --skip-build           复用上次打好的包（只在确认代码没变时用）
    publish.py --dry-run              只生成本地清单，不联网、不提交

token 文件（只放一行，绝不提交）：
    Windows: C:\\Users\\<你>\\.workbuddy\\ptw_token.txt
    权限至少 repo（经典 token 勾 repo；fine-grained 需 Contents+Administration 写权限）

顺序不能颠倒（原因：客户端先读 version.json 再去下载，早了就是 404）：
    1) 打包 zip  2) 建 Release 传 zip  3) 最后写 version.json
"""
import base64
import hashlib
import importlib.util
import json
import os
import re
import ssl
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
PROJ = HERE.parent.parent                      # 01_PTW平台
APP_SRC = PROJ / "开发" / "程序"
ARCH = PROJ / "04_归档"
TOKEN_CANDIDATES = [
    Path.home() / ".workbuddy" / "ptw_token.txt",
    Path.home() / ".workbuddy" / "github_token.txt",
    # 同一台机器上已经配过的 MDMQ token 直接复用（同一个 GitHub 账号）
    Path.home() / ".workbuddy" / "mdmq_token.txt",
]
API = "https://api.github.com"
APP_NAME = "PTW"
MANIFEST_REMOTE = "version.json"               # 客户端从 main 分支根目录读它
MANIFEST_LOCAL = ARCH / "version.json"


def die(msg):
    print("\n[失败] " + msg)
    sys.exit(1)


def ok(msg):
    print("  [OK] " + msg)


def _ctx():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


_CTX = _ctx()
_OPENER = urllib.request.build_opener(urllib.request.HTTPSHandler(context=_CTX))


# ------------------------------------------------------------------ 基础
def token():
    """按顺序找 token 文件，取第一个有效的；绝不打印内容。"""
    found = [p for p in TOKEN_CANDIDATES if p.exists()]
    if not found:
        die("找不到 token 文件。任选一个位置放：\n  " +
            "\n  ".join(str(p) for p in TOKEN_CANDIDATES) +
            "\n（只放一行，别加引号）\n生 token：GitHub → Settings → Developer settings → "
            "Personal access tokens → Tokens (classic) → 勾 repo")
    for p in found:
        t = p.read_text(encoding="utf-8-sig", errors="replace").strip()
        t = t.replace("\ufeff", "").strip().strip('"').strip("'")
        if not t:
            continue
        if not (t.startswith("ghp_") or t.startswith("github_pat_")):
            continue
        print("  token 来源：%s" % p)
        return t
    die("找到了 token 文件但内容不可用（空 / 不以 ghp_ 或 github_pat_ 开头）：\n  " +
        "\n  ".join(str(p) for p in found))


def api(path, method="GET", body=None, tok=None, allow_404=False, raw=False):
    url = path if path.startswith("http") else API + path
    data = None
    headers = {
        "Authorization": "Bearer " + (tok or token()),
        "Accept": "application/vnd.github+json",
        "User-Agent": "PTW-release",
    }
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, method=method, data=data, headers=headers)
    last = None
    for attempt in range(4):
        try:
            with _OPENER.open(req, timeout=90) as r:
                buf = r.read()
            return buf if raw else json.loads(buf.decode("utf-8"))
        except urllib.error.HTTPError as e:
            # 409 = 空仓库（Git Repository is empty），语义上等同"还不存在"
            if allow_404 and e.code in (404, 409):
                return None
            if e.code >= 500 and attempt < 3:
                last = "HTTP %s" % e.code
                time.sleep(1.5 * (attempt + 1))
                continue
            detail = e.read().decode("utf-8", "replace")[:400]
            die("%s %s → HTTP %s\n%s" % (method, path, e.code, detail))
        except Exception as e:                      # 连接被掐、SSL EOF、超时
            last = "%s: %s" % (type(e).__name__, str(e)[:80])
            if attempt < 3:
                print("    网络抖动（%s），重试 %d/3" % (last, attempt + 1), flush=True)
                time.sleep(1.5 * (attempt + 1))
                continue
            die("连不上 GitHub：%s\n代理不稳定时重跑一次即可（已建的 blob 会复用）。" % last)


def app_version_module():
    spec = importlib.util.spec_from_file_location("av", APP_SRC / "app_version.py")
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def repo_name():
    r = str(getattr(app_version_module(), "UPDATE_REPO", "") or "").strip()
    if not r:
        die("app_version.py 里 UPDATE_REPO 还是空的。\n"
            "先跑一次：publish.py --setup（会自动建仓库并填好）")
    if r.count("/") != 1:
        die("UPDATE_REPO 格式应为 owner/repo，当前是：%s" % r)
    return r


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def _tracked_files():
    out = subprocess.run(["git", "-C", str(PROJ), "ls-files", "-z"],
                         capture_output=True)
    return [p for p in (out.stdout or b"").decode("utf-8").split("\0") if p]


def api_push(repo, message, tok, note=""):
    """用 GitHub 的 Git Data API 推代码——不走 git 协议。

    为什么：本机到 github.com 的 443 被网络拦（git push 必失败），
    而 api.github.com 通。blobs → tree(base_tree 保底) → commit → ref，效果等同 push。
    用 base_tree 是为了不删掉通过 API 单独维护的文件（如根目录 version.json）。
    """
    # 空仓库不能直接建 blob（HTTP 409 Git Repository is empty）：
    # 先用 contents API 落一个小文件把 main 分支生出来，再走 blob/tree/commit。
    ref = api("/repos/%s/git/ref/heads/main" % repo, tok=tok, allow_404=True)
    if not ref:
        seed = (PROJ / "README.md")
        body = seed.read_bytes() if seed.exists() else b"# PTW\n"
        api("/repos/%s/contents/%s" % (repo, "README.md"), "PUT", {
            "message": "chore: 初始化仓库",
            "content": base64.b64encode(body).decode(),
        }, tok=tok)
        ok("空仓库已初始化 main 分支")

    files = _tracked_files()
    print("  待推送文件 %d 个" % len(files))
    entries = []
    for i, rel in enumerate(files, 1):
        p = PROJ / rel
        if not p.is_file():
            continue
        data = p.read_bytes()
        if len(data) > 20 * 1024 * 1024:
            print("    跳过过大文件 %s" % rel)
            continue
        blob = api("/repos/%s/git/blobs" % repo, "POST",
                   {"content": base64.b64encode(data).decode(), "encoding": "base64"},
                   tok=tok)
        entries.append({"path": rel.replace("\\", "/"), "mode": "100644",
                        "type": "blob", "sha": blob["sha"]})
        if i % 40 == 0:
            print("    %d/%d" % (i, len(files)), flush=True)

    parent = base_tree = None
    ref = api("/repos/%s/git/ref/heads/main" % repo, tok=tok, allow_404=True)
    if isinstance(ref, dict) and ref.get("object"):
        parent = ref["object"]["sha"]
        pc = api("/repos/%s/git/commits/%s" % (repo, parent), tok=tok)
        base_tree = ((pc.get("tree") or {}).get("sha"))

    body = {"tree": entries}
    if base_tree:
        body["base_tree"] = base_tree
    tree = api("/repos/%s/git/trees" % repo, "POST", body, tok=tok)
    cb = {"message": message, "tree": tree["sha"]}
    if parent:
        cb["parents"] = [parent]
    commit = api("/repos/%s/git/commits" % repo, "POST", cb, tok=tok)
    if parent:
        api("/repos/%s/git/refs/heads/main" % repo, "PATCH",
            {"sha": commit["sha"], "force": False}, tok=tok)
    else:
        api("/repos/%s/git/refs" % repo, "POST",
            {"ref": "refs/heads/main", "sha": commit["sha"]}, tok=tok)
    ok("已推送 %s（commit %s）%s" % ("main", commit["sha"][:8], note))
    return commit["sha"]


def local_commit(message):
    """本地先提交（失败不致命：可能没有改动）。"""
    git = ["git", "-C", str(PROJ)]
    subprocess.run(git + ["add", "-A"], capture_output=True)
    r = subprocess.run(git + ["commit", "-q", "-m", message], capture_output=True)
    return r.returncode == 0


def write_repo_config(full):
    """把 UPDATE_REPO 写死进 app_version.py（打包才带着它）。"""
    p = APP_SRC / "app_version.py"
    s = p.read_text(encoding="utf-8")
    if not re.search(r'(?m)^UPDATE_REPO\s*=', s):
        die("没找到 UPDATE_REPO 那一行，请手动改 %s" % p)
    new = re.sub(r'(?m)^UPDATE_REPO\s*=.*$',
                 'UPDATE_REPO = "%s"' % full, s, count=1)
    if new == s:
        ok("app_version.py 已是 %s（无需改动）" % full)
        return
    p.write_text(new, encoding="utf-8")
    ok("app_version.py → UPDATE_REPO = %s" % full)


# ------------------------------------------------------------------ 步骤
def check():
    print("1) 验证 token 与仓库")
    me = api("/user")
    ok("账号 %s" % me.get("login"))
    r = str(getattr(app_version_module(), "UPDATE_REPO", "") or "").strip()
    if not r:
        print("  [注意] UPDATE_REPO 尚未配置，跑 --setup 建仓库")
        return
    repo = api("/repos/" + r)
    perm = (repo.get("permissions") or {})
    if not perm.get("push"):
        die("该 token 对 %s 没有写权限" % r)
    ok("仓库 %s 可写（private=%s）" % (r, repo.get("private")))
    rel = api("/repos/%s/releases/latest" % r, allow_404=True)
    if rel:
        ok("最新 Release：%s（%s）" % (rel.get("tag_name"), (rel.get("published_at") or "")[:10]))
    else:
        print("  [注意] 还没有任何 Release")
    m = api("/repos/%s/contents/%s" % (r, MANIFEST_REMOTE), allow_404=True)
    print("  version.json 在仓库里：%s" % ("有" if m else "还没有（首次发版会写）"))
    print("\n当前代码版本 v%s" % app_version_module().APP_VERSION)


def setup(argv):
    """首次：建仓库 + 写 UPDATE_REPO + 推代码。已存在就只补配置。"""
    tok = token()
    me = api("/user", tok=tok)
    login = me.get("login")
    ok("账号 %s" % login)

    name = "PTW"
    private = False
    for i, a in enumerate(argv):
        if a == "--repo" and i + 1 < len(argv):
            name = argv[i + 1]
        if a == "--private":
            private = True
    full = "%s/%s" % (login, name)

    print("1) 仓库 %s" % full)
    repo = api("/repos/" + full, tok=tok, allow_404=True)
    if repo:
        ok("已存在，跳过创建（private=%s）" % repo.get("private"))
    else:
        repo = api("/user/repos", "POST", {
            "name": name,
            "description": "PTW 平台：PBC to WP 审计底稿工具（客户资料 → 底稿）",
            "private": private, "has_issues": True, "has_wiki": False,
            "auto_init": False,
        }, tok=tok)
        ok("已创建（private=%s）" % repo.get("private"))

    print("2) 写入 UPDATE_REPO")
    write_repo_config(full)

    print("3) 推送代码到 main")
    git = ["git", "-C", str(PROJ)]
    remotes = subprocess.run(git + ["remote"], capture_output=True, text=True).stdout
    url = "https://github.com/%s.git" % full
    if "origin" in remotes.split():
        subprocess.run(git + ["remote", "set-url", "origin", url], capture_output=True)
    else:
        subprocess.run(git + ["remote", "add", "origin", url], capture_output=True)
    ok("remote origin → %s" % url)

    local_commit("chore: 配置更新源 %s" % full)
    api_push(full, "chore: 初始化 %s" % APP_NAME, tok)
    print("\n下一步：publish.py --notes \"首个版本\" —— 打包并发第一版 Release")


def build_package():
    """调 build.py 全量打包，返回 (zip路径, 版本)。"""
    print("1) 打包（launcher + 主程序 + zip）")
    py = sys.executable
    r = subprocess.run([py, str(HERE / "build.py")], cwd=str(HERE),
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        die("打包失败：\n%s" % ((r.stdout or "")[-800:] + (r.stderr or "")[-800:]))
    ver = app_version_module().APP_VERSION
    zp = ARCH / ("PTW_v%s_win64.zip" % ver)
    if not zp.exists():
        die("打包完成但找不到 %s" % zp)
    ok("zip %s（%.1f MB）" % (zp.name, zp.stat().st_size / 1048576))
    return zp, ver


def make_manifest(ver, zp, notes, repo, asset=None):
    size = zp.stat().st_size
    tag = "v" + ver
    urls = []
    # 资产 API 地址放第一位：它走 api.github.com，国内比 releases/download 稳，
    # 公开仓库无需 token（客户端下载时会带 Accept: application/octet-stream）
    if asset and asset.get("url"):
        urls.append(asset["url"])
    urls.append("https://github.com/%s/releases/download/%s/%s" % (repo, tag, zp.name))
    doc = {
        "version": ver,
        "date": time.strftime("%Y-%m-%d"),
        "notes": notes,
        "urls": urls,
        "sha256": sha256(zp),
        "size": size,
        "page_url": "https://github.com/%s/releases/latest" % repo,
    }
    MANIFEST_LOCAL.write_text(json.dumps(doc, ensure_ascii=False, indent=2),
                              encoding="utf-8")
    ok("清单 → %s" % MANIFEST_LOCAL)
    return doc


def upload_asset(repo, release_id, zp, tok):
    import http.client
    total = zp.stat().st_size
    conn = http.client.HTTPSConnection("uploads.github.com", timeout=300, context=_CTX)
    path = "/repos/%s/releases/%d/assets?name=%s" % (
        repo, release_id, urllib.parse.quote(zp.name))
    conn.putrequest("POST", path)
    conn.putheader("Authorization", "Bearer " + tok)
    conn.putheader("Accept", "application/vnd.github+json")
    conn.putheader("Content-Type", "application/zip")
    conn.putheader("Content-Length", str(total))
    conn.putheader("User-Agent", "PTW-release")
    conn.endheaders()
    sent = last = 0
    t0 = time.time()
    with open(zp, "rb") as f:
        while True:
            buf = f.read(1 << 20)
            if not buf:
                break
            conn.send(buf)
            sent += len(buf)
            if sent - last >= 10 * 1048576:
                last = sent
                sp = sent / max(time.time() - t0, 1e-6) / 1024
                print("    %.1f/%.1f MB  %.0f KB/s" % (sent / 1048576, total / 1048576, sp),
                      flush=True)
    r = conn.getresponse()
    body = r.read().decode("utf-8", "replace")
    if r.status not in (200, 201):
        die("上传失败 HTTP %s\n%s" % (r.status, body[:400]))
    d = json.loads(body)
    ok("资产 %s  id=%s  %s 字节" % (d["name"], d["id"], d.get("size")))
    return d


def put_file(repo, remote, doc, message, tok):
    b64 = base64.b64encode(json.dumps(doc, ensure_ascii=False, indent=2)
                           .encode("utf-8")).decode()
    cur = api("/repos/%s/contents/%s" % (repo, remote), tok=tok, allow_404=True)
    body = {"message": message, "content": b64}
    if isinstance(cur, dict) and cur.get("sha"):
        body["sha"] = cur["sha"]
    api("/repos/%s/contents/%s" % (repo, remote), "PUT", body, tok=tok)
    ok("%s 已更新（main 分支）" % remote)


def purge_silent(repo):
    for p in ("gh/%s@main/version.json", "gh/%s/version.json"):
        try:
            _OPENER.open("https://purge.jsdelivr.net/" + (p % repo), timeout=20).read()
        except Exception:
            pass


def release(argv):
    notes = ""
    skip_build = "--skip-build" in argv
    dry = "--dry-run" in argv
    for i, a in enumerate(argv):
        if a == "--notes" and i + 1 < len(argv):
            notes = argv[i + 1]

    repo = str(getattr(app_version_module(), "UPDATE_REPO", "") or "").strip() or "owner/repo"

    if dry:
        print("== dry-run：只生成本地清单，不联网 ==")
        ver = app_version_module().APP_VERSION
        zp = ARCH / ("PTW_v%s_win64.zip" % ver)
        if not zp.exists():
            die("还没有包：%s\n先跑 publish.py（不加 --dry-run）或 build.py" % zp)
        doc = make_manifest(ver, zp, notes or "（dry-run）", repo)
        print(json.dumps({k: doc[k] for k in ("version", "size", "sha256")},
                         ensure_ascii=False))
        print("\n将向 GitHub 发：\n  Release  v%s  ← %s\n  PUT      %s"
              % (doc["version"], zp.name, MANIFEST_REMOTE))
        return

    tok = token()
    ver = app_version_module().APP_VERSION
    tag = "v" + ver
    today = time.strftime("%Y-%m-%d")
    print("发版 %s %s（%s）  仓库 %s\n" % (APP_NAME, tag, today, repo))

    if skip_build:
        zp = ARCH / ("PTW_v%s_win64.zip" % ver)
        if not zp.exists():
            die("--skip-build 但找不到 %s" % zp)
        print("1) 复用已有包 %s" % zp.name)
    else:
        zp, ver = build_package()
        tag = "v" + ver

    print("1.5) 推送代码到 main")
    local_commit("release: %s" % tag)
    api_push(repo, "release: %s" % tag, tok)

    print("2) 建 Release（已存在则复用）")
    rel = api("/repos/%s/releases/tags/%s" % (repo, tag), tok=tok, allow_404=True)
    if rel:
        ok("复用已有 Release id=%s" % rel["id"])
    else:
        rel = api("/repos/%s/releases" % repo, "POST", {
            "tag_name": tag, "target_commitish": "main",
            "name": "%s %s" % (APP_NAME, tag),
            "body": notes or ("%s %s" % (APP_NAME, tag)),
            "draft": False, "prerelease": False,
        }, tok=tok)
        ok("新建 Release id=%s" % rel["id"])

    print("3) 上传安装包")
    for a in (rel.get("assets") or []):
        if a.get("name") == zp.name:
            api("/repos/%s/releases/assets/%s" % (repo, a["id"]), "DELETE", tok=tok)
            print("    删掉同名旧资产 %s" % a["id"])
    asset = upload_asset(repo, rel["id"], zp, tok)

    print("4) 最后写 version.json（客户端据此判断版本）")
    doc = make_manifest(ver, zp, notes or ("%s %s" % (APP_NAME, tag)), repo, asset)
    put_file(repo, MANIFEST_REMOTE, doc, "release: %s" % tag, tok)

    print("5) 清缓存")
    try:
        purge_silent(repo)
    except Exception:
        pass

    print("\n完成：https://github.com/%s/releases/tag/%s" % (repo, tag))
    print("包：%s（%s 字节）" % (zp, doc["size"]))
    print("用户端：打开 PTW → 检查更新 → 立即更新")


def purge_silent(repo):
    for p in ("gh/%s@main/version.json", "gh/%s/version.json"):
        try:
            _OPENER.open("https://purge.jsdelivr.net/" + (p % repo), timeout=20).read()
        except Exception:
            pass


def main():
    argv = sys.argv[1:]
    if "--help" in argv or "-h" in argv:
        print(__doc__)
        return
    if "--check" in argv:
        check()
        return
    if "--setup" in argv:
        setup(argv)
        return
    release(argv)


if __name__ == "__main__":
    main()
