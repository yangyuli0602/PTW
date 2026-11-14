# -*- coding: utf-8 -*-
"""打包后冒烟：把刚打出来的 runtime 真起一次，确认「能跑、版本对、板块全、模板在」。

为什么要有：打包成功 ≠ 能运行。PyInstaller 少收一个 datas、少一个 hiddenimport，
zip 照样生成，客户一点就白屏。这一步 20 秒，能拦住整类事故。

用法：
    python smoke.py                  # 用 runtime/PTWApp.exe 起一次，验完自动关
    python smoke.py --exe <path>     # 指定可执行文件
    python smoke.py --keep           # 不关进程（排查用）

判定（任一不过即退出码 1）：
  1. 10 秒内 /api/modules 有响应
  2. 返回的 version == app_version.APP_VERSION
  3. 板块数 >= 3（ingest / analysis / bkd）
  4. 首页 HTML 能取到且含 PTW
  5. 首次启动把模板释放到 <数据目录>/bkd_tpl（至少 1 个 xlsx）
"""
import json
import os
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
APP = os.path.join(HERE, "..", "程序")


def free_port(start=8895):
    for p in range(start, start + 30):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    raise RuntimeError("找不到空闲端口")


def ver():
    sys.path.insert(0, APP)
    try:
        from app_version import APP_VERSION
        return APP_VERSION
    finally:
        sys.path.pop(0)


def main():
    argv = sys.argv[1:]
    exe = os.path.join(ROOT, "runtime", "PTWApp.exe")
    for i, a in enumerate(argv):
        if a == "--exe" and i + 1 < len(argv):
            exe = argv[i + 1]
    if not os.path.exists(exe):
        print("[冒烟] 找不到 %s" % exe)
        return 1

    port = free_port()
    data = tempfile.mkdtemp(prefix="smoke_", dir=os.path.join(ROOT, "tmp"))
    env = dict(os.environ, KZ_PORT=str(port), KZ_NO_BROWSER="1", KZ_DATA_DIR=data)
    print("[冒烟] 起 %s（端口 %d，数据目录 tmp/%s）"
          % (os.path.relpath(exe, ROOT), port, os.path.basename(data)))
    p = subprocess.Popen([exe], cwd=os.path.dirname(exe), env=env,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    base = "http://127.0.0.1:%d" % port
    fails = []
    try:
        info = None
        for _ in range(40):                       # 最多等 20s
            try:
                with urllib.request.urlopen(base + "/api/modules", timeout=2) as r:
                    info = json.loads(r.read().decode("utf-8"))
                break
            except Exception:
                time.sleep(0.5)
        if not info:
            fails.append("10 秒内 /api/modules 没有响应（进程可能启动即崩）")
        else:
            want = ver()
            got = info.get("version")
            print("    版本 打包 %s / 源码 %s" % (got, want))
            if got != want:
                fails.append("版本不一致：包内 %s，源码 %s" % (got, want))
            mods = info.get("modules") or []
            print("    板块 %d 个：%s" % (len(mods), "、".join(m.get("key", "?") for m in mods)))
            if len(mods) < 3:
                fails.append("板块只有 %d 个（应 ≥3）" % len(mods))
            try:
                with urllib.request.urlopen(base + "/", timeout=5) as r:
                    html = r.read().decode("utf-8", "replace")
                if "PTW" not in html:
                    fails.append("首页没有 PTW 字样（web 资源没打进包？）")
                else:
                    print("    首页 OK（%d 字节）" % len(html))
            except Exception as e:
                fails.append("首页取不到：%s" % e)

        tpl = os.path.join(data, "bkd_tpl")
        files = [f for f in os.listdir(tpl)] if os.path.isdir(tpl) else []
        print("    模板释放 %d 个" % len(files))
        if not any(f.lower().endswith(".xlsx") for f in files):
            fails.append("模板没有释放到 %s（首启 seed 失效）" % os.path.relpath(tpl, ROOT))
    finally:
        if "--keep" not in argv:
            try:
                p.terminate()
                p.wait(timeout=8)
            except Exception:
                try:
                    p.kill()
                except Exception:
                    pass

    if fails:
        print("\n[冒烟] 失败 %d 项：" % len(fails))
        for f in fails:
            print("   ✗", f)
        return 1
    print("\n[冒烟] 通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
