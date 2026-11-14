# -*- coding: utf-8 -*-
"""一键构建：打包 launcher + 主程序 + 出发布包 + 生成 version.json 片段。

用法（在 开发/构建/ 下）：
    ..\..\tmp\buildenv\Scripts\python.exe build.py            # 全流程
    ..\..\tmp\buildenv\Scripts\python.exe build.py --app      # 只重打主程序（常用）
    ..\..\tmp\buildenv\Scripts\python.exe build.py --launcher # 只重打启动器

产物：
    平台根/PTW.exe                 启动器（永不更新）
    平台根/runtime/                主程序目录（更新时整体替换）
    04_归档/PTW_v{ver}_win64.zip   对外分发包（不含 data）
    04_归档/version.json           仓库根要放的内容（含 url/sha256 占位）
"""
import hashlib
import io
import json
import os
import shutil
import subprocess
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))     # 平台根
APP = os.path.join(HERE, "..", "程序")
PY = os.path.join(ROOT, "tmp", "buildenv", "Scripts", "python.exe")
if not os.path.exists(PY):
    PY = sys.executable
WORK = os.path.join(ROOT, "tmp", "pyi-build")
ARCH = os.path.join(ROOT, "04_归档")


def sh(*args):
    print("»", " ".join(os.path.basename(a) if os.sep in a else a for a in args))
    r = subprocess.run(list(args), cwd=HERE)
    if r.returncode != 0:
        raise SystemExit("构建失败：%s" % args[0])


def version():
    sys.path.insert(0, APP)
    try:
        from app_version import APP_VERSION
        return APP_VERSION
    finally:
        sys.path.pop(0)


def clean_pycache(base):
    n = 0
    for root, dirs, _ in os.walk(base):
        for d in list(dirs):
            if d == "__pycache__":
                shutil.rmtree(os.path.join(root, d), ignore_errors=True)
                dirs.remove(d)
                n += 1
    return n


def build_launcher():
    sh(PY, "-m", "PyInstaller", "--noconfirm", "--clean",
       "--distpath", ROOT, "--workpath", os.path.join(WORK, "launcher"),
       os.path.join(HERE, "launcher.spec"))
    exe = os.path.join(ROOT, "PTW.exe")
    print("   launcher →", exe if os.path.exists(exe) else "(缺失!)")


def build_app():
    """打包主程序。

    注意：不能直接 --distpath 到 平台根 让 PyInstaller 覆盖 runtime/——
    它 --clean 时会批量删除旧目录（上百个文件），会被批量删除保护拦下，
    而且中途失败会留下半个 runtime。这里改成：
        ① 先打到临时目录 tmp/_dist_new/runtime
        ② 成功后把旧 runtime 改名为 runtime_old_<时间戳>（改名不删文件）
        ③ 新目录改名成 runtime
    """
    dist = os.path.join(ROOT, "tmp", "_dist_new")
    old = os.path.join(ROOT, "runtime_old_build")
    if os.path.isdir(old):
        old = old + "_" + time.strftime("%H%M%S")
    if os.path.isdir(dist):
        # 上一次的临时产物：改名而不是删
        try:
            os.rename(dist, dist + "_" + time.strftime("%H%M%S"))
        except OSError:
            pass
    sh(PY, "-m", "PyInstaller", "--noconfirm", "--clean",
       "--distpath", dist, "--workpath", os.path.join(WORK, "app"),
       os.path.join(HERE, "PTWApp.spec"))
    new_rt = os.path.join(dist, "runtime")
    if not os.path.isdir(new_rt):
        raise SystemExit("打包失败：没有产出 %s" % new_rt)
    rt = os.path.join(ROOT, "runtime")
    if os.path.isdir(rt):
        os.rename(rt, old)
        print("   旧 runtime → %s" % os.path.basename(old))
    os.rename(new_rt, rt)
    print("   主程序 →", rt)


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def make_zip(ver):
    os.makedirs(ARCH, exist_ok=True)
    out = os.path.join(ARCH, "PTW_v%s_win64.zip" % ver)
    if os.path.exists(out):
        os.remove(out)
    n = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        z.write(os.path.join(ROOT, "PTW.exe"), "PTW.exe")
        rt = os.path.join(ROOT, "runtime")
        for root, dirs, files in os.walk(rt):
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            for f in files:
                if f.endswith(".pyc"):
                    continue
                p = os.path.join(root, f)
                z.write(p, os.path.join("runtime", os.path.relpath(p, rt)))
                n += 1
        for extra in ("启动说明.txt",):
            p = os.path.join(ROOT, extra)
            if os.path.exists(p):
                z.write(p, extra)
    size = os.path.getsize(out)
    print("   发布包 → %s（%d 文件 / %.1f MB）" % (out, n, size / 1048576))
    return out, size


def write_version_json(ver, zpath=None, size=None):
    """生成仓库根要放的 version.json 骨架（url/sha256 填好，仓库地址留空由人填）。"""
    doc = {
        "version": ver,
        "date": time.strftime("%Y-%m-%d"),
        "notes": "（在此写本次更新说明，会显示在用户的更新提示里）",
        "urls": ["https://github.com/<owner>/<repo>/releases/download/v%s/PTW_v%s_win64.zip"
                 % (ver, ver)],
        "sha256": sha256(zpath) if zpath else "",
        "size": size or 0,
    }
    p = os.path.join(ARCH, "version.json")
    with io.open(p, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    print("   version.json → %s" % p)
    return doc


def main():
    args = sys.argv[1:]
    do_app = not args or "--app" in args
    do_launcher = not args or "--launcher" in args
    ver = version()
    print("=" * 60)
    print("构建 PTW v%s  根目录=%s" % (ver, ROOT))
    print("=" * 60)
    clean_pycache(os.path.join(APP, "modules"))
    clean_pycache(os.path.join(APP, "core"))
    if do_launcher:
        build_launcher()
    if do_app:
        build_app()
    zp, size = make_zip(ver)
    write_version_json(ver, zp, size)
    # 本机版本指针
    with io.open(os.path.join(ROOT, "version.json"), "w", encoding="utf-8") as f:
        json.dump({"current": ver, "built_at": time.strftime("%Y-%m-%d %H:%M:%S")},
                  f, ensure_ascii=False, indent=2)
    print("\n完成。发版顺序：① 打包 → ② 建 Release 传 zip → ③ 最后改仓库 version.json")


if __name__ == "__main__":
    main()
