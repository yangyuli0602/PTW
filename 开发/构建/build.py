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


def _dir_size(p):
    n = 0
    for root, _dirs, files in os.walk(p):
        for f in files:
            try:
                n += os.path.getsize(os.path.join(root, f))
            except OSError:
                pass
    return n


def prune_leftovers():
    """把上一次构建的残留挪进 tmp/_build_prune/，不在构建过程中删任何东西。

    为什么是「挪」不是「删」：本机有批量删除保护，构建里删 80 个文件的目录会被拦，
    构建直接失败。挪进 tmp 后由 tidy_workspace.py（需要人确认）统一清，
    这样「构建永不删除」+「临时区有出口」两头都满足。
    """
    import glob
    dest = os.path.join(ROOT, "tmp", "_build_prune")
    moved = []
    for p in glob.glob(os.path.join(ROOT, "runtime_old_build*")) + \
             glob.glob(os.path.join(ROOT, "tmp", "_dist_new*")):
        tgt = os.path.join(dest, os.path.basename(p))
        n = 1
        while os.path.exists(tgt):
            tgt = os.path.join(dest, "%s_%d" % (os.path.basename(p), n))
            n += 1
        os.makedirs(dest, exist_ok=True)
        try:
            os.rename(p, tgt)
            moved.append(os.path.relpath(p, ROOT))
        except OSError as e:
            print("   挪不动 %s（%s）" % (os.path.basename(p), e))
    if moved:
        print("» 上次构建残留已挪进 tmp/_build_prune（%d 项）" % len(moved))


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
    prune_leftovers()
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


def published_record():
    """publish.py 发版成功后写下的记录：{版本: {sha256,size,tag,at}}。

    用来判断「这个版本号是不是已经发出去过」——已发布的版本号不能再用不同字节覆盖，
    否则客户端拿着 1.3.4 却有两份不同的包，追溯就断了。
    """
    p = os.path.join(ARCH, "已发布.json")
    if not os.path.exists(p):
        return {}
    try:
        with io.open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def guard_republish(ver, args):
    """包已存在且该版本已发布过 → 拒绝，除非显式 --allow-republish。"""
    rec = published_record().get(ver)
    zp = os.path.join(ARCH, "PTW_v%s_win64.zip" % ver)
    if not rec:
        return
    print("  注意：v%s 已经发布过（%s，sha %s…）"
          % (ver, rec.get("at", "?"), (rec.get("sha256") or "")[:12]))
    if os.path.exists(zp):
        old = sha256(zp)
        if old.lower() == (rec.get("sha256") or "").lower():
            print("  本地包与已发布版本字节一致，可以继续（等同重新打包，不影响追溯）")
            return
    if "--allow-republish" in args:
        print("  !! --allow-republish：允许用不同字节覆盖已发布版本号（追溯会断，慎用）")
        return
    raise SystemExit(
        "\n构建已中止：v%s 已经发布过，但当前字节与已发布的不同。\n"
        "  已发布 sha256：%s\n"
        "  请先把 app_version.py 的 APP_VERSION 升一位（推荐），\n"
        "  或确认要覆盖时加 --allow-republish。\n"
        % (ver, (rec.get("sha256") or "")[:32]))


def run_smoke():
    """打包后冒烟：真起一次，验版本/板块/首页/模板。"""
    r = subprocess.run([sys.executable, os.path.join(HERE, "smoke.py")])
    if r.returncode != 0:
        raise SystemExit("\n构建已中止：打包产物冒烟未通过（上面有具体原因）。")


def main():
    args = sys.argv[1:]
    do_app = not args or "--app" in args
    do_launcher = not args or "--launcher" in args
    ver = version()
    print("=" * 60)
    print("构建 PTW v%s  根目录=%s" % (ver, ROOT))
    print("=" * 60)
    prune_leftovers()
    guard_republish(ver, args)
    # 构建前先做结构体检：根目录有野文件、tmp 爆了，就在这里拦住（--force 可越过）
    if "--force" not in args:
        rc = subprocess.run([sys.executable, os.path.join(HERE, "check_workspace.py"),
                             "--quiet"]).returncode
        if rc != 0:
            raise SystemExit("\n构建已中止：工作区结构不达标。\n"
                             "先跑 tidy_workspace.py 整理，或确认无所谓时加 --force。")
    clean_pycache(os.path.join(APP, "modules"))
    clean_pycache(os.path.join(APP, "core"))
    if do_launcher:
        build_launcher()
    if do_app:
        build_app()
    zp, size = make_zip(ver)
    if do_app and "--no-smoke" not in args:
        run_smoke()
    write_version_json(ver, zp, size)
    # 本机版本指针
    with io.open(os.path.join(ROOT, "version.json"), "w", encoding="utf-8") as f:
        json.dump({"current": ver, "built_at": time.strftime("%Y-%m-%d %H:%M:%S")},
                  f, ensure_ascii=False, indent=2)
    print("\n完成。发版顺序：① 打包 → ② 建 Release 传 zip → ③ 最后改仓库 version.json")
    # 构建成功后不再需要上一版 runtime 备份（新包已进 04_归档，随时能回滚），
    # 留一份在根目录就是每构建一次多 64MB，根目录也不该有临时目录。
    prune_leftovers()


if __name__ == "__main__":
    main()
