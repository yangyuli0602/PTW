# -*- coding: utf-8 -*-
"""工作区结构体检：把「乱」变成机器能拦的规则。

查什么（每条都能一眼看懂、能立刻修）：
  1. 平台根只允许固定几项，多出来的（runtime_old_build、_dist_new、散落的 zip）一律报错
  2. tmp 只允许 buildenv / pyi-build 这类固定项 + 少量临时目录
  3. 04_归档 里的分发包不超过 3 个（保留最近几版，其余移去 02_备份）
  4. runtime 必须存在且完整（有 PTWApp.exe），且不允许出现第二份

用法：
    python check_workspace.py            # 体检，有问题退出码 1
    python check_workspace.py --quiet    # 只报错，不打印清单

被谁调用：build.py（构建前）、.githooks/pre-commit（提交前）、整理工作区.bat
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))

ALLOW_ROOT_FILES = {"PTW.exe", "version.json", "README.md", "启动说明.txt",
                    ".gitignore", ".gitattributes"}
ALLOW_ROOT_DIRS = {"runtime", "data", "开发", "02_备份", "04_归档", "tmp",
                   ".git", ".githooks", ".workbuddy"}
TMP_ALLOW = {"buildenv", "pyi-build", "flowtest", "_regdata", "_fixture"}
ARCH_MAX_ZIP = 3
TMP_MAX_ITEMS = 12
TMP_MAX_MB = 500

fails = []
warns = []


def bad(msg, fix=""):
    fails.append((msg, fix))


def _dir_size(p):
    n = 0
    for root, _d, files in os.walk(p):
        for f in files:
            try:
                n += os.path.getsize(os.path.join(root, f))
            except OSError:
                pass
    return n


def check_root():
    for name in sorted(os.listdir(ROOT)):
        p = os.path.join(ROOT, name)
        if os.path.isdir(p):
            if name not in ALLOW_ROOT_DIRS:
                bad("平台根出现多余目录：%s" % name,
                    "移到 02_备份/ 或删掉（它不该待在根目录）")
            continue
        if name not in ALLOW_ROOT_FILES:
            bad("平台根出现多余文件：%s" % name,
                "移到 04_归档/ 或 02_备份/")


def check_runtime():
    rt = os.path.join(ROOT, "runtime")
    if not os.path.isdir(rt):
        bad("缺少 runtime/（主程序目录）", "跑 build.py 重新打包")
        return
    if not os.path.exists(os.path.join(rt, "PTWApp.exe")):
        bad("runtime/ 不完整（没有 PTWApp.exe）", "跑 build.py --app 重新打包")
    for name in os.listdir(ROOT):
        if name.startswith("runtime_old") or name.startswith("runtime_new"):
            bad("平台根残留 %s" % name,
                "跑 build.py（会自动清）或手动移到 02_备份/")


def check_tmp():
    t = os.path.join(ROOT, "tmp")
    if not os.path.isdir(t):
        return
    items = [n for n in os.listdir(t) if n != "desktop.ini"]
    extra = [n for n in items if n not in TMP_ALLOW]
    if len(items) > TMP_MAX_ITEMS:
        bad("tmp 顶层有 %d 项（上限 %d）：%s" % (len(items), TMP_MAX_ITEMS,
                                             "、".join(sorted(items)[:8]) + " …"),
            "跑 3_整理工作区.bat，或把不再需要的临时目录移到 02_备份/")
    for n in extra:
        if n.startswith("_dist_new"):
            bad("tmp 里有构建中间产物：%s" % n, "跑 build.py（会自动清）")
    mb = _dir_size(t) / 1048576.0
    if mb > TMP_MAX_MB:
        bad("tmp 占 %.0f MB（上限 %d MB）" % (mb, TMP_MAX_MB),
            "buildenv 保留，其余临时目录移到 02_备份/")


def check_archive():
    a = os.path.join(ROOT, "04_归档")
    if not os.path.isdir(a):
        return
    zips = sorted([n for n in os.listdir(a) if n.lower().endswith(".zip")])
    if len(zips) > ARCH_MAX_ZIP:
        bad("04_归档 有 %d 个分发包（上限 %d）：%s"
            % (len(zips), ARCH_MAX_ZIP, "、".join(zips)),
            "只留最近 2-3 版，其余移到 02_备份/（GitHub 上都有，删了也能取回）")


def main():
    quiet = "--quiet" in sys.argv
    check_root()
    check_runtime()
    check_tmp()
    check_archive()

    if not quiet:
        for n in ("runtime", "data", "开发", "04_归档", "tmp"):
            p = os.path.join(ROOT, n)
            if os.path.isdir(p):
                print("  %-8s %6.0f MB" % (n, _dir_size(p) / 1048576.0))
        print("  根目录固定项：%s" % "、".join(sorted(n for n in os.listdir(ROOT)
                                                if not os.path.isdir(os.path.join(ROOT, n)))))
    if fails:
        print("\n[结构体检] 发现 %d 个问题：" % len(fails))
        for msg, fix in fails:
            print("  ✗ %s" % msg)
            if fix:
                print("    → %s" % fix)
        return 1
    print("\n[结构体检] 通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
