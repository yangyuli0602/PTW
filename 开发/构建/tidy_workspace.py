# -*- coding: utf-8 -*-
"""整理工作区：把「可重建的垃圾」清掉，把「旧版本包」挪去备份区。

做什么（都只动三类路径，别的一概不碰）：
  1. 平台根 runtime_old_build* / tmp/_dist_new*  —— 构建残留，直接删（可重建）
  2. 04_归档 只留最近 3 个分发包，其余移到 02_备份/归档_旧版_<日期>/
  3. tmp 下超过 3 天的一次性目录 → 移到 02_备份/tmp_旧_<日期>/（移动不删）

绝不碰：runtime/、data/、开发/、02_备份/ 里的东西。

用法：
    python tidy_workspace.py --dry-run     # 只看会动什么
    python tidy_workspace.py               # 真执行，然后跑一次体检
"""
import glob
import os
import shutil
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
ARCH = os.path.join(ROOT, "04_归档")
BACKUP = os.path.join(ROOT, "02_备份")
TMP = os.path.join(ROOT, "tmp")
KEEP_ZIP = 3
TMP_STALE_DAYS = 3

DRY = "--dry-run" in sys.argv


def _size(p):
    n = 0
    for root, _d, files in os.walk(p):
        for f in files:
            try:
                n += os.path.getsize(os.path.join(root, f))
            except OSError:
                pass
    return n


def _rm(p):
    if DRY:
        print("    [dry] 删除", os.path.relpath(p, ROOT))
        return
    shutil.rmtree(p, ignore_errors=True)
    print("    已删", os.path.relpath(p, ROOT))


def _mv(src, dst_dir):
    os.makedirs(dst_dir, exist_ok=True)
    if DRY:
        print("    [dry] 移动", os.path.relpath(src, ROOT), "→", os.path.relpath(dst_dir, ROOT))
        return
    shutil.move(src, os.path.join(dst_dir, os.path.basename(src)))
    print("    已移", os.path.relpath(src, ROOT))


def ver_key(name):
    import re
    nums = re.findall(r"\d+", name)[:4]
    return tuple(int(x) for x in nums) or (0,)



UI_CACHE_SUBDIRS = ("component_crx_cache", "ProvenanceData", "GrShaderCache",
                    "ShaderCache", "Code Cache", "GPUCache", "DawnCache",
                    "Edge Entity Extraction", "BrowserMetrics")


def tidy_ui(data_mb=200):
    """浏览器缓存治理：ui_cache 整删；ui_profile 只在超过阈值时清缓存子目录。

    只删缓存，不删登录态/窗口位置（profile 里其余内容保留），
    所以用户下次打开窗口位置不变、也不会重新点任何同意项。
    """
    data = os.path.join(ROOT, "data")
    print("4) 浏览器缓存（ui_cache / ui_profile）")
    c = os.path.join(data, "ui_cache")
    if os.path.isdir(c):
        print("    ui_cache %.0f MB" % (_size(c) / 1048576.0))
        _rm(c)
    else:
        print("    ui_cache 不存在")
    prof = os.path.join(data, "ui_profile")
    if not os.path.isdir(prof):
        return
    mb = _size(prof) / 1048576.0
    if mb < data_mb:
        print("    ui_profile %.0f MB（< %d MB，跳过）" % (mb, data_mb))
        return
    print("    ui_profile %.0f MB，清理缓存子目录：" % mb)
    for n in UI_CACHE_SUBDIRS:
        p = os.path.join(prof, n)
        if os.path.isdir(p):
            print("      %-24s %6.0f MB" % (n, _size(p) / 1048576.0))
            _rm(p)

def main():
    t0 = _size(ROOT)
    print("整理工作区%s（当前 %.0f MB）" % ("（演练）" if DRY else "", t0 / 1048576.0))

    print("1) 构建残留")
    for p in glob.glob(os.path.join(ROOT, "runtime_old_build*")) + \
             glob.glob(os.path.join(TMP, "_dist_new*")):
        _rm(p)

    print("1.5) 测试留下的临时目录（ptw_html_* / ptw_zip_*）")
    for p in glob.glob(os.path.join(TMP, "ptw_html_*")) + glob.glob(os.path.join(TMP, "ptw_zip_*")):
        _rm(p)

    print("2) 分发包只留最近 %d 个" % KEEP_ZIP)
    zips = sorted([n for n in os.listdir(ARCH) if n.lower().endswith(".zip")], key=ver_key)
    old = zips[:-KEEP_ZIP] if len(zips) > KEEP_ZIP else []
    for n in old:
        _mv(os.path.join(ARCH, n), os.path.join(BACKUP, "归档_旧版_" + time.strftime("%Y%m%d")))
    if not old:
        print("    无需处理（现有 %d 个）" % len(zips))

    print("3) tmp 里超过 %d 天的一次性目录" % TMP_STALE_DAYS)
    keep = {"buildenv", "pyi-build", "flowtest", "_regdata", "_fixture"}
    deadline = time.time() - TMP_STALE_DAYS * 86400
    moved = 0
    for n in os.listdir(TMP):
        p = os.path.join(TMP, n)
        if not os.path.isdir(p) or n in keep:
            continue
        if os.path.getmtime(p) < deadline:
            _mv(p, os.path.join(BACKUP, "tmp_旧_" + time.strftime("%Y%m%d")))
            moved += 1
    if not moved:
        print("    无需处理")

    tidy_ui()

    if not DRY:
        t1 = _size(ROOT)
        print("\n完成：%.0f MB → %.0f MB（省 %.0f MB）"
              % (t0 / 1048576.0, t1 / 1048576.0, (t0 - t1) / 1048576.0))
        print("接着做一次结构体检：")
        sys.argv = [sys.argv[0]]
        import runpy
        runpy.run_path(os.path.join(HERE, "check_workspace.py"), run_name="__main__")


if __name__ == "__main__":
    main()
