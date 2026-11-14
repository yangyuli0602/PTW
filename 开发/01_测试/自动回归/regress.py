# -*- coding: utf-8 -*-
"""统一回归入口：一次跑完 开发/01_测试/自动回归 下的所有 test_*.py。

用法（平台根执行）：
    tmp/buildenv/Scripts/python.exe 开发/01_测试/自动回归/regress.py
    ... regress.py --only fa_swp            # 只跑名字匹配的（可多次给）
    ... regress.py --list                   # 只列清单
    ... regress.py --timeout 300            # 单个脚本超时（秒，默认 180）

判定：**退出码全为 0 才算通过**；打印"通过"但退出码非 0 的算失败。
所有输出落到 tmp/_regress/<名字>.log，失败时自动打印最后 15 行。
"""
import glob
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))     # 平台根
PY = sys.executable
LOGDIR = os.path.join(ROOT, "tmp", "_regress")


def _tests():
    return sorted(glob.glob(os.path.join(HERE, "test_*.py")))


def main():
    argv = sys.argv[1:]
    only = [argv[i + 1] for i, a in enumerate(argv) if a == "--only" and i + 1 < len(argv)]
    timeout = 180
    for i, a in enumerate(argv):
        if a == "--timeout" and i + 1 < len(argv):
            timeout = int(argv[i + 1])

    tests = _tests()
    if only:
        tests = [t for t in tests if any(o in os.path.basename(t) for o in only)]
    if not tests:
        print("没有匹配的测试脚本")
        return 1
    if "--list" in argv:
        for t in tests:
            print(" ", os.path.basename(t))
        return 0

    os.makedirs(LOGDIR, exist_ok=True)
    print("=" * 66)
    print("回归：%d 个脚本（单个超时 %ds）" % (len(tests), timeout))
    print("=" * 66)
    results, t0 = [], time.time()
    for t in tests:
        name = os.path.basename(t)[:-3]
        log = os.path.join(LOGDIR, name + ".log")
        st = time.time()
        try:
            with open(log, "w", encoding="utf-8", errors="replace") as f:
                rc = subprocess.run([PY, t], cwd=HERE, stdout=f, stderr=subprocess.STDOUT,
                                    timeout=timeout).returncode
        except subprocess.TimeoutExpired:
            rc = 124
            with open(log, "a", encoding="utf-8") as f:
                f.write("\n[regress] 超时 %ds\n" % timeout)
        dt = time.time() - st
        ok = rc == 0
        results.append((name, ok, rc, dt, log))
        print("  %-4s %-26s %5.1fs%s" % ("PASS" if ok else "FAIL", name, dt,
                                         "" if ok else "  rc=%s" % rc), flush=True)

    bad = [r for r in results if not r[1]]
    print("-" * 66)
    print("通过 %d / 失败 %d / 共 %d，用时 %.1fs"
          % (len(results) - len(bad), len(bad), len(results), time.time() - t0))
    if bad:
        for name, _ok, rc, _dt, log in bad:
            print("\n[%s] 最后 15 行（完整看 %s）：" % (name, os.path.relpath(log, ROOT)))
            try:
                with open(log, encoding="utf-8", errors="replace") as f:
                    for line in f.read().splitlines()[-15:]:
                        print("   ", line)
            except OSError:
                pass
        return 1
    print("全部通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
