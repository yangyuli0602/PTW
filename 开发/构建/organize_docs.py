# -*- coding: utf-8 -*-
"""把 03_文档 的散装 skill 收进 skills/，按「用途」分目录，并自动改写互相引用。

设计：
  03_文档/
  ├─ SKILL.md            ← 调度入口（只做分类 + 指路，不复述细节）
  ├─ skills/
  │   ├─ 平台运行/        平台本身怎么用（导入清洗、达标判定、复跑）
  │   ├─ 底稿/            每张表怎么做（细节全留，别精简）
  │   ├─ 开发/            改代码前读的约定
  │   └─ 历史版本说明/      V1.2.x 版本说明（旧，但留档）
  └─ _历史/              更早的历史文档（原样不动）

为什么保留原文件名：正文里的互相引用全是「裸文件名」，
同目录引用天然有效；跨目录的由本脚本统一改写成相对路径。
用法：
    python organize_docs.py --dry-run     # 只打印计划
    python organize_docs.py               # 真执行（git mv + 改写引用）
"""
import io
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))          # 平台根
DOCS = os.path.join(ROOT, "开发", "03_文档")
DRY = "--dry-run" in sys.argv

GROUPS = {
    "平台运行": ["01_平台使用说明.md", "02_导入清洗.md", "03_清洗达标提示词.md",
              "_参考_清洗达标规则_逐字段.md", "复跑样本放在哪.md"],
    "底稿": ["00_总则.md", "04_货币资金_总览.md", "05_货币资金_BKD.md",
           "06_货币资金_Lead.md", "07_货币资金_C01b.md", "08_货币资金_C02a.md",
           "09_货币资金_C03.md", "_参考_货币资金口径.md", "AI提示词_口径判断.md",
           "10_费用.md", "11_固定资产.md"],
    "开发": ["项目约定_实现细节_20261106.md"],
}
VER_PREFIX = re.compile(r"^V\d+\.\d+\.")


def plan():
    """返回 {旧文件名: 新相对路径（相对 DOCS）}"""
    m = {}
    for g, names in GROUPS.items():
        for n in names:
            m[n] = os.path.join("skills", g, n)
    for n in sorted(os.listdir(DOCS)):
        if n.endswith(".md") and VER_PREFIX.match(n):
            m[n] = os.path.join("skills", "历史版本说明", n)
    return m


def target(abs_dir):
    os.makedirs(abs_dir, exist_ok=True)


def move(mapping):
    for old, new in mapping.items():
        src = os.path.join(DOCS, old)
        dst = os.path.join(DOCS, new)
        if not os.path.exists(src):
            print("  !! 源文件不存在：%s" % old)
            continue
        if DRY:
            print("  %-40s → %s" % (old, new))
            continue
        target(os.path.dirname(dst))
        r = subprocess.run(["git", "-C", ROOT, "mv", src, dst],
                           capture_output=True, text=True)
        if r.returncode != 0:
            subprocess.run(["cmd", "/c", "move", src, dst], capture_output=True)
        print("  %-40s → %s" % (old, new))


def rewrite(mapping):
    """把正文里的裸文件名改成相对当前文件位置的路径。"""
    total = 0
    for old, new in mapping.items():
        dst = os.path.join(DOCS, new)
        if not os.path.exists(dst):
            continue
        here = os.path.dirname(dst)
        s = io.open(dst, encoding="utf-8").read()
        orig = s
        for name, tgt in mapping.items():
            if name == old:
                continue
            rel = os.path.relpath(os.path.join(DOCS, tgt), here).replace("\\", "/")
            # 只替换「裸文件名」或「文档/前缀」两种写法，避免二次改写
            s = re.sub(r"(?<![/\w])" + re.escape(name), rel, s)
        if s != orig:
            n = sum(1 for a, b in zip(orig.splitlines(), s.splitlines()) if a != b)
            total += n
            if not DRY:
                io.open(dst, "w", encoding="utf-8").write(s)
            print("  改写 %-36s %d 行" % (new, n))
    print("  合计改写 %d 行" % total)


def verify():
    """校验：所有 .md 引用都能在磁盘上找到。"""
    bad = []
    for root, _d, files in os.walk(DOCS):
        for f in files:
            if not f.endswith(".md"):
                continue
            p = os.path.join(root, f)
            s = io.open(p, encoding="utf-8", errors="replace").read()
            for ref in re.findall(r"[\w\u4e00-\u9fa5_\-/]+\.md", s):
                if ref.startswith(("http", "MEMORY")):
                    continue
                if not os.path.exists(os.path.join(root, ref)):
                    bad.append((os.path.relpath(p, DOCS), ref))
    if bad:
        print("\n[引用校验] %d 处找不到目标：" % len(bad))
        for a, b in bad[:20]:
            print("  %s → %s" % (a, b))
    else:
        print("\n[引用校验] 全部引用均可解析。")


def main():
    m = plan()
    print("计划：移动 %d 个文件%s" % (len(m), "（演练）" if DRY else ""))
    move(m)
    if not DRY:
        print("\n改写跨目录引用：")
        rewrite(m)
        verify()


if __name__ == "__main__":
    main()
