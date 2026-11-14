# -*- coding: utf-8 -*-
"""提交前结构体检 —— 防「乱存文件」和「规则文件膨胀」。

为什么要有它：AI 每次会话都是新来的，靠"文档里写清楚"约束不住——
想起来才遵守，忘了就随手在根目录建 md、在 tmp 堆脚本、把 MEMORY 越写越长。
这里把四条规矩变成**提交时的硬拦截**，不靠自觉。

五查（任一不过即拒绝提交）：
  1. 新增文件必须落在目录白名单内（根目录零新增）
  2. MEMORY.md ≤ 150 行
  3. .workbuddy/memory/ 顶层日志 ≤ 5 个（多的按日期归档到 _archive/）
  4. 03_文档 根目录 md ≤ 30 个（防野文件堆积）
  5. 不得引用工作区之外的路径（2026-11-13 用户红线：只允许在 01_PTW框架 内工作）

用 `git commit --no-verify` 可强行绕过——那是人为故意，请在提交说明里写原因。
"""
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 允许新增文件的目录（以 / 结尾表示目录）
ALLOW_DIRS = (
    "开发/", "02_备份/", "04_归档/", ".workbuddy/", ".githooks/",
    "tmp/", "data/",
)
# 允许在根目录存在的文件（固定入口，不许新增其它）
ALLOW_ROOT = {
    ".gitignore", ".gitattributes", "README.md", "启动说明.txt", "AGENTS.md",
}

LIMITS = {
    "memory_md_lines": 150,
    "memory_logs": 5,
    "doc_md": 30,
}

# 工作区越界引用：出现即拒绝。2026-11-13 用户红线——只允许在工作区内工作。
# 不是查"有没有访问过"（查不到），而是查"代码/文档里有没有把工作区外的路径写死"——
# 写死的路径就是下一次会话、下一个人的默认行为。
OUTSIDE_PAT = re.compile(
    r"(?:FDE[\\/]+工作[\\/]+(?!01_PTW框架))"      # D:\桌面\FDE工作\ 后面不是本工作区
    r"|02_PTW无skill测试"                         # 工作区外的旧工程目录
    r"|10_工具_非PTW"                             # 工作区外的工具目录
)
# 只查会被人照着执行的文本；memory 是历史日志、tmp 是产物、02_备份/04_归档 是回滚与交付件
OUTSIDE_SKIP_DIRS = (".workbuddy/memory/", "tmp/", "02_备份/", "04_归档/")
OUTSIDE_SKIP_FILES = (".githooks/check_structure.py", "开发/.githooks/check_structure.py")   # 规则自身的注释，不算引用
OUTSIDE_EXTS = (".py", ".md", ".json", ".js", ".bat", ".ps1", ".sh", ".txt", ".css", ".html")

fails = []


def run(*args):
    # core.quotepath=false：否则中文路径会被 git 转成 "/346/240/271" 这种八进制，报错看不懂
    r = subprocess.run(("git", "-c", "core.quotepath=false") + args,
                       cwd=ROOT, capture_output=True)
    if r.returncode != 0:
        # 体检自身的命令失败绝不能当成"通过"——那等于门禁形同虚设
        fails.append("结构体检执行失败（git %s 返回 %d）：%s"
                     % (" ".join(args[:2]), r.returncode,
                        (r.stderr or b"").decode("utf-8", "replace")[:120]))
        return ""
    return (r.stdout or b"").decode("utf-8", "replace")


def check_new_files():
    """新增文件必须落在白名单目录内。"""
    # 注意：run() 内部已经拼了 git，这里只传子命令
    out = run("diff", "--cached", "--name-status", "--diff-filter=A")
    bad = []
    for line in out.splitlines():
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        path = parts[-1].replace("\\", "/")
        if path in ALLOW_ROOT:
            continue
        if any(path.startswith(d) for d in ALLOW_DIRS):
            continue
        bad.append(path)
    if bad:
        fails.append("根目录/白名单外新增了 %d 个文件（只允许放在 %s）：\n    %s"
                     % (len(bad), " / ".join(d.rstrip("/") for d in ALLOW_DIRS),
                        "\n    ".join(bad[:10])))


def count_lines(path):
    try:
        with open(os.path.join(ROOT, path), encoding="utf-8") as f:
            return sum(1 for _ in f)
    except OSError:
        return 0


def check_memory_size():
    n = count_lines(os.path.join(".workbuddy", "memory", "MEMORY.md"))
    if n > LIMITS["memory_md_lines"]:
        fails.append("MEMORY.md 已 %d 行（上限 %d）：请把「只在改某模块才需要」的内容"
                     "移到 03_文档/项目约定_实现细节_*.md，本文件只留每次都必须知道的。"
                     % (n, LIMITS["memory_md_lines"]))


def check_memory_logs():
    d = os.path.join(ROOT, ".workbuddy", "memory")
    if not os.path.isdir(d):
        return
    logs = [f for f in os.listdir(d) if f.endswith(".md") and f != "MEMORY.md"]
    if len(logs) > LIMITS["memory_logs"]:
        fails.append("memory/ 顶层有 %d 个工作日志（上限 %d）：把旧的合并进 "
                     "memory/_archive/ 后再提交。示例：%s"
                     % (len(logs), LIMITS["memory_logs"], "、".join(sorted(logs)[:3])))


def check_docs():
    d = os.path.join(ROOT, "03_文档")
    if not os.path.isdir(d):
        return
    mds = [f for f in os.listdir(d) if f.endswith(".md")]
    if len(mds) > LIMITS["doc_md"]:
        fails.append("03_文档 根目录有 %d 个 md（上限 %d）：版本类/一次性文档请移进 "
                     "_历史/，首页只放当前有效的。" % (len(mds), LIMITS["doc_md"]))


def check_outside_refs():
    """代码/文档里不许写死工作区外的路径。

    查的是"写死"，不是"访问过"——但差别在后果：写死的路径会被下一次会话、
    下一个人原样执行，工作区就慢慢变成"备份区 + 工程目录"的混合体，
    数据到底哪份是准的都说不清（2026-11-13 瑞云复跑就踩过一次）。
    """
    out = run("diff", "--cached", "--name-only", "--diff-filter=ACM")
    hits = []
    for path in [p.strip().replace("\\", "/") for p in out.splitlines() if p.strip()]:
        if not path.lower().endswith(OUTSIDE_EXTS):
            continue
        if path in OUTSIDE_SKIP_FILES:
            continue
        if any(path.startswith(d) for d in OUTSIDE_SKIP_DIRS):
            continue
        full = os.path.join(ROOT, path)
        if not os.path.isfile(full):
            continue
        try:
            with open(full, encoding="utf-8", errors="replace") as fh:
                for no, line in enumerate(fh, 1):
                    if OUTSIDE_PAT.search(line):
                        hits.append("%s:%d  %s" % (path, no, line.strip()[:90]))
        except OSError:
            continue
    if hits:
        fails.append("有 %d 处引用工作区外的路径（红线：只允许在 01_PTW框架 内工作）：\n    %s\n"
                     "    改法：路径一律用工作区内的相对路径；历史来源只写「11-13 备份区」这类描述，"
                     "不写可复制的绝对路径。" % (len(hits), "\n    ".join(hits[:8])))


def main():
    check_new_files()
    check_memory_size()
    check_memory_logs()
    check_docs()
    check_outside_refs()
    if fails:
        print("\n[结构体检] 拒绝提交：\n")
        for i, f in enumerate(fails, 1):
            print("  %d) %s\n" % (i, f))
        print("  豁免：git commit --no-verify（请在提交说明里写明原因）\n")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
