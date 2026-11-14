# -*- coding: utf-8 -*-
"""「需要人或外部 AI 补的东西」的统一登记与落盘。

为什么要有它
------------
这类东西原先只写在 warnings 和批注里，调用方（尤其是外部 AI）很容易漏读；
而 warnings 是给人看的散文，没法机器核对"结清了没有"。

货币资金那套（C.02a 汇率公告截图）验证过有效：**任务卡自包含**——去哪取、取什么、
存到哪个路径、怎么自检、做完做什么，全部写死；AI 换一台电脑、没有任何项目上下文，
照着卡也能做完。费用这边的 OAR 口径判定与两段分析文字，同样按这套办。

两件事分开：
  · `add_task()`     把任务登记进生成结果的 `tasks[]`（结构化，可机器核对）
  · `task_file()`    把待办清单写成固定路径的 JSON，AI 只读它、只写结果文件
"""
import json
import os

#: 待办清单目录名（相对 <data>/bkd_work/）
TASKS_DIRNAME = "ai_tasks"

#: 结果文件的统一后缀：写在原文件旁边，加载时自动合并。
#: 让 AI 写 `xxx_result.json` 而不是直接改 `xxx.json`，是为了不覆盖已有的
#: 人工/AI 结论（合并语义 = 同键覆盖，但旧键保留）。
RESULT_SUFFIX = "_result"


def add_task(res, id="", type="manual", title="", target="", saveTo="",
             url="", prompt="", blocking=False, afterDone="",
             readPath="", expectCount=None):
    """把一条待办登记进 res["tasks"]（同 id + 同 target 只留一条）。

    blocking=True 表示不结清就不能出稿；费用这边都是 False（缺 AI 也能出稿，
    只是文字退化成规则版），靠 warnings 让调用方知道"还有几件事没做"。
    """
    t = {"id": id or "", "type": type, "title": title, "target": target,
         "saveTo": saveTo, "url": url, "prompt": prompt, "blocking": bool(blocking)}
    # 清单路径与应有条数：调用方（外部 AI）据此直接读、并在写完后核对条数，
    # 不用从 title 里正则抠数字。加字段不改语义，老调用方忽略即可。
    if readPath:
        t["readPath"] = readPath
    if expectCount is not None:
        t["expectCount"] = int(expectCount)
    if afterDone:
        t["afterDone"] = afterDone
    lst = res.setdefault("tasks", [])
    if not any((x.get("id") == t["id"] and x.get("target") == t["target"]) for x in lst):
        lst.append(t)
    return t


def tasks_dir(data_dir):
    p = os.path.join(data_dir, "bkd_work", TASKS_DIRNAME)
    try:
        os.makedirs(p, exist_ok=True)
    except Exception:
        return ""
    return p


def task_file(data_dir, name, payload):
    """把 payload 写进 <data>/bkd_work/ai_tasks/<name>，返回路径（失败返回空）。"""
    d = tasks_dir(data_dir)
    if not d:
        return ""
    p = os.path.join(d, name)
    try:
        with open(p, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=1)
    except Exception:
        return ""
    return p


def result_path(base_path):
    """`<x>.json` → `<x>_result.json`（AI 该写的那个文件）。"""
    if not base_path:
        return ""
    root, ext = os.path.splitext(base_path)
    return root + RESULT_SUFFIX + (ext or ".json")


def load_side(path):
    """读 AI 写的结果侧文件：{"items":[...]} 或 {"键": "文字"}，读不到返回空。"""
    if not path or not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            d = json.load(fh)
    except Exception:
        return {}
    return d if isinstance(d, dict) else {}
