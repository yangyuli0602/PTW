# -*- coding: utf-8 -*-
"""统一路径解析 —— 让「源码运行」与「PyInstaller 打包运行」走同一套逻辑。

为什么要单独一个模块：原来 APP_DIR / KZ_DATA 的计算散在 server.py、core/db.py、
modules/bkd/core.py 三处，各自用 `__file__ + ../..` 推，打包后（`_MEIPASS` 布局）
会推出错误的目录。集中到一处，改一次就够。

三种运行态：
  源码   python 程序/server.py            APP_DIR=程序/   DATA=程序/../data
  测试   KZ_DATA_DIR=tmp/xxx              DATA 由环境变量指定（优先级最高）
  打包   runtime/PTWApp.exe               APP_DIR=_MEIPASS  DATA=平台根/data
         其中：_MEIPASS = runtime/_internal（web、modules 都在它下面）
"""
import os
import sys

#: 是否运行在 PyInstaller 打包产物里
FROZEN = bool(getattr(sys, "frozen", False))

#: 静态资源根（web/、modules/、rules/ 的父目录）
APP_DIR = (getattr(sys, "_MEIPASS", None) or
           os.path.dirname(os.path.abspath(__file__)) + os.sep + "..")
APP_DIR = os.path.abspath(APP_DIR)

#: 可执行文件所在目录（未打包时是 程序/ 的父目录，仅用于日志提示）
EXE_DIR = os.path.dirname(os.path.abspath(sys.executable)) if FROZEN else os.path.abspath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))


def _default_data_dir():
    """未指定 KZ_DATA_DIR 时的数据目录。

    打包态：运行时目录（runtime/）的同级 data/ —— 即「平台根/data/」，
            程序升级替换 runtime/ 时数据不受影响。
    源码态：程序/../data（与原行为一致）。
    """
    if FROZEN:
        return os.path.join(os.path.dirname(EXE_DIR), "data")
    return os.path.abspath(os.path.join(APP_DIR, "..", "data"))


#: 数据目录（KZ_DATA_DIR 环境变量优先，测试隔离靠它）
DATA_DIR = os.path.abspath(os.environ.get("KZ_DATA_DIR") or _default_data_dir())

#: 运行数据子目录
TPL_DIR = os.path.join(DATA_DIR, "bkd_tpl")
LOG_DIR = os.path.join(DATA_DIR, "logs")
DB_PATH = os.path.join(DATA_DIR, "ledger.duckdb")


def ensure_dirs(*names):
    """确保数据目录下的子目录存在，返回绝对路径列表。"""
    out = []
    for n in names:
        p = os.path.join(DATA_DIR, n)
        os.makedirs(p, exist_ok=True)
        out.append(p)
    return out


def describe():
    """启动横幅用：一眼看出实际用的是哪套目录。"""
    return ("运行态=%s | 程序=%s | 数据=%s"
            % ("打包" if FROZEN else "源码", APP_DIR, DATA_DIR))


def seed_templates():
    """首次启动把内置模板释放到 data/bkd_tpl/（已存在的不覆盖）。

    模板是程序资产（`程序/tpl/`，打包后在 `_internal/bkd_tpl_seed/`），
    数据目录属于用户——升级程序不会冲掉客户改过的模板。
    返回本次释放的文件名列表。
    """
    import shutil
    seed = os.path.join(APP_DIR, "bkd_tpl_seed")
    if not os.path.isdir(seed):
        seed = os.path.join(APP_DIR, "tpl")         # 源码态直接用 程序/tpl
    if not os.path.isdir(seed):
        return []
    os.makedirs(TPL_DIR, exist_ok=True)
    out = []
    for f in os.listdir(seed):
        if not f.lower().endswith((".xlsx", ".xlsm")):
            continue
        dst = os.path.join(TPL_DIR, f)
        if os.path.exists(dst):
            continue
        try:
            shutil.copy2(os.path.join(seed, f), dst)
            out.append(f)
        except OSError:
            pass
    return out
