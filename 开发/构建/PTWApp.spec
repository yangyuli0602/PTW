# -*- mode: python ; coding: utf-8 -*-
"""主程序打包：onedir，输出到 平台根/runtime/。

为什么 onedir 不是 onefile：
  · server.py 用 os.listdir(modules/) 动态扫描板块，需要真实目录；
  · onefile 每次启动都要把 ~150MB 解压到临时目录，慢且被杀软重点照顾；
  · 更新时整体替换 runtime/ 目录即可，不涉及单文件替换。

datas 里带 modules / web / rules：这些是「运行时从文件系统加载」的，
不能只塞进 pyz——带出来才能加板块不改代码、换前端不重打包。
"""
import os

HERE = os.path.abspath(SPECPATH)      # SPECPATH 本身就是 spec 所在目录
APP = os.path.abspath(os.path.join(HERE, "..", "程序"))

datas = [
    (os.path.join(APP, "web"), "web"),
    (os.path.join(APP, "modules"), "modules"),
    (os.path.join(APP, "rules"), "rules"),
    (os.path.join(APP, "tpl"), "bkd_tpl_seed"),      # 内置模板：首次启动释放到 data/bkd_tpl
]
datas = [(p, d) for p, d in datas if os.path.isdir(p)]

hidden = [
    "duckdb", "openpyxl", "et_xmlfile",
    "urllib.request", "http.client", "ssl", "_ssl", "certifi",
    "core.paths", "core.logutil", "core.ptw_ai", "core.updater",
    "core.audit", "core.clean", "core.recon", "core.db", "core.fields",
    "core.recipe", "core.xlsread", "core.xlsxinfo", "core.htmlread",
    "app_version",
]

a = Analysis(
    [os.path.join(APP, "server.py")],
    pathex=[APP],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter.test", "test", "unittest", "pydoc_data",
              "pandas", "numpy", "scipy", "matplotlib", "PIL", "IPython",
              "notebook", "pytest", "bokeh", "lxml", "cryptography"],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PTWApp",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # duckdb 的 .pyd 用 upx 压过容易出问题，不压
    console=False,             # 无控制台；日志一律落 data/logs/
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="runtime",
)
