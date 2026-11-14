# -*- mode: python ; coding: utf-8 -*-
"""启动器打包：单文件 PTW.exe，放在平台根，永不更新。"""
import os

HERE = os.path.abspath(SPECPATH)      # SPECPATH 本身就是 spec 所在目录

a = Analysis(
    [os.path.join(HERE, "launcher.py")],
    pathex=[HERE],
    binaries=[],
    datas=[],
    hiddenimports=["tkinter", "tkinter.messagebox"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["duckdb", "openpyxl", "pandas", "numpy", "PIL", "matplotlib"],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="PTW",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
