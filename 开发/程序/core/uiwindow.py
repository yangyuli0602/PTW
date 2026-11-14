# -*- coding: utf-8 -*-
"""把界面以「独立窗口」打开，而不是一个浏览器标签。

优先用机器上已装的 Edge / Chrome 的 --app 模式：
  - 没有地址栏、书签栏、标签页，任务栏是独立窗口；
  - 用独立 user-data-dir（放在数据目录下），所以不会跟你日常用的浏览器混在一起，
    关掉窗口就是关掉它，不会连带关掉你正在看的网页。
找不到浏览器时退回系统默认浏览器（行为与旧版一致，不会因此打不开）。
"""
import os
import subprocess
import sys
import webbrowser

from core import paths


def _candidates():
    """按可能性排序列出浏览器 exe 路径（只列存在性检查用，不写死优先级到某台机器）。"""
    env = os.environ
    pf = env.get("ProgramFiles", r"C:\Program Files")
    pf86 = env.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    local = env.get("LOCALAPPDATA", "")
    out = []
    for base in (pf86, pf, local):
        if base:
            out.append(os.path.join(base, r"Microsoft\Edge\Application\msedge.exe"))
    for base in (pf, pf86, local):
        if base:
            out.append(os.path.join(base, r"Google\Chrome\Application\chrome.exe"))
    # 便携版常见位置（尽力而为，不存在就跳过）
    out.append(os.path.join(local, r"Chromium\Application\chrome.exe"))
    if local:
        out.append(os.path.join(local, r"Microsoft\Edge SxS\Application\msedge.exe"))
    return [p for p in out if p and os.path.exists(p)]


def profile_dir():
    """独立配置目录：放在数据目录下，不污染用户自己的浏览器配置。"""
    d = os.path.join(paths.DATA_DIR, "ui_profile")
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        return None
    return d


def open_ui(url, log=None):
    """打开界面，返回 'app:<浏览器名>' 或 'browser'。任何异常都退回默认浏览器。"""
    def _say(msg):
        if log:
            try:
                log(msg)
            except Exception:
                pass

    if os.name == "nt":
        prof = profile_dir()
        for exe in _candidates():
            args = [exe, "--app=" + url, "--no-first-run",
                    "--no-default-browser-check", "--disable-session-crashed-bubble",
                    "--window-size=1440,900"]
            if prof:
                args.insert(2, "--user-data-dir=" + prof)
            try:
                subprocess.Popen(args, close_fds=True,
                                 creationflags=getattr(subprocess, "DETACHED_PROCESS", 0))
                name = "Edge" if "msedge" in exe.lower() else "Chrome"
                _say("已用 %s 独立窗口打开界面" % name)
                return "app:" + name
            except Exception as e:      # 例如被策略禁止启动
                _say("用 %s 打开失败（%s），继续尝试下一个" % (os.path.basename(exe), e))
    webbrowser.open(url)
    _say("已用系统默认浏览器打开界面")
    return "browser"
