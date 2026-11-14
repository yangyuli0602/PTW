# -*- coding: utf-8 -*-
"""PTW 启动器（PTW.exe）—— 永不更新的那一层。

职责只有三件：
  1. 拉起主程序 runtime/PTWApp.exe
  2. 有更新待应用时，等主程序退出 → 整体替换 runtime/ → 再拉起
  3. 记录 launcher 自己的日志（更新失败要能查）

为什么要 launcher：Windows 不能覆盖正在运行的程序。把「可更新部分」全部放进
runtime/，替换时目录整体改名即可，不碰任何被占用的文件。

用法：
  PTW.exe                 正常启动（有 update_pending.json 就先应用更新）
  PTW.exe --apply-update  主程序调用：等它退出后替换并重启
  PTW.exe --version       打印版本
"""
import io
import json
import os
import shutil
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False)
                                        else __file__))
RUNTIME = os.path.join(ROOT, "runtime")
STAGE = os.path.join(ROOT, "runtime_new")
BACKUP = os.path.join(ROOT, "runtime_old")
PENDING = os.path.join(ROOT, "update_pending.json")
VERFILE = os.path.join(ROOT, "version.json")
APP_EXE = os.path.join(RUNTIME, "PTWApp.exe")
LOG = os.path.join(ROOT, "data", "logs", "launcher.log")
DETACH = 0x00000008          # DETACHED_PROCESS
NO_WINDOW = 0x08000000       # CREATE_NO_WINDOW


def log(msg):
    try:
        os.makedirs(os.path.dirname(LOG), exist_ok=True)
        with io.open(LOG, "a", encoding="utf-8") as f:
            f.write("%s %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass


def read_json(p, default=None):
    try:
        with io.open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def write_json(p, obj):
    with io.open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def launch_app(wait=False, no_ui=False):
    """拉起主程序；wait=True 时阻塞到它退出（launcher 常驻模式）。

    常驻模式的理由：更新时主程序只需「下完包、写标记、自己退出」，
    launcher 一定在场，能立刻接手替换。不依赖主程序反过来启动 launcher——
    那条路要处理 onefile 解压耗时、父进程退出等一堆时序问题（踩过）。
    """
    if not os.path.exists(APP_EXE):
        log("!! 找不到主程序：%s" % APP_EXE)
        _msg("找不到主程序 runtime\\PTWApp.exe。\n请重新解压完整安装包。")
        return None
    env = dict(os.environ)
    env.setdefault("KZ_DATA_DIR", os.path.join(ROOT, "data"))
    env["PTW_LAUNCHED"] = "1"          # 主程序据此知道「有人等我退出，直接退就行」
    if no_ui:
        # 更新后重启：旧窗口会自己轮询刷新回来，别急着重开一个（会变成两个窗口）
        env["PTW_NO_UI"] = "1"
    try:
        p = subprocess.Popen([APP_EXE], cwd=RUNTIME, env=env, close_fds=True,
                             creationflags=NO_WINDOW)
        log("已启动主程序（PID %d）" % p.pid)
        if wait:
            p.wait()
            log("主程序已退出（code=%s）" % p.returncode)
        return p
    except Exception as e:
        log("!! 启动失败：%s" % e)
        _msg("启动失败：%s" % e)
        return None


def _msg(text, title="PTW 平台"):
    """无控制台时用一个最小弹窗告知（tkinter 是标准库）。"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk(); r.withdraw()
        messagebox.showinfo(title, text)
        r.destroy()
    except Exception:
        log("提示：" + text.replace("\n", " "))


def wait_main_exit(timeout=180):
    """等主程序退出：端口不通 或 进程已不存在。"""
    sys.path.insert(0, RUNTIME)
    lock = os.path.join(ROOT, "data", "server.lock")
    t0 = time.time()
    while time.time() - t0 < timeout:
        info = read_json(lock, {}) or {}
        pid = info.get("pid")
        if not pid:
            return True                     # 锁没了 = 已退出
        if not _pid_alive(pid):
            return True
        time.sleep(0.7)
    return False


def _pid_alive(pid):
    try:
        import ctypes
        h = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))
        if not h:
            return False
        ctypes.windll.kernel32.CloseHandle(h)
        return True
    except Exception:
        return False


def apply_update():
    """替换 runtime：runtime → runtime_old，runtime_new → runtime。失败回滚。"""
    doc = read_json(PENDING, {}) or {}
    log("开始应用更新 → %s（来自 %s）" % (doc.get("version"), doc.get("from")))
    if not os.path.isdir(STAGE):
        log("!! 暂存目录不存在：%s" % STAGE)
        _msg("更新文件缺失，已取消。请重新点「检查更新」。")
        try: os.remove(PENDING)
        except OSError: pass
        return 1

    # 清掉上次残留的 backup
    if os.path.isdir(BACKUP):
        shutil.rmtree(BACKUP, ignore_errors=True)

    # 改名可能因文件锁失败 → 重试
    ok = False
    for i in range(60):
        try:
            if os.path.isdir(RUNTIME):
                os.rename(RUNTIME, BACKUP)
            os.rename(STAGE, RUNTIME)
            ok = True
            break
        except OSError as e:
            if i == 0:
                log("  等待主程序释放文件…（%s）" % e)
            time.sleep(1)
    if not ok:
        log("!! 替换失败（60s 仍未释放）")
        if not os.path.isdir(RUNTIME) and os.path.isdir(BACKUP):
            try: os.rename(BACKUP, RUNTIME)
            except OSError: pass
        _msg("更新失败：程序文件被占用。\n请关闭所有 PTW 窗口后重新打开。")
        return 1

    write_json(VERFILE, {"current": doc.get("version") or "",
                         "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")})
    try: os.remove(PENDING)
    except OSError: pass
    log("更新完成，当前版本 %s" % doc.get("version"))
    return 0


def main():
    arg = (sys.argv[1] if len(sys.argv) > 1 else "").strip().lower()
    if arg in ("--version", "-v"):
        v = (read_json(VERFILE, {}) or {}).get("current", "?")
        print(v)
        return 0

    if arg == "--apply-update":
        # 兼容「直接双击主程序」的场景：主程序会自己拉起本进程
        log("=== launcher --apply-update ===")
        if not wait_main_exit():
            log("!! 等主程序退出超时（180s），放弃本次替换")
            _msg("更新超时：主程序没有退出。\n请手动关闭 PTW 后再次打开，会自动完成更新。")
            return 1
        rc = apply_update()
        if rc == 0:
            launch_app(wait=True, no_ui=True)
        return rc

    # ---------------- 常驻模式 ----------------
    log("=== launcher start（常驻）===")
    if os.path.isdir(BACKUP):
        shutil.rmtree(BACKUP, ignore_errors=True)
    if os.path.exists(PENDING):
        apply_update()
    no_ui = False      # 更新后重启的那一轮不新开窗口（旧窗口会自己刷回来）
    while True:
        p = launch_app(wait=True, no_ui=no_ui)
        if p is None:
            return 1
        # 主程序退出：如果它留了更新标记，就地替换后再拉起
        if os.path.exists(PENDING):
            log("检测到更新标记，开始替换")
            no_ui = True
            if apply_update() != 0:
                log("替换失败，按旧版本启动")
            continue
        log("主程序正常退出，launcher 结束")
        return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        log("!! launcher 异常：%s" % e)
        import traceback
        log(traceback.format_exc())
        raise
