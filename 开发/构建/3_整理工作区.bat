@echo off
setlocal
cd /d "%~dp0"
set PY=..\..\tmp\buildenv\Scripts\python.exe
if not exist "%PY%" goto NOPY
echo ============================================================
echo  整理工作区（清构建残留 / 归档只留最近 3 版 / tmp 清理）
echo ============================================================
echo.
echo 先演练看会动什么（不改任何东西）：
"%PY%" tidy_workspace.py --dry-run
echo.
set GO=
set /p GO=确认执行吗？输入 y 回车执行，其它键退出: 
if /i not "%GO%"=="y" goto END
echo.
"%PY%" tidy_workspace.py
goto END

:NOPY
echo [错误] 找不到打包环境：%PY%

:END
echo.
pause
