@echo off
setlocal
cd /d "%~dp0"
set PY=..\..\tmp\buildenv\Scripts\python.exe
if not exist "%PY%" goto NOPY
if not exist "%USERPROFILE%\.workbuddy\ptw_token.txt" goto NOTOKEN
echo ============================================================
echo  PTW 首次建仓库（GitHub）
echo ============================================================
echo.
"%PY%" publish.py --setup
goto END

:NOPY
echo [错误] 找不到打包环境：%PY%
echo 先按 README 建 tmp\buildenv
goto END

:NOTOKEN
echo [需要先做一步] 没找到 token 文件。
echo   1) 打开 GitHub - Settings - Developer settings - Personal access tokens
echo   2) 生成 Tokens (classic)，勾选 repo
echo   3) 把 token 粘进这个文件（只放一行）：
echo      %USERPROFILE%\.workbuddy\ptw_token.txt
echo   4) 再双击本脚本
goto END

:END
echo.
pause
