@echo off
setlocal
cd /d "%~dp0"
set PY=..\..\tmp\buildenv\Scripts\python.exe
if not exist "%PY%" goto NOPY
echo ============================================================
echo  PTW 发版（打包 + GitHub Release + 更新清单）
echo  版本号取 开发\程序\app_version.py 的 APP_VERSION
echo ============================================================
echo.
set NOTES=
set /p NOTES=本次更新说明（会显示给用户，可留空）: 
echo.
if "%NOTES%"=="" goto PLAIN
"%PY%" publish.py --notes "%NOTES%"
goto END

:PLAIN
"%PY%" publish.py
goto END

:NOPY
echo [错误] 找不到打包环境：%PY%

:END
echo.
pause
