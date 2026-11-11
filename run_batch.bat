@echo off
chcp 936 >nul 2>nul
set "PYTHONIOENCODING=gbk"
setlocal

rem =====================================================================
rem  dwg-material-to-xls  -  portable drag & drop batch runner (no AI)
rem  Drop one or more folders (or .dwg files) onto this .bat file.
rem  Scripts and templates live inside this folder: copy the whole folder anywhere.
rem
rem  Target PC needs (checked on startup):
rem    - Python 3.8+        (auto-installed via winget if available)
rem    - pywin32            (auto-installed via pip if missing)
rem    - COM-capable CAD: GStarCAD / AutoCAD / ZWCAD
rem    - Microsoft Excel or WPS Spreadsheets (COM), to write .xls output
rem
rem  NOTE: no ( ... ) blocks around %* (dropped args). A ")" in a file name
rem  would close such a block early and crash the batch. All branches use goto.
rem
rem  Optional environment variables:
rem    DWG_BATCH_NOPAUSE=1   no pause when finished (scheduled / scripted runs)
rem    DWG_BATCH_DRYRUN=1    show plan only, do not write .xls
rem    DWG_BATCH_NOCACHE=1   always re-read CAD (ignore cache)
rem    DWG_BATCH_NODEPS=1    skip startup dependency self-check
rem =====================================================================

title DWG 材料平差批处理 v2.4.6
echo  DWG 材料平差批处理 v2.4.6  -  作者 lonyii  lonyii@qq.com

set "HERE=%~dp0"
set "SCRIPT=%HERE%scripts\auto_export.py"
set "DEPS=%HERE%scripts\check_deps.py"
set "CACHE=%HERE%_cache"
set "LOG=%HERE%run_batch.log"

if "%~1"=="" goto NOARGS
if not exist "%SCRIPT%" goto NOSCRIPT

rem --- cache dir INSIDE this folder: portable, no user-profile writes ---
if not exist "%CACHE%" mkdir "%CACHE%" >nul 2>nul

rem --- locate Python 3 ---------------------------------------------------
set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY where python >nul 2>nul
if not errorlevel 1 if not defined PY set "PY=python"
if not defined PY goto AUTOPY
goto AFTERPY

rem --- try to auto-install Python via winget -----------------------------
:AUTOPY
echo  未找到 Python 3，正在尝试自动安装（winget）...
where winget >nul 2>nul
if errorlevel 1 goto NOPY
winget install -e --id Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY where python >nul 2>nul
if not errorlevel 1 if not defined PY set "PY=python"
if not defined PY goto NOPY
echo  Python 已安装。若 COM 组件异常，请重新打开窗口后再拖放。
goto AFTERPY

:AFTERPY

rem --- output folder: Desktop\材料平差输出\<date-time> ---------------
set "DESKTOP=%USERPROFILE%\Desktop"
if not exist "%DESKTOP%" set "DESKTOP=%USERPROFILE%\桌面"
for /f "usebackq delims=" %%D in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do set "DESKTOP=%%D"
set "OUTROOT=%DESKTOP%\材料平差输出"
for /f "usebackq delims=" %%T in (`powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"`) do set "STAMP=%%T"
if not defined STAMP set "STAMP=unknown"
set "OUT=%OUTROOT%\%STAMP%"
if not exist "%OUTROOT%" mkdir "%OUTROOT%" >nul 2>nul
if not exist "%OUT%" mkdir "%OUT%" >nul 2>nul

rem --- dependency self-check (Python / pywin32 / CAD / Excel) ------------
if "%DWG_BATCH_NODEPS%"=="1" goto AFTERDEPS
if not exist "%DEPS%" goto AFTERDEPS
%PY% "%DEPS%"
if errorlevel 1 goto DEPFAIL
goto AFTERDEPS
:DEPFAIL
echo.
echo  依赖不全，请按上面提示处理后重试。
if not "%DWG_BATCH_NOPAUSE%"=="1" pause
:AFTERDEPS

rem --- run arguments -----------------------------------------------------
set "DRYARG="
if "%DWG_BATCH_DRYRUN%"=="1" set "DRYARG=--dry-run"

echo.

rem --- run: %* expanded OUTSIDE any ( ... ) block ------------------------
rem progress prints live here; auto_export.py also tees to %LOG%.
if "%DWG_BATCH_NOCACHE%"=="1" goto RUNNOCACHE
%PY% "%SCRIPT%" %* --cache-dir "%CACHE%" --out-dir "%OUT%" %DRYARG% --log-file "%LOG%"
goto AFTERRUN
:RUNNOCACHE
%PY% "%SCRIPT%" %* --out-dir "%OUT%" %DRYARG% --log-file "%LOG%"
:AFTERRUN
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" goto DONEOK
echo  完成但有错误（%RC%）；日志：%LOG%
goto END
:DONEOK
echo  成品目录：%OUT%
echo  完成；日志：%LOG%
if not "%DRYARG%"=="" goto END
explorer "%OUT%" >nul 2>nul
:END
if not "%DWG_BATCH_NOPAUSE%"=="1" pause
exit /b %RC%

:NOARGS
echo.
echo  用法：把图纸文件夹或多个 .dwg 文件拖到本 bat 上。
echo.
if not "%DWG_BATCH_NOPAUSE%"=="1" pause
exit /b 1

:NOSCRIPT
echo  [错误] 找不到 %SCRIPT%
echo  请把 run_batch.bat、scripts 文件夹和 assets 文件夹放在同一目录。
if not "%DWG_BATCH_NOPAUSE%"=="1" pause
exit /b 1

:NOPY
echo.
echo  [错误] 未找到 Python 3，且 winget 也未能自动安装。
echo  请从 https://www.python.org/downloads/ 下载安装 Python 3.8 及以上版本，
echo  安装时务必勾选 "Add Python to PATH"，然后重新拖放运行本 bat。
echo.
if not "%DWG_BATCH_NOPAUSE%"=="1" pause
exit /b 1
