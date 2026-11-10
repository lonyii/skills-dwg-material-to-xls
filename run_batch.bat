@echo off
chcp 65001 >nul 2>nul
set "PYTHONIOENCODING=utf-8"
setlocal

rem =====================================================================
rem  dwg-material-to-xls  -  drag & drop batch runner (no AI needed)
rem  Drop ONE OR MORE folders (or .dwg files) onto this .bat file.
rem  Requirements on this PC: Python 3 + pywin32 + a CAD (GStarCAD /
rem  AutoCAD / ZWCAD) + Microsoft Excel. pywin32 is installed if missing.
rem
rem  NOTE: This script deliberately avoids ( ... ) command blocks around
rem  %* (the dropped arguments). A file name containing ")" would close
rem  such a block early and crash the batch ("... was unexpected at this
rem  time"). All branches use goto labels instead.
rem
rem  Optional environment variables:
rem    DWG_BATCH_NOPAUSE=1   do not wait for a key (scheduled / scripted runs)
rem    DWG_BATCH_DRYRUN=1    only show the plan, do not write .xls files
rem    DWG_BATCH_NOCACHE=1   always re-read drawings from CAD
rem =====================================================================

set "HERE=%~dp0"
set "SCRIPT=%HERE%scripts\auto_export.py"

if "%~1"=="" goto NOARGS
if not exist "%SCRIPT%" goto NOSCRIPT

rem --- locate Python 3 ---------------------------------------------------
set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY where python >nul 2>nul
if not errorlevel 1 if not defined PY set "PY=python"
if not defined PY goto NOPY

rem --- check / install pywin32 -------------------------------------------
%PY% -c "import win32com.client" >nul 2>nul
if errorlevel 1 goto INSTALLPY
goto AFTERPY
:INSTALLPY
echo  Installing missing dependency: pywin32 ...
%PY% -m pip install pywin32
:AFTERPY

rem --- entity cache: reused only while the DWG files are unchanged --------
set "DRYARG="
if "%DWG_BATCH_DRYRUN%"=="1" set "DRYARG=--dry-run"
if not exist "%USERPROFILE%\.workbuddy" mkdir "%USERPROFILE%\.workbuddy" >nul 2>nul
if not exist "%USERPROFILE%\.workbuddy\dwg-entity-cache" mkdir "%USERPROFILE%\.workbuddy\dwg-entity-cache" >nul 2>nul

set "LOG=%HERE%run_batch.log"
echo.
echo  Processing dropped items. CAD and Excel run in the background, please wait...
echo.
if exist "%HERE%version.json" findstr /c:"\"version\"" "%HERE%version.json"
echo  Using script: %SCRIPT%
echo.

rem --- run: %* is expanded OUTSIDE any ( ... ) block ----------------------
if "%DWG_BATCH_NOCACHE%"=="1" goto RUNNOCACHE
%PY% "%SCRIPT%" %* --cache-dir "%USERPROFILE%\.workbuddy\dwg-entity-cache" %DRYARG% > "%LOG%" 2>&1
goto AFTERRUN
:RUNNOCACHE
%PY% "%SCRIPT%" %* %DRYARG% > "%LOG%" 2>&1
:AFTERRUN
set "RC=%ERRORLEVEL%"
echo  ---- run log ----
type "%LOG%"
echo  -----------------

echo.
if "%RC%"=="0" goto DONEOK
echo  Finished with errors (code %RC%). Full log: "%LOG%"
goto END
:DONEOK
echo  Done. See the output folder printed above. Full log: "%LOG%"
:END
if not "%DWG_BATCH_NOPAUSE%"=="1" pause
exit /b %RC%

:NOARGS
echo.
echo  Usage: drag a folder ^(or several .dwg files^) onto this .bat file.
echo  The folder may contain drawings of several projects.
echo  Output folder is printed below when processing starts.
echo.
if not "%DWG_BATCH_NOPAUSE%"=="1" pause
exit /b 1

:NOSCRIPT
echo  [ERROR] Cannot find %SCRIPT%
if not "%DWG_BATCH_NOPAUSE%"=="1" pause
exit /b 1

:NOPY
echo  [ERROR] Python 3 not found in PATH. Install from https://www.python.org/
if not "%DWG_BATCH_NOPAUSE%"=="1" pause
exit /b 1
