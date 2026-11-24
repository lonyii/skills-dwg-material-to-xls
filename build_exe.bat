@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Building material balance EXE, please wait...
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 build_exe.py %*
) else (
  python build_exe.py %*
)
if errorlevel 1 (
  echo Build failed. Check the log above.
) else (
  echo Build done. EXE is in the dist folder.
)
pause
