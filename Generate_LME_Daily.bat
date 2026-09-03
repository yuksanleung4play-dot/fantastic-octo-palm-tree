@echo off
setlocal EnableExtensions EnableDelayedExpansion
pushd "%~dp0"
if errorlevel 1 (
  echo ERROR: cannot open the script folder.
  pause
  exit /b 1
)

echo ============================================
echo  LME daily report
echo  Folder: %CD%
echo ============================================
echo.

if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" generate_lme_daily.py %*
  set "RC=!ERRORLEVEL!"
  goto finish
)
if exist "venv\Scripts\python.exe" (
  "venv\Scripts\python.exe" generate_lme_daily.py %*
  set "RC=!ERRORLEVEL!"
  goto finish
)

where py >nul 2>&1
if not errorlevel 1 (
  py -3 generate_lme_daily.py %*
  set "RC=!ERRORLEVEL!"
  goto finish
)

where python >nul 2>&1
if not errorlevel 1 (
  python generate_lme_daily.py %*
  set "RC=!ERRORLEVEL!"
  goto finish
)

echo ERROR: Python not found.
echo Install Python 3.10+ from python.org
echo Tick "Add python.exe to PATH"
echo Then double-click install_deps.bat
pause
popd
exit /b 1

:finish
echo.
if not "!RC!"=="0" (
  echo FAILED. Read the messages above.
) else (
  echo DONE.
)
pause
popd
exit /b !RC!
