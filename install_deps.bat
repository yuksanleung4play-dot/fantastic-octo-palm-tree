@echo off
setlocal EnableExtensions
pushd "%~dp0"
if errorlevel 1 (
  echo ERROR: cannot open the script folder.
  pause
  exit /b 1
)

echo Installing Python packages...
where py >nul 2>&1
if not errorlevel 1 (
  py -3 -m pip install -r requirements.txt
  py -3 -m pip install pywin32 pywinauto
  goto done
)
where python >nul 2>&1
if not errorlevel 1 (
  python -m pip install -r requirements.txt
  python -m pip install pywin32 pywinauto
  goto done
)
echo ERROR: Python not found.
pause
popd
exit /b 1

:done
echo.
echo Install finished. Next: double-click RUN_LME.bat
pause
popd
exit /b 0
