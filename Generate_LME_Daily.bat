@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
set "LME_ONECLICK=1"

echo ============================================================
echo   一鍵生成 LME每日報價yyyymmdd.xlsx
echo   工作目錄：%CD%
echo ============================================================
echo.

set "PYTHON="
if exist "%~dp0.venv\Scripts\python.exe" set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not defined PYTHON if exist "%~dp0venv\Scripts\python.exe" set "PYTHON=%~dp0venv\Scripts\python.exe"

if not defined PYTHON (
  where py >nul 2>&1
  if not errorlevel 1 (
    py -3 "%~dp0generate_lme_daily.py" %*
    goto :after
  )
  where python >nul 2>&1
  if not errorlevel 1 (
    python "%~dp0generate_lme_daily.py" %*
    goto :after
  )
  echo 找不到 Python。請先安裝 Python 3.10+，或在本目錄建立 .venv。
  echo 例如：
  echo   py -3 -m venv .venv
  echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
  echo   .venv\Scripts\python.exe -m pip install pywin32 pywinauto
  goto :fail
)

"%PYTHON%" "%~dp0generate_lme_daily.py" %*

:after
if errorlevel 1 goto :fail
echo.
echo 完成。
echo.
if not defined LME_NO_PAUSE pause
exit /b 0

:fail
echo.
echo 產生失敗，請看上方錯誤（常見：config 路徑、Excel 未開、Bloomberg 未登入、巨集名稱不對）。
echo.
if not defined LME_NO_PAUSE pause
exit /b 1
