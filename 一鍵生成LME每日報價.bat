@echo off
cd /d "%~dp0"
call "%~dp0Generate_LME_Daily.bat" %*
exit /b %ERRORLEVEL%
