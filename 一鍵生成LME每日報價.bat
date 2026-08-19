@echo off
setlocal EnableExtensions
pushd "%~dp0"
call RUN_LME.bat %*
exit /b %ERRORLEVEL%
