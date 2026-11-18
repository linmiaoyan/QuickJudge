@echo off
REM ASCII only on purpose: see start_service.bat
cd /d "%~dp0"
set "PY=C:\Users\Administrator\.workbuddy\binaries\python\envs\quickjudge\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
"%PY%" stop_service.py
echo.
"%PY%" start_service.py
echo.
pause
