@echo off
REM ASCII only on purpose: cmd decodes .bat with the system code page,
REM so non-ASCII comments here can be mis-parsed. All Chinese messages
REM are printed by the Python scripts instead.
cd /d "%~dp0"
set "PY=C:\Users\Administrator\.workbuddy\binaries\python\envs\quickjudge\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
"%PY%" start_service.py
echo.
pause
