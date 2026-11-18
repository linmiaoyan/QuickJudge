@echo off
REM ============================================================
REM Foreground launcher: a console window stays open, and closing
REM that window stops the service. Use the background launcher in
REM this folder instead if you want no window at all.
REM
REM ASCII only on purpose. cmd decodes .bat files with the system
REM code page (936 here), so UTF-8 Chinese in comments can swallow
REM the line break and make the parser run the comment as a
REM command. That is exactly what used to break this file: the
REM "set PY" line got eaten and cmd fell through to a bare python.
REM All Chinese output comes from app.py, which prints through the
REM wide-char console API and never garbles.
REM ============================================================
cd /d "%~dp0"
title QuickJudge - foreground

set "PY=C:\Users\Administrator\.workbuddy\binaries\python\envs\quickjudge\Scripts\python.exe"
if not exist "%PY%" (
  where python >nul 2>&1
  if errorlevel 1 (
    echo.
    echo [ERROR] Python not found. Install it, or edit the PY path in this file.
    echo.
    pause
    exit /b 1
  )
  set "PY=python"
)

"%PY%" app.py

echo.
echo Service stopped.
pause
