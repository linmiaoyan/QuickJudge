@echo off
chcp 65001 >nul
cd /d "%~dp0"
title QuickJudge 服务

REM 优先用项目自带的虚拟环境，找不到再退回系统 python
set "PY=C:\Users\Administrator\.workbuddy\binaries\python\envs\quickjudge\Scripts\python.exe"
if not exist "%PY%" (
  where python >nul 2>&1
  if errorlevel 1 (
    echo.
    echo [错误] 没有找到 Python，请先安装或在脚本里改成你的 python 路径。
    echo.
    pause
    exit /b 1
  )
  set "PY=python"
)

echo ============================================
echo    QuickJudge 服务启动中
echo    监听 0.0.0.0:5001，同一局域网内的设备可访问
echo    关闭此窗口即停止服务
echo ============================================
echo.

"%PY%" app.py

echo.
echo 服务已停止。
pause
