@echo off
chcp 65001 >nul
title 开放 QuickJudge 局域网端口 5001

REM ============================================================
REM  让同一局域网内的手机 / 其它电脑访问本机的 QuickJudge 服务
REM  用法：右键本文件 →「以管理员身份运行」，运行一次即可长期生效
REM ============================================================

net session >nul 2>&1
if errorlevel 1 (
  echo.
  echo  [需要管理员权限]
  echo  请右键本文件，选择「以管理员身份运行」。
  echo.
  pause
  exit /b 1
)

echo 正在添加防火墙入站规则：TCP 5001（范围：专用网络 + 域网络）...
netsh advfirewall firewall delete rule name="QuickJudge 5001" >nul 2>&1
netsh advfirewall firewall add rule name="QuickJudge 5001" dir=in action=allow protocol=TCP localport=5001 profile=private,domain description="QuickJudge 局域网访问" >nul
if errorlevel 1 (
  echo.
  echo  [失败] 规则没能添加，请确认是以管理员身份运行。
  echo.
  pause
  exit /b 1
)

echo  [完成] 已放行 TCP 5001。
echo.
echo  其它设备请访问（把 IP 换成本机当前地址）：
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do (
  for /f "tokens=1" %%b in ("%%a") do echo      http://%%b:5001/
)
echo.
echo  撤销放行：右键以管理员身份运行下面的命令
echo      netsh advfirewall firewall delete rule name="QuickJudge 5001"
echo.
pause
