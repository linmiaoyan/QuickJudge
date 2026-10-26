@echo off
REM Scanner soft-reset launcher. Double-click this file when the scanner
REM reports offline / busy and you do not want to unplug the USB cable.
setlocal
set LOG=%TEMP%\qj_scanner_reset.log
if exist "%LOG%" del "%LOG%" >nul 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0reset_scanner.ps1"
echo Resetting scanner, please wait...
timeout /t 25 /nobreak >nul
chcp 65001 >nul
type "%LOG%" 2>nul
echo.
echo ------------------------------------------------------------
echo If the scanner still does not work, unplug the USB cable,
echo wait 10 seconds, then plug it back into a REAR USB port.
echo ------------------------------------------------------------
pause
