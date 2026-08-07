@echo off
REM Una sola vez: registra inicio automatico + intenta abrir firewall
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0registrar_servicio_estudio.ps1"
echo.
pause
