@echo off
REM Ejecutar como Administrador: abre el puerto 8501 para la red de la oficina
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0abrir_firewall_8501.ps1\"'"
