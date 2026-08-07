@echo off
REM Ejecutar UNA VEZ en cada PC de la oficina (como Administrador)
REM Deja el nombre facil: http://estudio:8501
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0configurar_acceso_facil.ps1\"'"
pause
