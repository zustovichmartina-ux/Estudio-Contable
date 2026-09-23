@echo off
title ARCA EJECUTOR — instalar dependencias
cd /d "%~dp0"
call "%~dp0_arca_env.cmd"
echo.
echo  ========================================
echo   Instalar Python deps + Playwright
echo  ========================================
echo.
echo  Python: %PY%
"%PY%" -V
if errorlevel 1 (
  echo No encuentro Python. Instala Python 3.12 y volve a correr esto.
  pause
  exit /b 1
)
echo.
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install -r "%~dp0scripts\requirements_arca_ejecutor.txt"
if errorlevel 1 (
  echo Fallo pip. Revisa la red / permisos.
  pause
  exit /b 1
)
"%PY%" -m playwright install chrome
if errorlevel 1 (
  echo playwright install chrome fallo; pruebo chromium...
  "%PY%" -m playwright install chromium
)
echo.
if not exist "%~dp0jobs" mkdir "%~dp0jobs"
if not exist "%~dp0jobs\pending" mkdir "%~dp0jobs\pending"
if not exist "%~dp0jobs\running" mkdir "%~dp0jobs\running"
if not exist "%~dp0jobs\done" mkdir "%~dp0jobs\done"
if not exist "%~dp0jobs\error" mkdir "%~dp0jobs\error"
if not exist "%~dp0jobs\needs_auth" mkdir "%~dp0jobs\needs_auth"
echo.
echo  Listo. Siguiente:
echo   1) iniciar_afip_sesion.bat  (guardar 1 clave del estudio)
echo   2) ejecutor_arca.bat        (dejar abierto)
echo.
echo  cloudflared: si no esta en PATH, ponelo en tools\cloudflared.exe
echo.
pause
