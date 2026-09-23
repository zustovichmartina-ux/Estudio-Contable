@echo off
title AFIP worker - EJECUTOR ARCA
REM Worker autonomo: cola + API + Chrome oculto. Se relanza si se cae.
cd /d "%~dp0"
call "%~dp0_arca_env.cmd"
set AFIP_CHROME_HEADED=0
echo [AFIP worker] Cola autonoma. Cerra esta ventana para detener.
echo Si la web no conecta, mira jobs\cloud_bridge.txt y actualiza Secrets.
echo Python: %PY%
echo.
:loop
"%PY%" -u -m afip_worker.main --live --serve --interval 3
echo.
echo El worker se detuvo. Reintento en 8 segundos.
timeout /t 8
goto loop
