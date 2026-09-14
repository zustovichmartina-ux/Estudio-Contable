@echo off
REM Worker AFIP autonomo + API (Chrome oculto). Solo iniciar_afip_sesion.bat usa Chrome visible.
cd /d "%~dp0"
set AFIP_CHROME_HEADED=0
echo [AFIP worker] API + cola en segundo plano. Ctrl+C para detener.
python -m afip_worker.main --live --serve --interval 3
echo.
echo Si aparecio una URL trycloudflare, pegala en Streamlit Secrets (AFIP_WORKER_URL).
echo El token esta en jobs\.worker_token y el texto listo en jobs\cloud_bridge.txt
pause
