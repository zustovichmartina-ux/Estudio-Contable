@echo off
REM Worker AFIP autonomo + API para Streamlit Cloud (tunel Cloudflare).
REM Dejar esta ventana abierta en RECEPCION.
cd /d "%~dp0"
echo [AFIP worker] API + cola. Ctrl+C para detener.
python -m afip_worker.main --dry-run --serve --interval 3
echo.
echo Si aparecio una URL trycloudflare, pegala en Streamlit Secrets (AFIP_WORKER_URL).
echo El token esta en jobs\.worker_token y el texto listo en jobs\cloud_bridge.txt
pause
