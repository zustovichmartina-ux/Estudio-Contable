@echo off
REM Worker AFIP autonomo + API para Streamlit Cloud (tunel Cloudflare).
REM Dejar esta ventana abierta en RECEPCION.
cd /d "%~dp0"
echo [AFIP worker] API + cola. Ctrl+C para detener.
python -m afip_worker.main --live --serve --interval 3
echo.
echo Si hay jobs\cloudflared\config.yml (tunel con nombre), la URL es estable: Secrets una sola vez.
echo Si usas tunel rapido (trycloudflare), la URL cambia: actualiza AFIP_WORKER_URL en Secrets.
echo El token esta en jobs\.worker_token y el texto listo en jobs\cloud_bridge.txt
pause
