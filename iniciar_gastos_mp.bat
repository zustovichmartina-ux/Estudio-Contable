@echo off
REM Mi Mercado Pago — libro personal permanente (no mezclar con el Estudio / sociedades)
cd /d "%~dp0"
netstat -ano | findstr ":8505" | findstr "LISTENING" >nul
if %ERRORLEVEL%==0 (
  start "" "http://localhost:8505"
  exit /b 0
)
start "" "http://localhost:8505"
python -m streamlit run gastos_mp_app.py --server.port 8505 --server.headless true --server.disconnectedSessionTTL 3600 --server.websocketPingInterval 10
pause
