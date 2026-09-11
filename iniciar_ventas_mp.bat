@echo off
REM Ventas Mercado Pago de una sociedad (no mezclar con Mi Mercado Pago personal en 8505)
cd /d "%~dp0"
netstat -ano | findstr ":8506" | findstr "LISTENING" >nul
if %ERRORLEVEL%==0 (
  start "" "http://localhost:8506"
  exit /b 0
)
start "" "http://localhost:8506"
python -m streamlit run ventas_mp_app.py --server.port 8506 --server.headless true --server.disconnectedSessionTTL 3600 --server.websocketPingInterval 10
pause
