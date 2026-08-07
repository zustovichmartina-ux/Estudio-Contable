@echo off
REM Estudio Contable — arranque con tolerancia de red (WebSocket / inactividad)
REM Parametros para evitar perdida de sesion por inactividad prolongada:
REM   --server.disconnectedSessionTTL 3600   (conserva sesion 1 hora tras desconexion)
REM   --server.websocketPingInterval 10      (ping cada 10 s para mantener vivo el socket)
cd /d "%~dp0"
python -m streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.disconnectedSessionTTL 3600 --server.websocketPingInterval 10
pause
