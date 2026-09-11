@echo off
title AFIP - Paso 1 login
cd /d "%~dp0"
echo.
echo  ========================================
echo    AFIP / ARCA — un solo paso ahora
echo  ========================================
echo.
echo  1. Se abre Chrome adelante (ARCA).
echo  2. Entras con clave + Guardar contrasena.
echo  3. Cuando veas el portal, esta ventana dice OK.
echo.
echo  El worker queda en pausa mientras tanto.
echo  No hay tiempo limite.
echo.
pause
set AFIP_CHROME_HEADED=1
set PYTHONUNBUFFERED=1
set "PY=%LocalAppData%\Programs\Python\Python314\python.exe"
if exist "%PY%" (
  "%PY%" -u -m afip_worker.login_once
) else (
  py -3 -u -m afip_worker.login_once
)
echo.
pause
