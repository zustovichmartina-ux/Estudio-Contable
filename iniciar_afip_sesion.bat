@echo off
title AFIP - Paso 1 login
cd /d "%~dp0"
call "%~dp0_arca_env.cmd"
echo.
echo  ========================================
echo    AFIP / ARCA — un solo paso ahora
echo  ========================================
echo.
echo  1. Se abre Chrome adelante (ARCA).
echo  2. Entras con clave + Guardar contrasena.
echo  3. Cuando veas el portal, ESTA ventana pide Enter.
echo  4. Apreta Enter recien cuando la clave este guardada.
echo.
echo  El worker queda en pausa mientras tanto.
echo  No hay tiempo limite.
echo.
pause
set AFIP_CHROME_HEADED=1
"%PY%" -u -m afip_worker.login_once
echo.
pause
