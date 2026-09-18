@echo off
title ARCA EJECUTOR — esta maquina abre AFIP
cd /d "%~dp0"
echo.
echo  ========================================
echo   ARCA EJECUTOR
echo  ========================================
echo.
echo  Esta PC SOLO EJECUTA: Chrome + AFIP + carpeta del cliente.
echo  El control es la web, desde la otra PC.
echo  No cierres esta ventana. No apagues esta maquina.
echo.
call "%~dp0iniciar_afip_worker.bat"
