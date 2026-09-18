@echo off
title ARCA CONTROL — esta PC no abre AFIP
cd /d "%~dp0"
set "PY=%LocalAppData%\Programs\Python\Python314\python.exe"
if not exist "%PY%" set "PY=python"
echo.
echo  ========================================
echo   ARCA CONTROL
echo  ========================================
echo.
echo  Esta PC solo controla. AFIP lo abre el EJECUTOR.
echo  Abriendo la web...
echo.
start "" "https://estudiocontablemdp.streamlit.app/"
"%PY%" -u "%~dp0scripts\control_arca.py"
echo.
pause
