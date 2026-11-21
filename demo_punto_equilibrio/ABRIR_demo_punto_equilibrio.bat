@echo off
REM Abre el tablero de punto de equilibrio en Chrome (http://localhost:8510)
cd /d "%~dp0"
start "Servidor demo punto de equilibrio" /min python -m http.server 8510
timeout /t 2 /nobreak >nul
start chrome "http://localhost:8510/punto_equilibrio_cafe.html" || start "" "http://localhost:8510/punto_equilibrio_cafe.html"
echo Tablero abierto en http://localhost:8510/punto_equilibrio_cafe.html
echo Para apagarlo, cerra la ventana "Servidor demo punto de equilibrio".
