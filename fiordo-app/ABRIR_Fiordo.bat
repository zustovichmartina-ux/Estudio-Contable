@echo off
REM Abre la app de punto de equilibrio de Fiordo Cafe en Chrome.
REM Tambien queda accesible desde el celular conectado al mismo WiFi.
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Falta instalar Node.js. Descargalo de https://nodejs.org ^(version LTS^) y volve a abrir este archivo.
  pause
  exit /b 1
)

if not exist node_modules (
  echo Instalando por unica vez... puede tardar un par de minutos.
  call npm install
)

echo.
echo Para abrirla desde el celular ^(mismo WiFi^) usa alguna de estas direcciones:
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do echo    http:%%a:3000
echo.
start "" cmd /c "timeout /t 6 /nobreak >nul & start chrome http://localhost:3000 || start http://localhost:3000"
call npx next dev -H 0.0.0.0 -p 3000
