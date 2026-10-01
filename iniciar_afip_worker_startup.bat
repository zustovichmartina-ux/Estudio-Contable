@echo off
REM Auto-start AFIP worker (Windows Startup). Se relanza si se cae.
cd /d "%~dp0"
call "%~dp0_arca_env.cmd"
if not exist "jobs" mkdir "jobs"
set AFIP_CHROME_HEADED=0
echo [%date% %time%] AFIP worker auto-start >> "jobs\worker_autostart.log"
:loop
"%PY%" -u -m afip_worker.main --live --serve --interval 3 >> "jobs\worker.log" 2>> "jobs\worker.log.err"
echo [%date% %time%] AFIP worker exited %ERRORLEVEL% — reintento 10s >> "jobs\worker_autostart.log"
timeout /t 10 /nobreak >nul
goto loop
