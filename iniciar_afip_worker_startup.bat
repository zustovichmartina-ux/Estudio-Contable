@echo off
REM Auto-start AFIP worker (Windows Startup). Background / hidden Chrome.
cd /d "%~dp0"
if not exist "jobs" mkdir "jobs"
set AFIP_CHROME_HEADED=0
echo [%date% %time%] AFIP worker auto-start >> "jobs\worker_autostart.log"
python -m afip_worker.main --live --serve --interval 3 >> "jobs\worker.log" 2>> "jobs\worker.log.err"
echo [%date% %time%] AFIP worker exited %ERRORLEVEL% >> "jobs\worker_autostart.log"
