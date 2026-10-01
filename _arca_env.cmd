@echo off
REM Resuelve Python para el ejecutor ARCA. Preferir 3.12 (paquete de la PC ociosa).
set "PY="
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
  set "PY=%LocalAppData%\Programs\Python\Python312\python.exe"
  goto :done
)
if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
  set "PY=%LocalAppData%\Programs\Python\Python313\python.exe"
  goto :done
)
if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
  set "PY=%LocalAppData%\Programs\Python\Python311\python.exe"
  goto :done
)
where py >nul 2>&1
if not errorlevel 1 (
  for /f "delims=" %%I in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PY=%%I"
)
if not defined PY set "PY=python"
:done
set PYTHONUNBUFFERED=1
