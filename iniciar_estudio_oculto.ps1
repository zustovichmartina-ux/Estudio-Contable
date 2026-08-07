#Requires -Version 5.1
<#
.SYNOPSIS
  Arranca Streamlit del Estudio Contable oculto (sin consola ni barra de tareas).
#>
$ErrorActionPreference = "SilentlyContinue"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$LogDir = Join-Path $Root "logs"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$Log = Join-Path $LogDir "servicio.log"
$PidFile = Join-Path $LogDir "streamlit.pid"

function Write-Log([string]$msg) {
    Add-Content -Path $Log -Value ("{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg) -Encoding UTF8
}

# Evitar duplicados: liberar puerto 8501
$listeners = Get-NetTCPConnection -LocalPort 8501 -State Listen -ErrorAction SilentlyContinue
foreach ($c in $listeners) {
    if ($c.OwningProcess) {
        Write-Log "Deteniendo proceso previo PID $($c.OwningProcess) en :8501"
        Stop-Process -Id $c.OwningProcess -Force -ErrorAction SilentlyContinue
    }
}
Start-Sleep -Seconds 2

$python = $null
foreach ($cand in @("pythonw.exe", "python.exe")) {
    $cmd = Get-Command $cand -ErrorAction SilentlyContinue
    if ($cmd) { $python = $cmd.Source; break }
}
if (-not $python) {
    Write-Log "ERROR: no se encontro python/pythonw en PATH"
    exit 1
}

# Argumentos como string unica (ProcessStartInfo.Arguments)
$arguments = @(
    "-m streamlit run app.py",
    "--server.port 8501",
    "--server.address 0.0.0.0",
    "--server.headless true",
    "--browser.gatherUsageStats false",
    "--server.disconnectedSessionTTL 3600",
    "--server.websocketPingInterval 10"
) -join " "

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $python
$psi.Arguments = $arguments
$psi.WorkingDirectory = $Root
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true
# No redirigir stdout/stderr: si se redirige y no se lee, Streamlit puede colgarse.

try {
    $proc = [System.Diagnostics.Process]::Start($psi)
    if (-not $proc) { throw "Process.Start devolvio null" }
    $proc.Id | Set-Content -Path $PidFile -Encoding ASCII
    Write-Log "Iniciado PID=$($proc.Id) python=$python bind=0.0.0.0:8501"
} catch {
    Write-Log "ERROR al iniciar: $($_.Exception.Message)"
    exit 1
}

# Arrancar vigilante en segundo plano (reinicia si :8501 cae)
$vigPidFile = Join-Path $LogDir "vigilancia.pid"
$vigRunning = $false
if (Test-Path $vigPidFile) {
    $vigPid = Get-Content $vigPidFile -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($vigPid -and (Get-Process -Id ([int]$vigPid) -ErrorAction SilentlyContinue)) {
        $vigRunning = $true
    }
}
if (-not $vigRunning) {
    $vigPsi = New-Object System.Diagnostics.ProcessStartInfo
    $vigPsi.FileName = "powershell.exe"
    $vigPsi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $Root 'vigilar_estudio_oculto.ps1')`""
    $vigPsi.WorkingDirectory = $Root
    $vigPsi.UseShellExecute = $false
    $vigPsi.CreateNoWindow = $true
    try {
        $vigProc = [System.Diagnostics.Process]::Start($vigPsi)
        if ($vigProc) { Write-Log "Vigilante lanzado PID=$($vigProc.Id)" }
    } catch {
        Write-Log "No se pudo lanzar vigilante: $($_.Exception.Message)"
    }
}
