#Requires -Version 5.1
<#
.SYNOPSIS
  Arranca el loop de smoke en segundo plano (no toca Streamlit).
#>
$ErrorActionPreference = "SilentlyContinue"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogDir = Join-Path $Root "logs"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$PidFile = Join-Path $LogDir "smoke_vigilancia.pid"
$Log = Join-Path $LogDir "smoke.log"

function Write-Log([string]$msg) {
    Add-Content -Path $Log -Value ("{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg) -Encoding UTF8
}

if (Test-Path $PidFile) {
    $old = Get-Content $PidFile -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($old -and (Get-Process -Id ([int]$old) -ErrorAction SilentlyContinue)) {
        Write-Output "Smoke-vigilante ya activo (PID $old)."
        exit 0
    }
}

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "powershell.exe"
$psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $Root 'vigilar_smoke_estudio.ps1')`""
$psi.WorkingDirectory = $Root
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true
try {
    $proc = [System.Diagnostics.Process]::Start($psi)
    if (-not $proc) { throw "Process.Start null" }
    Start-Sleep -Seconds 2
    $saved = if (Test-Path $PidFile) { Get-Content $PidFile | Select-Object -First 1 } else { $proc.Id }
    Write-Log "Smoke-vigilante lanzado launcher=$($proc.Id) vig=$saved"
    Write-Output "Smoke-vigilante activo (PID $saved). Cada 30 min. No reinicia la web."
} catch {
    Write-Log "ERROR al lanzar smoke-vigilante: $($_.Exception.Message)"
    Write-Output "FAIL: $($_.Exception.Message)"
    exit 1
}
