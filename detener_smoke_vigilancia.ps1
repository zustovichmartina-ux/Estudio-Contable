#Requires -Version 5.1
<#
.SYNOPSIS
  Detiene el loop de smoke (no toca Streamlit ni el vigilante de vida).
#>
$ErrorActionPreference = "SilentlyContinue"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$PidFile = Join-Path $Root "logs\smoke_vigilancia.pid"
$Log = Join-Path $Root "logs\smoke.log"

function Write-Log([string]$msg) {
    if (Test-Path (Split-Path $Log)) {
        Add-Content -Path $Log -Value ("{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg) -Encoding UTF8
    }
}

if (Test-Path $PidFile) {
    $pidSaved = Get-Content $PidFile -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($pidSaved) {
        Stop-Process -Id ([int]$pidSaved) -Force -ErrorAction SilentlyContinue
        Write-Log "Detenido smoke-vigilante PID $pidSaved"
    }
    Remove-Item $PidFile -Force -ErrorAction SilentlyContinue
}
Write-Output "Smoke-vigilancia detenida."
