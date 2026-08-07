#Requires -Version 5.1
<#
.SYNOPSIS
  Detiene el Streamlit oculto del Estudio Contable (puerto 8501).
#>
$ErrorActionPreference = "SilentlyContinue"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$PidFile = Join-Path $Root "logs\streamlit.pid"
$Log = Join-Path $Root "logs\servicio.log"

function Write-Log([string]$msg) {
    if (Test-Path (Split-Path $Log)) {
        Add-Content -Path $Log -Value ("{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg) -Encoding UTF8
    }
}

if (Test-Path $PidFile) {
    $pidSaved = Get-Content $PidFile -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($pidSaved) {
        Stop-Process -Id ([int]$pidSaved) -Force -ErrorAction SilentlyContinue
        Write-Log "Detenido PID $pidSaved (pidfile)"
    }
    Remove-Item $PidFile -Force -ErrorAction SilentlyContinue
}

Get-NetTCPConnection -LocalPort 8501 -State Listen -ErrorAction SilentlyContinue | ForEach-Object {
    Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue
    Write-Log "Detenido PID $($_.OwningProcess) (puerto 8501)"
}

Write-Output "Estudio Contable detenido."
