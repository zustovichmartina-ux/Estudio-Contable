#Requires -Version 5.1
<#
.SYNOPSIS
  Loop liviano: cada N minutos corre smoke_estudio.py.
  NO reinicia Streamlit. NO usa el navegador. Solo escribe logs/smoke_*.
#>
$ErrorActionPreference = "SilentlyContinue"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogDir = Join-Path $Root "logs"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$PidFile = Join-Path $LogDir "smoke_vigilancia.pid"
$Log = Join-Path $LogDir "smoke.log"
$IntervalMin = 30

function Write-Log([string]$msg) {
    $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $Log -Value $line -Encoding UTF8
}

if (Test-Path $PidFile) {
    $old = Get-Content $PidFile -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($old) {
        $procOld = Get-Process -Id ([int]$old) -ErrorAction SilentlyContinue
        if ($procOld -and $procOld.Id -ne $PID) {
            Write-Output "Ya hay smoke-vigilante (PID $old). Saliendo."
            exit 0
        }
    }
}
$PID | Set-Content $PidFile -Encoding ASCII
Write-Log ("Smoke-vigilante iniciado PID={0} cada {1}m" -f $PID, $IntervalMin)
Write-Output ("Smoke cada {0} min (PID {1}) - no toca la web" -f $IntervalMin, $PID)

while ($true) {
    try {
        & (Join-Path $Root "correr_smoke_estudio.ps1") | Out-Null
    } catch {
        $err = $_.Exception.Message
        Write-Log ("ERROR smoke: {0}" -f $err)
    }
    Start-Sleep -Seconds ($IntervalMin * 60)
}
