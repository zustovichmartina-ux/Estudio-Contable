#Requires -Version 5.1
<#
.SYNOPSIS
  Loop de vigilancia: si la web del Estudio Contable (:8501) no responde, la reinicia sola.
#>
$ErrorActionPreference = "SilentlyContinue"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogDir = Join-Path $Root "logs"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$Log = Join-Path $LogDir "vigilancia.log"
$PidFile = Join-Path $LogDir "vigilancia.pid"
$IntervalSec = 45
$FailThreshold = 2
$Url = "http://127.0.0.1:8501/"

function Write-Log([string]$msg) {
    Add-Content -Path $Log -Value ("{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg) -Encoding UTF8
}

# Evitar vigilantes duplicados
if (Test-Path $PidFile) {
    $old = Get-Content $PidFile -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($old) {
        $procOld = Get-Process -Id ([int]$old) -ErrorAction SilentlyContinue
        if ($procOld -and $procOld.Id -ne $PID) {
            Write-Output "Ya hay un vigilante (PID $old). Saliendo."
            exit 0
        }
    }
}
$PID | Set-Content -Path $PidFile -Encoding ASCII
Write-Log "Vigilante iniciado PID=$PID interval=${IntervalSec}s"
Write-Output "Vigilando $Url cada ${IntervalSec}s (PID $PID)"

$fails = 0
while ($true) {
    $ok = $false
    try {
        $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 10
        if ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 400) {
            $ok = $true
        }
    } catch {
        $ok = $false
    }

    if ($ok) {
        if ($fails -gt 0) { Write-Log "Recuperado OK tras $fails fallo(s)" }
        $fails = 0
    } else {
        $fails++
        Write-Log "Fallo health-check ($fails/$FailThreshold)"
        if ($fails -ge $FailThreshold) {
            Write-Log "Reiniciando Estudio Contable..."
            try {
                & (Join-Path $Root "detener_estudio_oculto.ps1") | Out-Null
            } catch {}
            Start-Sleep -Seconds 3
            try {
                & (Join-Path $Root "iniciar_estudio_oculto.ps1") | Out-Null
                Write-Log "Reinicio disparado"
            } catch {
                Write-Log ("ERROR al reiniciar: " + $_.Exception.Message)
            }
            $fails = 0
            Start-Sleep -Seconds 15
        }
    }

    Start-Sleep -Seconds $IntervalSec
}
