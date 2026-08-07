#Requires -Version 5.1
<#
.SYNOPSIS
  Registra inicio automatico al iniciar sesion + regla de firewall (LAN).
  Ejecutar una vez en esta PC (RECEPCION).
#>
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Vbs = Join-Path $Root "iniciar_estudio_oculto.vbs"
$TaskName = "EstudioContable_Streamlit_Oculto"

if (-not (Test-Path $Vbs)) {
    throw "No se encuentra $Vbs"
}

# --- Tarea programada al iniciar sesion (usuario actual, oculta) ---
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$Vbs`"" -WorkingDirectory $Root
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Estudio Contable (Streamlit) en segundo plano, sin ventana. Puerto 8501." `
    -Force | Out-Null

Write-Output "OK tarea programada: $TaskName (al iniciar sesion de $env:USERNAME)"

# --- Firewall: permitir 8501 en red privada (oficina) ---
$fwName = "Estudio Contable Streamlit 8501"
try {
    $fw = Get-NetFirewallRule -DisplayName $fwName -ErrorAction SilentlyContinue
    if ($fw) { Remove-NetFirewallRule -DisplayName $fwName -ErrorAction SilentlyContinue }
    New-NetFirewallRule `
        -DisplayName $fwName `
        -Direction Inbound `
        -Protocol TCP `
        -LocalPort 8501 `
        -Action Allow `
        -Profile Private, Domain `
        -Description "Acceso LAN a Estudio Contable (Streamlit)" | Out-Null
    Write-Output "OK firewall: puerto 8501 abierto (Private/Domain)"
} catch {
    Write-Output "AVISO firewall (hace falta admin): $($_.Exception.Message)"
    Write-Output "Si otras PCs no entran, abri PowerShell como Administrador y volve a correr este script."
}

$hostName = $env:COMPUTERNAME
$ips = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
    Where-Object { $_.IPAddress -notlike "127.*" -and $_.PrefixOrigin -ne "WellKnown" } |
    Select-Object -ExpandProperty IPAddress)

Write-Output ""
Write-Output "Desde esta PC:     http://127.0.0.1:8501"
Write-Output "Desde la oficina:  http://${hostName}:8501"
foreach ($ip in $ips) {
    Write-Output "                   http://${ip}:8501"
}
Write-Output ""
Write-Output "Para arrancar ahora (oculto):  wscript.exe `"$Vbs`""
Write-Output "Para detener:  powershell -File `"$Root\detener_estudio_oculto.ps1`""
