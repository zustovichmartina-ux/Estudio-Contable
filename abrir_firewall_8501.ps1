#Requires -Version 5.1
# Abrir puerto 8501 (correr como Administrador)
$ErrorActionPreference = "Stop"
$fwName = "Estudio Contable Streamlit 8501"
Get-NetFirewallRule -DisplayName $fwName -ErrorAction SilentlyContinue | Remove-NetFirewallRule -ErrorAction SilentlyContinue
New-NetFirewallRule `
    -DisplayName $fwName `
    -Direction Inbound `
    -Protocol TCP `
    -LocalPort 8501 `
    -Action Allow `
    -Profile Private, Domain `
    -Description "Acceso LAN a Estudio Contable (Streamlit)" | Out-Null
Write-Output "OK: puerto 8501 abierto para redes Private/Domain."
Write-Output "Proba desde otra PC: http://$($env:COMPUTERNAME):8501"
Start-Sleep -Seconds 4
