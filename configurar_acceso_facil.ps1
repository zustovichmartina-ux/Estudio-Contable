#Requires -Version 5.1
<#
.SYNOPSIS
  Configura el nombre facil http://estudio:8501 en esta PC (hosts)
  y crea accesos directos "Estudio Contable" en Escritorio y Menu Inicio.
#>
$ErrorActionPreference = "Stop"
$HostNameFacil = "estudio"
$Puerto = 8501
$IpServidor = "192.168.1.8"  # PC RECEPCION
$UrlFacil = "http://${HostNameFacil}:${Puerto}"
$UrlRed = "http://RECEPCION:${Puerto}"
$UrlLocal = "http://127.0.0.1:${Puerto}"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Escritorio = [Environment]::GetFolderPath("Desktop")
$Inicio = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"

function Add-HostsEntry {
    param([string]$Ip, [string]$Name)
    $hosts = "$env:SystemRoot\System32\drivers\etc\hosts"
    $contenido = @(Get-Content $hosts -ErrorAction SilentlyContinue)
    $ya = $contenido | Where-Object {
        $_ -match ("\b{0}\b" -f [regex]::Escape($Name)) -and $_ -notmatch "^\s*#"
    }
    if ($ya) {
        Write-Output "Hosts: ya existe entrada para '$Name'"
        return $true
    }
    Add-Content -Path $hosts -Value "`r`n# Estudio Contable`r`n$Ip`t$Name" -Encoding ascii
    Write-Output "Hosts: agregado $Ip  $Name"
    return $true
}

function New-UrlShortcut([string]$Path, [string]$TargetUrl) {
    @"
[InternetShortcut]
URL=$TargetUrl
"@ | Set-Content -Path $Path -Encoding ASCII
}

$esAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).
    IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$hostsOk = $false
if ($esAdmin) {
    try {
        $hostsOk = Add-HostsEntry -Ip $IpServidor -Name $HostNameFacil
    } catch {
        Write-Output "AVISO hosts: $($_.Exception.Message)"
    }
} else {
    Write-Output "AVISO: sin admin no se puede crear http://estudio:8501 (hosts)."
    Write-Output "       Ejecuta CREAR_acceso_Estudio_Contable.bat y acepta UAC."
}

# Acceso principal: nombre facil si hosts OK, si no RECEPCION (siempre anda en la red)
$urlPrincipal = if ($hostsOk) { $UrlFacil } else { $UrlRed }

$nombreAcceso = "Estudio Contable.url"
New-UrlShortcut (Join-Path $Escritorio $nombreAcceso) $urlPrincipal
New-UrlShortcut (Join-Path $Inicio $nombreAcceso) $urlPrincipal
New-UrlShortcut (Join-Path $Root $nombreAcceso) $urlPrincipal
New-UrlShortcut (Join-Path $Escritorio "Estudio Contable (esta PC).url") $UrlLocal

Write-Output ""
Write-Output "Listo."
Write-Output "  Menu Inicio / Escritorio:  Estudio Contable"
Write-Output "  URL principal:             $urlPrincipal"
if ($hostsOk) {
    Write-Output "  Tambien podes escribir:    $UrlFacil"
}
Write-Output "  Por nombre de PC:          $UrlRed"
Write-Output ""
Write-Output "En otras PCs: copia CREAR_acceso_Estudio_Contable.bat y ejecutalos como Administrador una vez."
