# Copia el paquete ejecutor ARCA a una carpeta standalone (Desktop o red).
# No incluye jobs/.worker_token ni el perfil Chrome (se copian a mano si hace falta).
param(
    [string]$Dest = ""
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
if (-not $Dest) {
    $Dest = Join-Path $Root "ARCA_ejecutor_pack"
}

Write-Host "Origen: $Root"
Write-Host "Destino: $Dest"

New-Item -ItemType Directory -Force -Path $Dest | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $Dest "scripts") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $Dest "tools") | Out-Null
foreach ($d in @("pending", "running", "done", "error", "needs_auth", "archive")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Dest "jobs\$d") | Out-Null
}

$files = @(
    "_arca_env.cmd",
    "instalar_todo.bat",
    "iniciar_afip_sesion.bat",
    "iniciar_afip_worker.bat",
    "iniciar_afip_worker_startup.bat",
    "ejecutor_arca.bat",
    "control_arca.bat"
)
foreach ($f in $files) {
    Copy-Item -Force (Join-Path $Root $f) (Join-Path $Dest $f)
}

Copy-Item -Force (Join-Path $Root "ARCA_ejecutor\LEEME.txt") (Join-Path $Dest "LEEME.txt")
Copy-Item -Force (Join-Path $Root "scripts\requirements_arca_ejecutor.txt") (Join-Path $Dest "scripts\requirements_arca_ejecutor.txt")
Copy-Item -Force (Join-Path $Root "scripts\control_arca.py") (Join-Path $Dest "scripts\control_arca.py")

$afipSrc = Join-Path $Root "afip_worker"
$afipDst = Join-Path $Dest "afip_worker"
if (Test-Path $afipDst) { Remove-Item -Recurse -Force $afipDst }
Copy-Item -Recurse -Force $afipSrc $afipDst
Get-ChildItem -Recurse -Directory -Filter "__pycache__" $afipDst | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

$tokenSrc = Join-Path $Root "jobs\.worker_token"
$tokenDst = Join-Path $Dest "jobs\.worker_token"
if (Test-Path $tokenSrc) {
    Copy-Item -Force $tokenSrc $tokenDst
    Write-Host "Copié jobs\.worker_token (mismo token que Secrets)."
} else {
    Write-Host "AVISO: no hay jobs\.worker_token en el origen. Copialo a mano al destino."
}

Write-Host ""
Write-Host "Listo. En la PC ociosa abrí: $Dest"
Write-Host "  1) instalar_todo.bat"
Write-Host "  2) iniciar_afip_sesion.bat"
Write-Host "  3) ejecutor_arca.bat"
