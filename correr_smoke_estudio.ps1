#Requires -Version 5.1
<#
.SYNOPSIS
  Corre el smoke funcional del Estudio Contable (NO reinicia ni toca Streamlit).
#>
$ErrorActionPreference = "Continue"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
$LogDir = Join-Path $Root "logs"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$Log = Join-Path $LogDir "smoke.log"

function Write-Log([string]$msg) {
    Add-Content -Path $Log -Value ("{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg) -Encoding UTF8
}

$python = $null
foreach ($cand in @("python.exe", "py.exe")) {
    $cmd = Get-Command $cand -ErrorAction SilentlyContinue
    if ($cmd) { $python = $cmd.Source; break }
}
if (-not $python) {
    Write-Log "ERROR: no hay python en PATH"
    Write-Output "FAIL: no hay python"
    exit 1
}

Write-Log "Smoke inicio"
& $python (Join-Path $Root "smoke_estudio.py")
$code = $LASTEXITCODE
$summary = Get-Content (Join-Path $LogDir "smoke_ultimo.txt") -TotalCount 1 -ErrorAction SilentlyContinue
Write-Log ("Smoke fin exit={0} {1}" -f $code, $summary)
Write-Output $summary
exit $code
