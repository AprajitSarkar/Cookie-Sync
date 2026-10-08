# Cookie Sync - Packaging Script for Chrome & Firefox (POSIX / AMO compliant)
$ErrorActionPreference = "Stop"

$baseDir = $PSScriptRoot
if (-not $baseDir) { $baseDir = Get-Location }

$pyScript = Join-Path $baseDir "package_amo.py"
$pythonExe = "C:\Users\Aprajit\AppData\Local\Python\pythoncore-3.14-64\python.exe"

Write-Host "Packaging Cookie Sync Extensions (POSIX / AMO Compliant)..." -ForegroundColor Cyan

if (Test-Path $pythonExe) {
    & $pythonExe "$pyScript"
} else {
    python "$pyScript"
}

Write-Host "All packages ready for deployment!" -ForegroundColor Green
