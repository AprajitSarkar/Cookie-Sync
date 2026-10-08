# Cookie Sync - Packaging Script for Chrome & Firefox
$ErrorActionPreference = "Stop"

$baseDir = $PSScriptRoot
if (-not $baseDir) { $baseDir = Get-Location }

$chromeDir = Join-Path $baseDir "Chrome"
$firefoxDir = Join-Path $baseDir "Firefox"

$chromeZip = Join-Path $baseDir "Cookie-Sync-Chrome.zip"
$firefoxZip = Join-Path $baseDir "Cookie-Sync-Firefox.zip"

Write-Host "Packaging Cookie Sync Extensions..." -ForegroundColor Cyan

# 1. Package Chrome Extension
if (Test-Path $chromeZip) { Remove-Item $chromeZip -Force }
Write-Host "Creating Chrome Web Store package: $chromeZip" -ForegroundColor Yellow
Compress-Archive -Path "$chromeDir\*" -DestinationPath $chromeZip -Force
Write-Host "Chrome extension packaged successfully!" -ForegroundColor Green

# 2. Package Firefox Extension (AMO upload format)
if (Test-Path $firefoxZip) { Remove-Item $firefoxZip -Force }
Write-Host "Creating Firefox Add-ons (AMO) package: $firefoxZip" -ForegroundColor Yellow
Compress-Archive -Path "$firefoxDir\*" -DestinationPath $firefoxZip -Force
Write-Host "Firefox extension packaged successfully!" -ForegroundColor Green

Write-Host "All packages ready for deployment!" -ForegroundColor Cyan
