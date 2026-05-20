# Launch BD web page — NO Node.js, NO admin rights.
param([int]$Port = 5173)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$json = Join-Path $root "web\public\data\bd-sheet.json"

if (-not (Test-Path $json)) {
    Write-Host ""
    Write-Host "Missing: web\public\data\bd-sheet.json" -ForegroundColor Red
    Write-Host "Ask a colleague to copy this file, or run export on a PC with Node:" -ForegroundColor Yellow
    Write-Host "  .\export-bd-data.ps1  (needs Node / Cursor only)" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$sizeMb = [math]::Round((Get-Item $json).Length / 1MB, 2)
Write-Host "Data file OK ($sizeMb MB)" -ForegroundColor Green

& (Join-Path $root "scripts\serve-web.ps1") -Port $Port -WebRoot (Join-Path $root "web")
