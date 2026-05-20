# Launch BD web page - NO Node.js, NO admin rights. Opens Microsoft Edge by default.
param(
    [int]$Port = 5173,
    [switch]$NoBrowser
)
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

$vendorVue = Join-Path $root "web\vendor\vue.esm-browser.js"
if (-not (Test-Path $vendorVue)) {
    Write-Host "Installing web vendor (Vue + HyperFormula)..." -ForegroundColor Yellow
    & (Join-Path $root "scripts\setup-web-vendor.ps1")
}

function Stop-ListenerOnPort([int]$p) {
    $conns = Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue
    foreach ($c in $conns) {
        $proc = Get-Process -Id $c.OwningProcess -ErrorAction SilentlyContinue
        if ($proc) {
            Write-Host "Stopping $($proc.ProcessName) (PID $($proc.Id)) on port $p..." -ForegroundColor Yellow
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 400
        }
    }
}

function Open-InMicrosoftEdge([string]$url) {
    $edgePaths = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    )
    $edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($edge) {
        Start-Process -FilePath $edge -ArgumentList $url
        Write-Host "Opened Microsoft Edge: $url" -ForegroundColor Cyan
    } else {
        Start-Process $url
        Write-Host "Edge not found - opened default browser: $url" -ForegroundColor Yellow
    }
}

# Always restart so Edge gets fresh HTML/JS (avoids blank page from cache)
Stop-ListenerOnPort $Port
Stop-ListenerOnPort 5174

$openBrowser = -not $NoBrowser
& (Join-Path $root "scripts\serve-web.ps1") -Port $Port -WebRoot (Join-Path $root "web") -OpenBrowser:$openBrowser
