# Launch BD web page - NO Node.js, NO admin rights. Opens Microsoft Edge by default.
param(
    [int]$Port = 5173,
    [switch]$NoBrowser
)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$webRoot = Join-Path $root "web"
$pidFile = Join-Path $webRoot ".wght-server.pid"

$json = Join-Path $root "web\public\data\bd-sheet.json"
if (-not (Test-Path $json)) {
    Write-Host ""
    Write-Host "Fichier manquant: web\public\data\bd-sheet.json" -ForegroundColor Red
    Write-Host "Copiez ce fichier depuis un collegue, ou exportez avec Node:" -ForegroundColor Yellow
    Write-Host "  .\export-bd-data.ps1" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
$sizeMb = [math]::Round((Get-Item $json).Length / 1MB, 2)
Write-Host "Fichier donnees OK ($sizeMb Mo)" -ForegroundColor Green

$vendorVue = Join-Path $root "web\vendor\vue.esm-browser.js"
if (-not (Test-Path $vendorVue)) {
    Write-Host "Installation Vue + HyperFormula..." -ForegroundColor Yellow
    & (Join-Path $root "scripts\setup-web-vendor.ps1")
}

function Stop-PreviousWghtServer {
    if (-not (Test-Path $pidFile)) { return }
    $oldPid = 0
    [void][int]::TryParse((Get-Content $pidFile -Raw -ErrorAction SilentlyContinue), [ref]$oldPid)
    if ($oldPid -le 0) { return }
    if ($oldPid -eq $PID) { return }
    $proc = Get-Process -Id $oldPid -ErrorAction SilentlyContinue
    if (-not $proc) {
        Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
        return
    }
    $name = $proc.ProcessName
    if ($name -notin @('powershell', 'pwsh', 'cmd')) {
        Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
        return
    }
    Write-Host "Arret du serveur precedent (PID $oldPid)..." -ForegroundColor Yellow
    Stop-Process -Id $oldPid -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
}

function Stop-UserListenerOnPort([int]$p) {
    $conns = Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue
    foreach ($c in $conns) {
        $procId = $c.OwningProcess
        if ($procId -le 4) { continue }
        $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if (-not $proc) { continue }
        if ($proc.ProcessName -notin @('powershell', 'pwsh', 'node', 'cmd')) { continue }
        if ($procId -eq $PID) { continue }
        Write-Host "Liberation port $p ($($proc.ProcessName) PID $procId)..." -ForegroundColor Yellow
        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 400
    }
}

Stop-PreviousWghtServer
Stop-UserListenerOnPort $Port
Stop-UserListenerOnPort 5174

$openBrowser = -not $NoBrowser

function Open-InMicrosoftEdge([string]$url) {
    $edgePaths = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    )
    $edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($edge) {
        Start-Process -FilePath $edge -ArgumentList $url
    } else {
        Start-Process $url
    }
}

$nodeExe = Get-Command node -ErrorAction SilentlyContinue
$serverMjs = Join-Path $root "web\server.mjs"
if ($nodeExe -and (Test-Path $serverMjs)) {
    $env:PORT = "$Port"
    $url = "http://127.0.0.1:$Port/"
    Set-Content -Path $pidFile -Value $PID -Encoding ASCII -NoNewline
    Write-Host ""
    Write-Host "BD page: $url" -ForegroundColor Green
    Write-Host "API progressive (Node): meta + cellules par blocs" -ForegroundColor Cyan
    Write-Host "Ctrl+C pour arreter." -ForegroundColor Gray
    Write-Host ""
    if ($openBrowser) {
        Open-InMicrosoftEdge $url
    }
    Push-Location (Join-Path $root "web")
    try {
        & node server.mjs
    } finally {
        Pop-Location
        if (Test-Path $pidFile) { Remove-Item $pidFile -Force -ErrorAction SilentlyContinue }
    }
    exit $LASTEXITCODE
}

Write-Host "Node.js absent — serveur statique (sans API chunks)." -ForegroundColor Yellow
Write-Host "Installez Node ou lancez: node web/server.mjs" -ForegroundColor Yellow

& (Join-Path $root "scripts\serve-web.ps1") -Port $Port -WebRoot $webRoot -OpenBrowser:$openBrowser
