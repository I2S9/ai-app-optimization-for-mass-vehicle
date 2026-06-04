# Launch BD web page - requires Node.js for /api proxy to Supabase.
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

$vendorVue = Join-Path $root "web\vendor\vue.esm-browser.prod.js"
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

function Resolve-NodeExe {
    # 1) node dans le PATH (cas normal apres installation Node.js LTS)
    $cmd = Get-Command node -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    # 2) emplacements d'installation classiques + Node embarque dans Cursor
    $candidates = @(
        (Join-Path $env:ProgramFiles 'nodejs\node.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'nodejs\node.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\nodejs\node.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\cursor\resources\app\resources\helpers\node.exe')
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path $c)) { return $c }
    }
    # 3) installations nvm-windows
    $nvm = Join-Path $env:APPDATA 'nvm'
    if (Test-Path $nvm) {
        $found = Get-ChildItem -Path $nvm -Recurse -Filter 'node.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    return $null
}

$nodePath = Resolve-NodeExe
$serverMjs = Join-Path $root "web\server.mjs"
if (-not $nodePath -or -not (Test-Path $serverMjs)) {
    Write-Host ""
    Write-Host "ERREUR: Node.js introuvable (necessaire pour /api/v1/config)." -ForegroundColor Red
    Write-Host "Installez Node.js LTS (https://nodejs.org) puis relancez run-bd-server.bat" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
Write-Host "Node utilise : $nodePath" -ForegroundColor Green

$env:PORT = "$Port"
$env:WGHT_PROXY_API = "http://127.0.0.1:8000"
$url = "http://127.0.0.1:$Port/"
Set-Content -Path $pidFile -Value $PID -Encoding ASCII -NoNewline
Write-Host ""
Write-Host "Application : $url" -ForegroundColor Green
Write-Host "Laissez cette fenetre ouverte. Ctrl+C pour arreter." -ForegroundColor Gray
Write-Host ""
if ($openBrowser) {
    Open-InMicrosoftEdge $url
}
Push-Location (Join-Path $root "web")
try {
    & $nodePath server.mjs
} finally {
    Pop-Location
    if (Test-Path $pidFile) { Remove-Item $pidFile -Force -ErrorAction SilentlyContinue }
}
exit $LASTEXITCODE
