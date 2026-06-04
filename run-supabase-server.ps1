# Lance l'API Python (FastAPI) adossee a Supabase (HTTPS port 443) et sert
# l'application web sur le MEME port (8000). Toute cellule modifiee est ecrite
# dans Supabase en temps reel et rechargee au redemarrage.
param(
    [int]$Port = 8000,
    [switch]$NoBrowser
)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$apiDir = Join-Path $root "api"
$python = Join-Path $apiDir ".venv\Scripts\python.exe"
$envFile = Join-Path $apiDir ".env"

if (-not (Test-Path $python)) {
    Write-Host ""
    Write-Host "ERREUR: environnement Python introuvable: $python" -ForegroundColor Red
    Write-Host "Creez-le avec:" -ForegroundColor Yellow
    Write-Host "  cd api; python -m venv .venv; .\.venv\Scripts\Activate.ps1; pip install -r requirements.txt" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
if (-not (Test-Path $envFile)) {
    Write-Host ""
    Write-Host "ERREUR: api\.env manquant (SUPABASE_URL + SUPABASE_SERVICE_KEY)." -ForegroundColor Red
    Write-Host "Copiez api\.env.example vers api\.env et renseignez les cles Supabase." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

function Stop-ListenerOnPort([int]$p) {
    $conns = Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue
    foreach ($c in $conns) {
        $procId = $c.OwningProcess
        if ($procId -le 4) { continue }
        $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if (-not $proc) { continue }
        if ($proc.ProcessName -notin @('python', 'pythonw', 'uvicorn')) { continue }
        Write-Host "Liberation port $p ($($proc.ProcessName) PID $procId)..." -ForegroundColor Yellow
        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 400
    }
}

Stop-ListenerOnPort $Port

$url = "http://127.0.0.1:$Port/"

function Open-BrowserWhenReady([string]$u, [int]$p) {
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Milliseconds 500
        try {
            $r = Invoke-WebRequest -UseBasicParsing -Uri "$u`api/v1/health" -TimeoutSec 2
            if ($r.StatusCode -eq 200) { break }
        } catch { }
    }
    $edgePaths = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    )
    $edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($edge) { Start-Process -FilePath $edge -ArgumentList $u } else { Start-Process $u }
}

if (-not $NoBrowser) {
    Start-Job -ScriptBlock ${function:Open-BrowserWhenReady} -ArgumentList $url, $Port | Out-Null
}

Write-Host ""
Write-Host "API + application : $url" -ForegroundColor Green
Write-Host "Backend           : Supabase (HTTPS 443) - persistance temps reel" -ForegroundColor Green
Write-Host "Laissez cette fenetre ouverte. Ctrl+C pour arreter." -ForegroundColor Gray
Write-Host ""

Push-Location $apiDir
try {
    $env:PYTHONPATH = $apiDir
    & $python -m uvicorn app.main:app --port $Port
} finally {
    Pop-Location
}
exit $LASTEXITCODE
