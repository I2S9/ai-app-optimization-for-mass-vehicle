# Static file server - no Node.js, no admin rights required.
param(
    [int]$Port = 5173,
    [string]$WebRoot = (Join-Path $PSScriptRoot "..\web"),
    [switch]$OpenBrowser
)
$WebRoot = (Resolve-Path $WebRoot).Path

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

$fallbackPorts = @($Port, 8080, 5174) | Select-Object -Unique
$listener = $null
$boundPort = $null

foreach ($tryPort in $fallbackPorts) {
    $prefix = "http://127.0.0.1:$tryPort/"
    $tryListener = New-Object System.Net.HttpListener
    $tryListener.Prefixes.Add($prefix)
    try {
        $tryListener.Start()
        $listener = $tryListener
        $boundPort = $tryPort
        break
    } catch {
        $tryListener.Close()
    }
}

if (-not $listener) {
    Write-Host "Cannot start server (ports tried: $($fallbackPorts -join ', '))." -ForegroundColor Red
    Write-Host "Close other WGHT server windows (Ctrl+C) or reboot, then run run-bd-server.bat again." -ForegroundColor Yellow
    exit 1
}

$url = "http://127.0.0.1:$boundPort/"
Write-Host "BD page: $url" -ForegroundColor Green
Write-Host "Folder:  $WebRoot" -ForegroundColor Gray
Write-Host "Press Ctrl+C to stop." -ForegroundColor Gray
Write-Host "Tip: always use Edge at this address, not index.html from the folder." -ForegroundColor Gray

if ($OpenBrowser) {
    Open-InMicrosoftEdge $url
}

$mime = @{
    '.html' = 'text/html; charset=utf-8'
    '.js'   = 'text/javascript; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.ico'  = 'image/x-icon'
    '.png'  = 'image/png'
    '.svg'  = 'image/svg+xml'
}
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        $urlPath = $request.Url.LocalPath
        if ($urlPath -eq '/') { $urlPath = '/index.html' }
        $relative = $urlPath.TrimStart('/').Replace('/', [IO.Path]::DirectorySeparatorChar)
        $filePath = Join-Path $WebRoot $relative
        if (-not (Test-Path $filePath -PathType Leaf)) {
            $response.StatusCode = 404
            $bytes = [Text.Encoding]::UTF8.GetBytes('Not found')
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
            $response.Close()
            continue
        }
        $ext = [IO.Path]::GetExtension($filePath).ToLowerInvariant()
        $contentType = $mime[$ext]
        if (-not $contentType) { $contentType = 'application/octet-stream' }
        $bytes = [IO.File]::ReadAllBytes($filePath)
        $response.StatusCode = 200
        $response.ContentType = $contentType
        $response.ContentLength64 = $bytes.Length
        if ($ext -in '.html', '.js', '.css', '.json') {
            $response.Headers.Add('Cache-Control', 'no-store, no-cache, must-revalidate')
            $response.Headers.Add('Pragma', 'no-cache')
        }
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        $response.Close()
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
