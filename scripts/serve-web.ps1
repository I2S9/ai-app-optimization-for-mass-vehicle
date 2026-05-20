# Static file server — no Node.js, no admin rights required.
param(
    [int]$Port = 5173,
    [string]$WebRoot = (Join-Path $PSScriptRoot "..\web")
)

$WebRoot = (Resolve-Path $WebRoot).Path
$prefix = "http://127.0.0.1:$Port/"

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
} catch {
    Write-Host "Cannot bind to $prefix" -ForegroundColor Red
    Write-Host "Try another port: .\run-bd-server.ps1 -Port 8080" -ForegroundColor Yellow
    Write-Host $_.Exception.Message
    exit 1
}

Write-Host "BD page: $prefix" -ForegroundColor Green
Write-Host "Folder:  $WebRoot" -ForegroundColor Gray
Write-Host "Press Ctrl+C to stop." -ForegroundColor Gray

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
        if ($ext -eq '.json') {
            $response.Headers.Add('Cache-Control', 'no-cache')
        }
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        $response.Close()
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
