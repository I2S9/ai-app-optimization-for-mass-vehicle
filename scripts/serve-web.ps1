# Static file server (TcpListener) — no Node.js, no HTTP.sys URL reservation.
param(
    [int]$Port = 5173,
    [string]$WebRoot = (Join-Path $PSScriptRoot "..\web"),
    [switch]$OpenBrowser
)
$ErrorActionPreference = "Stop"
$WebRoot = (Resolve-Path $WebRoot).Path
$pidFile = Join-Path $WebRoot ".wght-server.pid"

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

function Send-HttpResponse($stream, [int]$code, [string]$contentType, [byte[]]$body) {
    $status = switch ($code) {
        200 { "OK" }
        404 { "Not Found" }
        default { "Error" }
    }
    $header = "HTTP/1.1 $code $status`r`n" +
        "Content-Type: $contentType`r`n" +
        "Content-Length: $($body.Length)`r`n" +
        "Connection: close`r`n"
    if ($contentType -match 'html|javascript|css|json') {
        $header += "Cache-Control: no-store, no-cache, must-revalidate`r`nPragma: no-cache`r`n"
    }
    $header += "`r`n"
    $bytes = [Text.Encoding]::ASCII.GetBytes($header)
    $stream.Write($bytes, 0, $bytes.Length)
    if ($body.Length -gt 0) {
        $stream.Write($body, 0, $body.Length)
    }
}

function Handle-Client($client) {
    try {
        $stream = $client.GetStream()
        $reader = New-Object System.IO.StreamReader($stream, [Text.Encoding]::ASCII, $false, 4096, $true)
        $requestLine = $reader.ReadLine()
        if (-not $requestLine) { return }
        do {
            $line = $reader.ReadLine()
        } while ($line -and $line.Length -gt 0)

        $parts = $requestLine -split ' '
        if ($parts.Length -lt 2 -or $parts[0] -ne 'GET') {
            Send-HttpResponse $stream 404 'text/plain; charset=utf-8' ([byte[]]@())
            return
        }
        $urlPath = [Uri]::UnescapeDataString(($parts[1] -split '\?')[0])
        if ($urlPath -eq '/') { $urlPath = '/index.html' }

        $apiJson = @{
            '/api/v1/config' = '{"mode":"static","chunkedLoad":false,"serverCalc":false,"cloudPersist":false,"remoteOnly":false,"version":2,"projectId":"default"}'
            '/api/v1/health' = '{"ok":true,"backend":"serve-web"}'
        }
        if ($apiJson.ContainsKey($urlPath)) {
            $body = [Text.Encoding]::UTF8.GetBytes($apiJson[$urlPath])
            Send-HttpResponse $stream 200 'application/json; charset=utf-8' $body
            return
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

        $relative = $urlPath.TrimStart('/').Replace('/', [IO.Path]::DirectorySeparatorChar)
        $filePath = [IO.Path]::GetFullPath((Join-Path $WebRoot $relative))
        if (-not $filePath.StartsWith($WebRoot, [StringComparison]::OrdinalIgnoreCase)) {
            Send-HttpResponse $stream 404 'text/plain; charset=utf-8' ([Text.Encoding]::UTF8.GetBytes('Forbidden'))
            return
        }
        if (-not (Test-Path $filePath -PathType Leaf)) {
            Send-HttpResponse $stream 404 'text/plain; charset=utf-8' ([Text.Encoding]::UTF8.GetBytes('Not found'))
            return
        }
        $ext = [IO.Path]::GetExtension($filePath).ToLowerInvariant()
        $contentType = $mime[$ext]
        if (-not $contentType) { $contentType = 'application/octet-stream' }
        $body = [IO.File]::ReadAllBytes($filePath)
        Send-HttpResponse $stream 200 $contentType $body
    } catch {
        # client disconnected
    } finally {
        try { $client.Close() } catch {}
    }
}

$fallbackPorts = @($Port, 8080, 5174, 5199, 5200) | Select-Object -Unique
$listener = $null
$boundPort = $null

foreach ($tryPort in $fallbackPorts) {
    try {
        $tryListener = [System.Net.Sockets.TcpListener]::new(
            [System.Net.IPAddress]::Loopback,
            $tryPort
        )
        $tryListener.Start()
        $listener = $tryListener
        $boundPort = $tryPort
        break
    } catch {
        if ($tryListener) {
            try { $tryListener.Stop() } catch {}
        }
    }
}

if (-not $listener) {
    Write-Host ""
    Write-Host "Impossible de demarrer le serveur (ports essayes: $($fallbackPorts -join ', '))." -ForegroundColor Red
    Write-Host "Fermez les autres fenetres WGHT (Ctrl+C) puis relancez run-bd-server.bat." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$url = "http://127.0.0.1:$boundPort/"
Set-Content -Path $pidFile -Value $PID -Encoding ASCII -NoNewline

Write-Host ""
Write-Host "BD page: $url" -ForegroundColor Green
Write-Host "Dossier: $WebRoot" -ForegroundColor Gray
Write-Host "Ctrl+C pour arreter." -ForegroundColor Gray
Write-Host ""

if ($OpenBrowser) {
    Open-InMicrosoftEdge $url
}

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()
        Handle-Client $client
    }
} finally {
    try { $listener.Stop() } catch {}
    if (Test-Path $pidFile) { Remove-Item $pidFile -Force -ErrorAction SilentlyContinue }
}
