# Download browser vendor bundles (Vue + HyperFormula) into web/vendor/
$ErrorActionPreference = "Stop"
$vendor = Join-Path $PSScriptRoot "..\web\vendor"
New-Item -ItemType Directory -Force -Path $vendor | Out-Null

$files = @{
    "vue.esm-browser.js" = "https://cdn.jsdelivr.net/npm/vue@3.5.13/dist/vue.esm-browser.js"
    "hyperformula.full.min.js" = "https://cdn.jsdelivr.net/npm/hyperformula@3.0.0/dist/hyperformula.full.min.js"
}

foreach ($entry in $files.GetEnumerator()) {
    $dest = Join-Path $vendor $entry.Key
    Write-Host "Downloading $($entry.Key) ..."
    Invoke-WebRequest -Uri $entry.Value -OutFile $dest
}

Write-Host "Vendor OK: $vendor" -ForegroundColor Green
