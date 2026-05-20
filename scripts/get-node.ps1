# Outputs full path to node.exe (stdout). Exits 1 if not found.
$nodeCmd = Get-Command node -ErrorAction SilentlyContinue
if ($nodeCmd) {
    Write-Output $nodeCmd.Source
    exit 0
}
$cursorNode = Join-Path $env:LOCALAPPDATA "Programs\cursor\resources\app\resources\helpers\node.exe"
if (Test-Path $cursorNode) {
    Write-Output $cursorNode
    exit 0
}
$programFilesNode = "C:\Program Files\nodejs\node.exe"
if (Test-Path $programFilesNode) {
    Write-Output $programFilesNode
    exit 0
}
Write-Error "Node.js not found. Install LTS from https://nodejs.org/ or run from Cursor terminal."
exit 1
