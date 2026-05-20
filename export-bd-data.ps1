# One-time export BD sheet from Excel (only if workbook changed)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$node = & (Join-Path $root "scripts\get-node.ps1")
& $node (Join-Path $root "tools\export-bd-sheet.mjs")
