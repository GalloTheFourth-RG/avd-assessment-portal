#!/usr/bin/env pwsh
Write-Host "=== AVD Assessment Portal - Startup ==="
Write-Host "Working directory: $(Get-Location)"
Write-Host ""
Write-Host "=== /app contents ==="
Get-ChildItem /app -Recurse -Depth 2 | ForEach-Object { Write-Host $_.FullName }
Write-Host ""
Write-Host "=== Checking server script ==="
$serverPath = "/app/backend/src/server.ps1"
if (Test-Path $serverPath) {
    Write-Host "Found server.ps1 at $serverPath"
} else {
    Write-Host "ERROR: server.ps1 NOT FOUND at $serverPath"
    Write-Host "Searching..."
    Get-ChildItem /app -Recurse -Filter "server.ps1" | ForEach-Object { Write-Host "  Found: $($_.FullName)" }
}
Write-Host ""
Write-Host "=== Checking frontend dist ==="
$distPath = "/app/frontend/dist"
if (Test-Path $distPath) {
    Write-Host "Found frontend dist at $distPath"
    Get-ChildItem $distPath | ForEach-Object { Write-Host "  $($_.Name)" }
} else {
    Write-Host "WARNING: frontend dist NOT FOUND at $distPath"
}
Write-Host ""
Write-Host "=== Starting server ==="
try {
    & $serverPath
} catch {
    Write-Host "FATAL: $($_.Exception.Message)"
    Write-Host "Sleeping to keep container alive for log inspection..."
    Start-Sleep -Seconds 3600
}
