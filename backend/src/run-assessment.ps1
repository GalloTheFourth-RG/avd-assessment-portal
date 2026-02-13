#!/usr/bin/env pwsh
<#
    Assessment runner — loads saved Az profile and runs the evidence pack.
    Called as a background process by the portal server.
#>
param(
    [string]$ParamsFile,
    [string]$RunId,
    [string]$StorageAccount,
    [string]$StorageContainer,
    [string]$AzProfilePath
)

$ErrorActionPreference = "Stop"

$outputDir = Join-Path ([System.IO.Path]::GetTempPath()) $RunId
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

try {
    Write-Host "[$RunId] Loading Az profile from $AzProfilePath..."
    if (Test-Path $AzProfilePath) {
        Import-AzContext -Path $AzProfilePath -ErrorAction Stop | Out-Null
        Write-Host "[$RunId] ✓ Az context loaded"
    } else {
        throw "Az profile not found at $AzProfilePath"
    }
    
    $azCtx = Get-AzContext
    Write-Host "[$RunId] Tenant: $($azCtx.Tenant.Id), Account: $($azCtx.Account.Id)"

    # Read parameters from file
    $raw = Get-Content $ParamsFile -Raw | ConvertFrom-Json
    $params = @{}
    foreach ($prop in $raw.PSObject.Properties) {
        $val = $prop.Value
        if ($val -is [System.Object[]]) { $val = [string[]]@($val) }
        $params[$prop.Name] = $val
    }
    Remove-Item $ParamsFile -Force -ErrorAction SilentlyContinue

    Write-Host "[$RunId] Running assessment in $outputDir..."
    Push-Location $outputDir
    & /app/scripts/Get-Enhanced-AVD-EvidencePack.ps1 @params
    Pop-Location

    # Re-load the profile since the script may have trashed the context
    Write-Host "[$RunId] Re-loading Az profile for upload..."
    Import-AzContext -Path $AzProfilePath -ErrorAction Stop | Out-Null

    # Upload results
    Write-Host "[$RunId] Uploading results..."
    $ctx = New-AzStorageContext -StorageAccountName $StorageAccount -UseConnectedAccount
    $files = Get-ChildItem -Path $outputDir -Recurse -File
    foreach ($file in $files) {
        $blobName = "$RunId/$($file.Name)"
        Set-AzStorageBlobContent -File $file.FullName -Container $StorageContainer -Blob $blobName -Context $ctx -Force | Out-Null
    }

    Write-Host "[$RunId] Complete. $($files.Count) files uploaded."
    @{ status = "completed"; runId = $RunId; fileCount = $files.Count } | ConvertTo-Json | Set-Content "/tmp/$RunId-result.json"
}
catch {
    Write-Host "[$RunId] FAILED: $($_.Exception.Message)"
    @{ status = "failed"; runId = $RunId; error = $_.Exception.Message } | ConvertTo-Json | Set-Content "/tmp/$RunId-result.json"
}
finally {
    if (Test-Path $outputDir) { Remove-Item $outputDir -Recurse -Force -ErrorAction SilentlyContinue }
}
