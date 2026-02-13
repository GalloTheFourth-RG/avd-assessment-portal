#!/usr/bin/env pwsh
<#
    Assessment runner — called by the portal server as a subprocess.
    Handles its own Az login and runs the evidence pack script.
#>
param(
    [string]$ParamsFile,
    [string]$RunId,
    [string]$StorageAccount,
    [string]$StorageContainer,
    [string]$ClientId
)

$ErrorActionPreference = "Stop"

$outputDir = Join-Path ([System.IO.Path]::GetTempPath()) $RunId
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

try {
    Write-Host "[$RunId] Connecting to Azure..."
    if ($ClientId) {
        Connect-AzAccount -Identity -AccountId $ClientId -ErrorAction Stop | Out-Null
    } else {
        Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
    }
    Write-Host "[$RunId] Connected."

    # Read parameters from file
    $raw = Get-Content $ParamsFile -Raw | ConvertFrom-Json
    $params = @{}
    foreach ($prop in $raw.PSObject.Properties) {
        $val = $prop.Value
        # Convert JSON arrays back to PowerShell string arrays
        if ($val -is [System.Object[]]) { $val = [string[]]@($val) }
        $params[$prop.Name] = $val
    }
    # Clean up params file
    Remove-Item $ParamsFile -Force -ErrorAction SilentlyContinue

    Write-Host "[$RunId] Running assessment in $outputDir..."
    Push-Location $outputDir
    & /app/scripts/Get-Enhanced-AVD-EvidencePack.ps1 @params
    Pop-Location

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
