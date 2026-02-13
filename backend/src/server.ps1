<#
.SYNOPSIS
    AVD Assessment Portal — Backend HTTP Server
    Lightweight PowerShell HTTP listener that serves the frontend and handles assessment API calls.
#>

param(
    [int]$Port = 3000
)

$ErrorActionPreference = "Stop"

# ============================================================================
# Configuration
# ============================================================================
$script:StorageAccount = $env:STORAGE_ACCOUNT_NAME
$script:StorageContainer = $env:STORAGE_CONTAINER ?? "results"
$script:ClientId = $env:AZURE_CLIENT_ID
$script:ScriptPath = "/app/scripts/Get-Enhanced-AVD-EvidencePack.ps1"
$script:FrontendPath = "/app/frontend/dist"
$script:ActiveJobs = @{}
$script:AzConnected = $false

# ============================================================================
# Connect to Azure on startup using Managed Identity
# ============================================================================
Write-Host "  Connecting to Azure..." -ForegroundColor Gray
Write-Host "  AZURE_CLIENT_ID: $(if ($script:ClientId) { $script:ClientId } else { 'NOT SET' })" -ForegroundColor Gray
try {
    if ($script:ClientId) {
        Connect-AzAccount -Identity -AccountId $script:ClientId -ErrorAction Stop | Out-Null
        Write-Host "  ✓ Connected with user-assigned managed identity" -ForegroundColor Green
    } else {
        # Try system-assigned identity as fallback
        Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
        Write-Host "  ✓ Connected with system-assigned managed identity" -ForegroundColor Green
    }
    $script:AzConnected = $true
    $ctx = Get-AzContext
    Write-Host "  ✓ Tenant: $($ctx.Tenant.Id)" -ForegroundColor Green
} catch {
    Write-Host "  ⚠ Azure login failed: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "  The portal will start but Azure API calls will fail." -ForegroundColor Yellow
    Write-Host "  Check that AZURE_CLIENT_ID is set and the managed identity is assigned to this Container App." -ForegroundColor Yellow
}

# ============================================================================
# MIME types for static file serving
# ============================================================================
$mimeTypes = @{
    ".html" = "text/html"
    ".css"  = "text/css"
    ".js"   = "application/javascript"
    ".json" = "application/json"
    ".png"  = "image/png"
    ".svg"  = "image/svg+xml"
    ".ico"  = "image/x-icon"
    ".woff2" = "font/woff2"
    ".woff" = "font/woff"
}

# ============================================================================
# Helper: JSON response
# ============================================================================
function Send-JsonResponse {
    param($Response, $Data, [int]$StatusCode = 200)
    $json = $Data | ConvertTo-Json -Depth 10 -Compress
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = "application/json; charset=utf-8"
    $Response.ContentLength64 = $buffer.Length
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $Response.OutputStream.Close()
}

# ============================================================================
# Helper: Serve static file
# ============================================================================
function Send-StaticFile {
    param($Response, [string]$FilePath)
    if (Test-Path $FilePath) {
        $ext = [System.IO.Path]::GetExtension($FilePath)
        $mime = $mimeTypes[$ext] ?? "application/octet-stream"
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        $Response.StatusCode = 200
        $Response.ContentType = $mime
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
        # SPA fallback — serve index.html for client-side routing
        $indexPath = Join-Path $script:FrontendPath "index.html"
        if (Test-Path $indexPath) {
            $bytes = [System.IO.File]::ReadAllBytes($indexPath)
            $Response.StatusCode = 200
            $Response.ContentType = "text/html"
            $Response.ContentLength64 = $bytes.Length
            $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $Response.StatusCode = 404
            $buffer = [System.Text.Encoding]::UTF8.GetBytes("Not found")
            $Response.ContentLength64 = $buffer.Length
            $Response.OutputStream.Write($buffer, 0, $buffer.Length)
        }
    }
    $Response.OutputStream.Close()
}

# ============================================================================
# API: GET /api/health
# ============================================================================
function Handle-Health {
    param($Response)
    $data = @{
        status = "healthy"
        timestamp = (Get-Date -Format "o")
        scriptExists = (Test-Path $script:ScriptPath)
        storageConfigured = (-not [string]::IsNullOrEmpty($script:StorageAccount))
        identityConfigured = (-not [string]::IsNullOrEmpty($script:ClientId))
        azureConnected = $script:AzConnected
    }
    Send-JsonResponse -Response $Response -Data $data
}

# ============================================================================
# API: GET /api/subscriptions — List accessible subscriptions
# ============================================================================
function Handle-ListSubscriptions {
    param($Response)
    try {
        if (-not $script:AzConnected) { throw "Not connected to Azure. Check managed identity configuration." }
        $subs = Get-AzSubscription -ErrorAction Stop | Select-Object @{N='id';E={$_.SubscriptionId}}, @{N='name';E={$_.Name}}, @{N='state';E={$_.State}}
        Send-JsonResponse -Response $Response -Data @{ subscriptions = @($subs) }
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{ error = $_.Exception.Message; subscriptions = @() } -StatusCode 500
    }
}

# ============================================================================
# API: POST /api/assess — Start an assessment run
# ============================================================================
function Handle-StartAssessment {
    param($Response, $Body)
    
    $config = $Body | ConvertFrom-Json
    $runId = "run-$(Get-Date -Format 'yyyyMMdd-HHmmss')-$([guid]::NewGuid().ToString().Substring(0,4))"
    
    # Build parameter set
    $params = @{
        TenantId = $config.tenantId
        SubscriptionIds = @($config.subscriptionIds)
        GenerateHtmlReport = $true
        SkipDisclaimer = $true
        CreateZip = $true
    }
    
    if ($config.logAnalyticsWorkspaceIds -and $config.logAnalyticsWorkspaceIds.Count -gt 0) {
        $params.LogAnalyticsWorkspaceResourceIds = @($config.logAnalyticsWorkspaceIds)
    }
    if ($config.metricsLookbackDays) { $params.MetricsLookbackDays = [int]$config.metricsLookbackDays }
    if ($config.includeAdvisor) { $params.IncludeAzureAdvisor = $true }
    if ($config.includeReservations) { $params.IncludeReservationAnalysis = $true }
    if ($config.skipCosts) { $params.SkipActualCosts = $true }
    if ($config.scrubPII) { $params.ScrubPII = $true }
    if ($config.quickSummary) { $params.QuickSummary = $true }
    if ($config.companyName) { $params.CompanyName = $config.companyName }
    if ($config.analystName) { $params.AnalystName = $config.analystName }
    
    # Start as background job
    $jobScript = {
        param($ScriptPath, $Params, $RunId, $StorageAccount, $StorageContainer, $ClientId)
        
        $ErrorActionPreference = "Stop"
        $outputDir = Join-Path ([System.IO.Path]::GetTempPath()) $RunId
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        
        try {
            # Connect with managed identity
            Connect-AzAccount -Identity -AccountId $ClientId -ErrorAction Stop | Out-Null
            
            # Set output location
            Push-Location $outputDir
            
            # Run the assessment
            & $ScriptPath @Params
            
            Pop-Location
            
            # Upload results to blob storage
            $ctx = New-AzStorageContext -StorageAccountName $StorageAccount -UseConnectedAccount
            $files = Get-ChildItem -Path $outputDir -Recurse -File
            foreach ($file in $files) {
                $blobName = "$RunId/$($file.Name)"
                Set-AzStorageBlobContent -File $file.FullName -Container $StorageContainer -Blob $blobName -Context $ctx -Force | Out-Null
            }
            
            @{ status = "completed"; runId = $RunId; fileCount = $files.Count }
        }
        catch {
            @{ status = "failed"; runId = $RunId; error = $_.Exception.Message }
        }
        finally {
            if (Test-Path $outputDir) { Remove-Item $outputDir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
    
    $job = Start-Job -ScriptBlock $jobScript -ArgumentList $script:ScriptPath, $params, $runId, $script:StorageAccount, $script:StorageContainer, $script:ClientId
    $script:ActiveJobs[$runId] = @{ Job = $job; StartTime = Get-Date; Config = $config }
    
    Send-JsonResponse -Response $Response -Data @{
        runId = $runId
        status = "started"
        message = "Assessment started. Poll /api/assess/$runId for status."
    }
}

# ============================================================================
# API: GET /api/assess/{runId} — Check assessment status
# ============================================================================
function Handle-AssessmentStatus {
    param($Response, [string]$RunId)
    
    if (-not $script:ActiveJobs.ContainsKey($RunId)) {
        Send-JsonResponse -Response $Response -Data @{ error = "Run not found"; runId = $RunId } -StatusCode 404
        return
    }
    
    $jobInfo = $script:ActiveJobs[$RunId]
    $job = $jobInfo.Job
    $elapsed = [math]::Round(((Get-Date) - $jobInfo.StartTime).TotalSeconds, 0)
    
    if ($job.State -eq "Running") {
        Send-JsonResponse -Response $Response -Data @{
            runId = $RunId
            status = "running"
            elapsedSeconds = $elapsed
        }
    }
    elseif ($job.State -eq "Completed") {
        $result = Receive-Job -Job $job
        Remove-Job -Job $job -ErrorAction SilentlyContinue
        $script:ActiveJobs.Remove($RunId)
        
        Send-JsonResponse -Response $Response -Data @{
            runId = $RunId
            status = $result.status ?? "completed"
            elapsedSeconds = $elapsed
            fileCount = $result.fileCount
            error = $result.error
        }
    }
    else {
        $errorMsg = if ($job.ChildJobs.Count -gt 0) { $job.ChildJobs[0].Error | Out-String } else { "Unknown error" }
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        $script:ActiveJobs.Remove($RunId)
        
        Send-JsonResponse -Response $Response -Data @{
            runId = $RunId
            status = "failed"
            elapsedSeconds = $elapsed
            error = $errorMsg
        }
    }
}

# ============================================================================
# API: GET /api/results/{runId} — List result files
# ============================================================================
function Handle-ListResults {
    param($Response, [string]$RunId)
    try {
        $ctx = New-AzStorageContext -StorageAccountName $script:StorageAccount -UseConnectedAccount
        $blobs = Get-AzStorageBlob -Container $script:StorageContainer -Prefix "$RunId/" -Context $ctx
        $prefix = "^$RunId/"
        $files = $blobs | ForEach-Object {
            $cleanName = $_.Name -replace $prefix, ''
            @{
                name = $cleanName
                size = $_.Length
                lastModified = $_.LastModified.ToString("o")
                url = "/api/results/$RunId/$cleanName"
            }
        }
        Send-JsonResponse -Response $Response -Data @{ runId = $RunId; files = @($files) }
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{ error = $_.Exception.Message } -StatusCode 500
    }
}

# ============================================================================
# API: GET /api/results/{runId}/{filename} — Download a result file
# ============================================================================
function Handle-DownloadResult {
    param($Response, [string]$RunId, [string]$FileName)
    try {
        $ctx = New-AzStorageContext -StorageAccountName $script:StorageAccount -UseConnectedAccount
        $blobName = "$RunId/$FileName"
        $tempFile = Join-Path ([System.IO.Path]::GetTempPath()) $FileName
        
        Get-AzStorageBlobContent -Container $script:StorageContainer -Blob $blobName -Destination $tempFile -Context $ctx -Force | Out-Null
        
        $bytes = [System.IO.File]::ReadAllBytes($tempFile)
        $ext = [System.IO.Path]::GetExtension($FileName)
        $mime = $mimeTypes[$ext] ?? "application/octet-stream"
        
        $Response.StatusCode = 200
        $Response.ContentType = $mime
        $Response.ContentLength64 = $bytes.Length
        $Response.AddHeader("Content-Disposition", "inline; filename=`"$FileName`"")
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $Response.OutputStream.Close()
        
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{ error = $_.Exception.Message } -StatusCode 500
    }
}

# ============================================================================
# API: GET /api/runs — List past assessment runs
# ============================================================================
function Handle-ListRuns {
    param($Response)
    try {
        $ctx = New-AzStorageContext -StorageAccountName $script:StorageAccount -UseConnectedAccount
        $blobs = Get-AzStorageBlob -Container $script:StorageContainer -Context $ctx
        $runs = @{}
        foreach ($blob in $blobs) {
            $parts = $blob.Name -split "/", 2
            if ($parts.Count -eq 2) {
                $rid = $parts[0]
                if (-not $runs.ContainsKey($rid)) {
                    $runs[$rid] = @{ runId = $rid; files = 0; lastModified = $blob.LastModified.ToString("o"); totalSize = 0 }
                }
                $runs[$rid].files++
                $runs[$rid].totalSize += $blob.Length
            }
        }
        $runList = $runs.Values | Sort-Object { $_.lastModified } -Descending
        Send-JsonResponse -Response $Response -Data @{ runs = @($runList) }
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{ error = $_.Exception.Message; runs = @() } -StatusCode 500
    }
}

# ============================================================================
# Main HTTP Listener
# ============================================================================
Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  AVD Assessment Portal — Server v1.0.0                      ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host "  Port:     $Port" -ForegroundColor Gray
Write-Host "  Script:   $(if (Test-Path $script:ScriptPath) { '✓ Found' } else { '✗ Missing!' })" -ForegroundColor $(if (Test-Path $script:ScriptPath) { 'Green' } else { 'Red' })
Write-Host "  Storage:  $(if ($script:StorageAccount) { $script:StorageAccount } else { '⚠ Not configured' })" -ForegroundColor $(if ($script:StorageAccount) { 'Green' } else { 'Yellow' })
Write-Host "  Identity: $(if ($script:ClientId) { $script:ClientId.Substring(0,8) + '...' } else { '⚠ Not configured' })" -ForegroundColor $(if ($script:ClientId) { 'Green' } else { 'Yellow' })
Write-Host ""

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://+:$Port/")
$listener.Start()
Write-Host "  🚀 Listening on http://0.0.0.0:$Port`n" -ForegroundColor Green

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        $path = $request.Url.AbsolutePath
        $method = $request.HttpMethod
        
        # Add CORS headers
        $response.AddHeader("Access-Control-Allow-Origin", "*")
        $response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        $response.AddHeader("Access-Control-Allow-Headers", "Content-Type")
        
        if ($method -eq "OPTIONS") {
            $response.StatusCode = 204
            $response.OutputStream.Close()
            continue
        }
        
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host "  [$timestamp] $method $path" -ForegroundColor Gray
        
        try {
            # Route: API endpoints
            if ($path -eq "/api/health" -and $method -eq "GET") {
                Handle-Health -Response $response
            }
            elseif ($path -eq "/api/subscriptions" -and $method -eq "GET") {
                Handle-ListSubscriptions -Response $response
            }
            elseif ($path -eq "/api/assess" -and $method -eq "POST") {
                $body = [System.IO.StreamReader]::new($request.InputStream).ReadToEnd()
                Handle-StartAssessment -Response $response -Body $body
            }
            elseif ($path -match "^/api/assess/([^/]+)$" -and $method -eq "GET") {
                Handle-AssessmentStatus -Response $response -RunId $Matches[1]
            }
            elseif ($path -match "^/api/results/([^/]+)$" -and $method -eq "GET") {
                Handle-ListResults -Response $response -RunId $Matches[1]
            }
            elseif ($path -match "^/api/results/([^/]+)/(.+)$" -and $method -eq "GET") {
                Handle-DownloadResult -Response $response -RunId $Matches[1] -FileName $Matches[2]
            }
            elseif ($path -eq "/api/runs" -and $method -eq "GET") {
                Handle-ListRuns -Response $response
            }
            # Route: Static files (frontend)
            else {
                $filePath = if ($path -eq "/") { Join-Path $script:FrontendPath "index.html" }
                            else { Join-Path $script:FrontendPath ($path.TrimStart("/")) }
                Send-StaticFile -Response $response -FilePath $filePath
            }
        }
        catch {
            Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
            Send-JsonResponse -Response $response -Data @{ error = $_.Exception.Message } -StatusCode 500
        }
    }
}
finally {
    $listener.Stop()
    Write-Host "`n  Server stopped." -ForegroundColor Yellow
}
