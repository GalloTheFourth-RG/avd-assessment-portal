#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Enable Entra ID authentication (Easy Auth) on the AVD Assessment Portal.

.DESCRIPTION
    This creates an Entra ID App Registration and configures Azure Container Apps
    built-in authentication so only authenticated users can access the portal.

    Run this AFTER the portal is deployed and working.

.EXAMPLE
    ./Setup-Auth.ps1 -ResourceGroup "rg-avd-assessment" -ContainerAppName "avdassess-portal"

.NOTES
    Requires: Az CLI with logged-in session that has permissions to create App Registrations.
#>

param(
    [string]$ResourceGroup = "rg-avd-assessment",
    [string]$ContainerAppName = "avdassess-portal"
)

$ErrorActionPreference = "Stop"

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  AVD Assessment Portal — Entra ID Authentication Setup      ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

# Get the portal URL
$fqdn = (az containerapp show -n $ContainerAppName -g $ResourceGroup --query "properties.configuration.ingress.fqdn" -o tsv)
if (-not $fqdn) { throw "Container App not found" }

$portalUrl = "https://$fqdn"
Write-Host "  Portal URL: $portalUrl" -ForegroundColor Gray

# Step 1: Create App Registration
Write-Host "`n── Step 1: Creating Entra ID App Registration ──" -ForegroundColor Yellow

$appName = "AVD Assessment Portal"
$redirectUri = "$portalUrl/.auth/login/aad/callback"

# Check if it already exists
$existing = az ad app list --display-name $appName --query "[0].appId" -o tsv 2>$null
if ($existing) {
    Write-Host "  App Registration already exists: $existing" -ForegroundColor Green
    $appId = $existing
} else {
    $appJson = az ad app create `
        --display-name $appName `
        --web-redirect-uris $redirectUri `
        --sign-in-audience "AzureADMyOrg" `
        --query "{appId:appId, id:id}" -o json | ConvertFrom-Json
    
    $appId = $appJson.appId
    Write-Host "  ✓ Created App Registration: $appId" -ForegroundColor Green
}

# Step 2: Create client secret
Write-Host "`n── Step 2: Creating Client Secret ──" -ForegroundColor Yellow
$secretJson = az ad app credential reset --id $appId --display-name "portal-auth" --query "{password:password}" -o json | ConvertFrom-Json
$clientSecret = $secretJson.password
Write-Host "  ✓ Client secret created" -ForegroundColor Green

# Step 3: Get Tenant ID
$tenantId = (az account show --query "tenantId" -o tsv)

# Step 4: Enable Easy Auth on Container App
Write-Host "`n── Step 3: Enabling Easy Auth on Container App ──" -ForegroundColor Yellow

az containerapp auth microsoft update `
    -n $ContainerAppName `
    -g $ResourceGroup `
    --client-id $appId `
    --client-secret $clientSecret `
    --tenant-id $tenantId `
    --issuer "https://sts.windows.net/$tenantId/v2.0" `
    --yes 2>$null

# Enable auth and set to require authentication
az containerapp auth update `
    -n $ContainerAppName `
    -g $ResourceGroup `
    --unauthenticated-client-action "RedirectToLoginPage" `
    --enabled true 2>$null

# Set REQUIRE_AUTH env var so the backend validates headers
az containerapp update `
    -n $ContainerAppName `
    -g $ResourceGroup `
    --set-env-vars "REQUIRE_AUTH=true" 2>$null

Write-Host "  ✓ Easy Auth enabled" -ForegroundColor Green

# Summary
Write-Host "`n══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  ✅ Authentication configured!" -ForegroundColor Green
Write-Host ""
Write-Host "  App Registration:  $appId" -ForegroundColor Gray
Write-Host "  Tenant:            $tenantId" -ForegroundColor Gray
Write-Host "  Portal:            $portalUrl" -ForegroundColor Gray
Write-Host ""
Write-Host "  Users in your tenant will be prompted to sign in with" -ForegroundColor White
Write-Host "  their Entra ID credentials before accessing the portal." -ForegroundColor White
Write-Host ""
Write-Host "  To restrict to specific users/groups, go to:" -ForegroundColor Gray
Write-Host "  Azure Portal → Entra ID → Enterprise Apps → $appName" -ForegroundColor Gray
Write-Host "  → Properties → Assignment Required = Yes" -ForegroundColor Gray
Write-Host "  Then add users/groups under 'Users and groups'" -ForegroundColor Gray
Write-Host ""
Write-Host "  To disable auth later:" -ForegroundColor Gray
Write-Host "  az containerapp auth update -n $ContainerAppName -g $ResourceGroup --enabled false" -ForegroundColor Gray
Write-Host "══════════════════════════════════════════════════════════════`n" -ForegroundColor Green
