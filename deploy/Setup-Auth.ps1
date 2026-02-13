#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Enable Entra ID authentication (Easy Auth) on the AVD Assessment Portal.

.DESCRIPTION
    Creates an Entra ID App Registration and configures Azure Container Apps
    built-in authentication so only signed-in users can access the portal.

    Run AFTER the portal is deployed and working.

.EXAMPLE
    ./Setup-Auth.ps1 -ResourceGroup "rg-avd-assessment" -ContainerAppName "avdassess-portal"
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
if (-not $fqdn) { throw "Container App '$ContainerAppName' not found in resource group '$ResourceGroup'" }

$portalUrl = "https://$fqdn"
$tenantId = (az account show --query "tenantId" -o tsv)
Write-Host "  Portal URL: $portalUrl" -ForegroundColor Gray
Write-Host "  Tenant:     $tenantId" -ForegroundColor Gray

# ── Step 1: Create or find App Registration ──
Write-Host "`n── Step 1: App Registration ──" -ForegroundColor Yellow

$appName = "AVD Assessment Portal"
$redirectUri = "$portalUrl/.auth/login/aad/callback"

$appId = az ad app list --display-name $appName --query "[0].appId" -o tsv 2>$null
if ($appId) {
    Write-Host "  ✓ Found existing: $appId" -ForegroundColor Green
    az ad app update --id $appId --web-redirect-uris $redirectUri | Out-Null
} else {
    $appId = (az ad app create `
        --display-name $appName `
        --web-redirect-uris $redirectUri `
        --sign-in-audience "AzureADMyOrg" `
        --query "appId" -o tsv)
    Write-Host "  ✓ Created: $appId" -ForegroundColor Green
}

# Enable ID token (required for Easy Auth callback)
az ad app update --id $appId --enable-id-token-issuance true | Out-Null
Write-Host "  ✓ ID token issuance enabled" -ForegroundColor Green

# ── Step 2: Create client secret ──
Write-Host "`n── Step 2: Client Secret ──" -ForegroundColor Yellow
$clientSecret = (az ad app credential reset --id $appId --display-name "portal-auth" --query "password" -o tsv)
Write-Host "  ✓ Secret created" -ForegroundColor Green

# ── Step 3: Configure Easy Auth ──
Write-Host "`n── Step 3: Container App Authentication ──" -ForegroundColor Yellow

# Configure Microsoft provider (--issuer and --tenant-id cannot both be set)
az containerapp auth microsoft update `
    -n $ContainerAppName `
    -g $ResourceGroup `
    --client-id $appId `
    --client-secret $clientSecret `
    --issuer "https://login.microsoftonline.com/$tenantId/v2.0" `
    --yes | Out-Null

Write-Host "  ✓ Microsoft provider configured" -ForegroundColor Green

# Enable auth — redirect unauthenticated users to login
az containerapp auth update `
    -n $ContainerAppName `
    -g $ResourceGroup `
    --unauthenticated-client-action RedirectToLoginPage `
    --enabled true | Out-Null

Write-Host "  ✓ Authentication enabled" -ForegroundColor Green

# Restart to pick up secret
$revision = (az containerapp revision list -n $ContainerAppName -g $ResourceGroup --query "[0].name" -o tsv)
if ($revision) {
    az containerapp revision restart -n $ContainerAppName -g $ResourceGroup --revision $revision 2>$null
    Write-Host "  ✓ Container restarted (may take 30-60s)" -ForegroundColor Green
}

# ── Summary ──
Write-Host "`n══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  ✅ Entra ID authentication configured!" -ForegroundColor Green
Write-Host ""
Write-Host "  App Registration:  $appId" -ForegroundColor Gray
Write-Host "  Tenant:            $tenantId" -ForegroundColor Gray
Write-Host "  Portal:            $portalUrl" -ForegroundColor Gray
Write-Host ""
Write-Host "  All users in your tenant can now sign in." -ForegroundColor White
Write-Host ""
Write-Host "  To restrict to specific users/groups:" -ForegroundColor Yellow
Write-Host "    Azure Portal → Entra ID → Enterprise Applications" -ForegroundColor Gray
Write-Host "    → '$appName' → Properties → Assignment Required = Yes" -ForegroundColor Gray
Write-Host "    → Users and groups → Add the allowed users/groups" -ForegroundColor Gray
Write-Host ""
Write-Host "  To disable auth:" -ForegroundColor Yellow
Write-Host "    az containerapp auth update -n $ContainerAppName -g $ResourceGroup --enabled false" -ForegroundColor Gray
Write-Host "══════════════════════════════════════════════════════════════`n" -ForegroundColor Green
