#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Post-deployment: Assigns Reader + Cost Management Reader roles to the managed identity
    on target subscriptions so the assessment can query Azure APIs.

.DESCRIPTION
    Run this AFTER deploying the Bicep template. The managed identity needs:
    - Reader: to enumerate VMs, host pools, session hosts, etc.
    - Cost Management Reader: to query actual billed costs
    - Log Analytics Reader: to run KQL queries

.EXAMPLE
    ./Setup-Permissions.ps1 -PrincipalId "abc-123" -SubscriptionIds @("sub-1", "sub-2")
#>

param(
    [Parameter(Mandatory)]
    [string]$PrincipalId,
    
    [Parameter(Mandatory)]
    [string[]]$SubscriptionIds,
    
    [switch]$IncludeCostReader,
    
    [switch]$IncludeLogAnalyticsReader
)

$ErrorActionPreference = "Stop"

# Role Definition IDs (built-in)
$roles = @{
    "Reader"                    = "acdd72a7-3385-48ef-bd42-f606fba81ae7"
    "Cost Management Reader"    = "72fafb9e-0641-4937-9268-a91bfd8191a3"
    "Log Analytics Reader"      = "73c42c96-874c-492b-b04d-ab87d138a893"
}

$rolesToAssign = @("Reader")
if ($IncludeCostReader) { $rolesToAssign += "Cost Management Reader" }
if ($IncludeLogAnalyticsReader) { $rolesToAssign += "Log Analytics Reader" }

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  AVD Assessment Portal — RBAC Permission Setup              ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

Write-Host "`nManaged Identity: $PrincipalId" -ForegroundColor Gray
Write-Host "Subscriptions:    $($SubscriptionIds.Count)" -ForegroundColor Gray
Write-Host "Roles:            $($rolesToAssign -join ', ')" -ForegroundColor Gray

foreach ($subId in $SubscriptionIds) {
    Write-Host "`n── Subscription: $subId ──" -ForegroundColor Yellow
    
    foreach ($roleName in $rolesToAssign) {
        $roleDefId = $roles[$roleName]
        $scope = "/subscriptions/$subId"
        
        try {
            # Check if assignment already exists
            $existing = Get-AzRoleAssignment -ObjectId $PrincipalId -RoleDefinitionId $roleDefId -Scope $scope -ErrorAction SilentlyContinue
            if ($existing) {
                Write-Host "  ✓ $roleName — already assigned" -ForegroundColor Green
                continue
            }
            
            New-AzRoleAssignment `
                -ObjectId $PrincipalId `
                -RoleDefinitionId $roleDefId `
                -Scope $scope `
                -ErrorAction Stop | Out-Null
            
            Write-Host "  ✓ $roleName — assigned" -ForegroundColor Green
        }
        catch {
            Write-Host "  ✗ $roleName — $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host "`n✅ Done! The managed identity can now access these subscriptions." -ForegroundColor Green
Write-Host "   Role assignments may take 1-2 minutes to propagate.`n" -ForegroundColor Gray
