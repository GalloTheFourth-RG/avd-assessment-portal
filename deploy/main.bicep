// ============================================================================
// AVD Assessment Portal — Azure Infrastructure
// Deploy: az deployment sub create --location eastus --template-file main.bicep
// ============================================================================

targetScope = 'subscription'

@description('Name for the resource group')
param resourceGroupName string = 'rg-avd-assessment'

@description('Azure region for all resources')
param location string = 'eastus'

@description('Name prefix for resources (lowercase, no spaces)')
@minLength(3)
@maxLength(12)
param namePrefix string = 'avdassess'

@description('Subscription IDs to grant the managed identity Reader access (comma-separated)')
param targetSubscriptionIds string = ''

// ============================================================================
// Resource Group
// ============================================================================
resource rg 'Microsoft.Resources/resourceGroups@2023-07-01' = {
  name: resourceGroupName
  location: location
}

// ============================================================================
// Module: All resources deployed into the resource group
// ============================================================================
module resources 'resources.bicep' = {
  scope: rg
  name: 'avd-assessment-resources'
  params: {
    location: location
    namePrefix: namePrefix
    targetSubscriptionIds: targetSubscriptionIds
  }
}

output portalUrl string = resources.outputs.portalUrl
output managedIdentityPrincipalId string = resources.outputs.managedIdentityPrincipalId
output storageAccountName string = resources.outputs.storageAccountName
