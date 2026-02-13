// ============================================================================
// AVD Assessment Portal — Resource Module
// Container App + ACR + Storage + Managed Identity
// ============================================================================

param location string
param namePrefix string
param targetSubscriptionIds string

// ============================================================================
// Managed Identity
// ============================================================================
resource managedIdentity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: '${namePrefix}-identity'
  location: location
}

// ============================================================================
// Azure Container Registry (Basic tier — stores the portal image)
// ============================================================================
resource acr 'Microsoft.ContainerRegistry/registries@2023-07-01' = {
  name: '${namePrefix}acr'
  location: location
  sku: { name: 'Basic' }
  properties: {
    adminUserEnabled: true
  }
}

// ============================================================================
// Storage Account — assessment results
// ============================================================================
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: '${namePrefix}store'
  location: location
  kind: 'StorageV2'
  sku: { name: 'Standard_LRS' }
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    supportsHttpsTrafficOnly: true
  }
}

resource blobService 'Microsoft.Storage/storageAccounts/blobServices@2023-01-01' = {
  parent: storageAccount
  name: 'default'
}

resource resultsContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-01-01' = {
  parent: blobService
  name: 'results'
  properties: {
    publicAccess: 'None'
  }
}

// Storage Blob Data Contributor for managed identity
resource storageBlobRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(storageAccount.id, managedIdentity.id, 'ba92f5b4-2d11-453d-a403-e96b0029c9fe')
  scope: storageAccount
  properties: {
    principalId: managedIdentity.properties.principalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', 'ba92f5b4-2d11-453d-a403-e96b0029c9fe')
  }
}

// ============================================================================
// Log Analytics Workspace — container logs
// ============================================================================
resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2022-10-01' = {
  name: '${namePrefix}-logs'
  location: location
  properties: {
    sku: { name: 'PerGB2018' }
    retentionInDays: 30
  }
}

// ============================================================================
// Container App Environment (consumption — scales to zero)
// ============================================================================
resource containerAppEnv 'Microsoft.App/managedEnvironments@2023-05-01' = {
  name: '${namePrefix}-env'
  location: location
  properties: {
    appLogsConfiguration: {
      destination: 'log-analytics'
      logAnalyticsConfiguration: {
        customerId: logAnalytics.properties.customerId
        sharedKey: logAnalytics.listKeys().primarySharedKey
      }
    }
  }
}

// ============================================================================
// Container App — the assessment portal
// ============================================================================
resource containerApp 'Microsoft.App/containerApps@2023-05-01' = {
  name: '${namePrefix}-portal'
  location: location
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${managedIdentity.id}': {}
    }
  }
  properties: {
    managedEnvironmentId: containerAppEnv.id
    configuration: {
      ingress: {
        external: true
        targetPort: 3000
        transport: 'http'
        allowInsecure: false
      }
      registries: []
      secrets: []
    }
    template: {
      containers: [
        {
          name: 'portal'
          // Placeholder image — replaced after ACR build (Step 2 in README)
          image: 'mcr.microsoft.com/azure-powershell:latest'
          command: [ 'pwsh', '-File', '/app/startup.ps1' ]
          resources: {
            cpu: json('1.0')
            memory: '2Gi'
          }
          env: [
            { name: 'AZURE_CLIENT_ID', value: managedIdentity.properties.clientId }
            { name: 'STORAGE_ACCOUNT_NAME', value: storageAccount.name }
            { name: 'STORAGE_CONTAINER', value: 'results' }
            { name: 'TARGET_SUBSCRIPTION_IDS', value: targetSubscriptionIds }
          ]
        }
      ]
      scale: {
        minReplicas: 0
        maxReplicas: 3
        rules: [
          {
            name: 'http-rule'
            http: {
              metadata: {
                concurrentRequests: '10'
              }
            }
          }
        ]
      }
    }
  }
}

// ============================================================================
// Outputs
// ============================================================================
output portalUrl string = 'https://${containerApp.properties.configuration.ingress.fqdn}'
output managedIdentityPrincipalId string = managedIdentity.properties.principalId
output managedIdentityClientId string = managedIdentity.properties.clientId
output storageAccountName string = storageAccount.name
output acrName string = acr.name
output acrLoginServer string = acr.properties.loginServer
