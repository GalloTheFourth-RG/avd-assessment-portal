# AVD Assessment Portal

A self-hosted web portal that runs the [Enhanced AVD Evidence Pack](https://github.com/your-org/enhanced-avd-evidence-pack) inside your Azure tenant. Deploy once, run assessments on-demand from your browser.

> 🔒 Private repo — all data stays in your Azure tenant.

---

## How It Works

```
┌─────────────────────────────────────────────────────────┐
│  Your Browser                                           │
│  ┌───────────────────────────────────────────────────┐  │
│  │  React Dashboard                                  │  │
│  │  • Configure assessment parameters                │  │
│  │  • Launch runs, monitor progress                  │  │
│  │  • View 17-tab HTML report in-browser             │  │
│  │  • Download CSVs and ZIP files                    │  │
│  └────────────────────┬──────────────────────────────┘  │
└───────────────────────┼─────────────────────────────────┘
                        │ HTTPS
┌───────────────────────┼─────────────────────────────────┐
│  Your Azure Tenant    │                                 │
│  ┌────────────────────▼──────────────────────────────┐  │
│  │  Container App (consumption — scales to zero)     │  │
│  │  • PowerShell 7 + Az Modules                     │  │
│  │  • Runs Get-Enhanced-AVD-EvidencePack.ps1         │  │
│  │  • Authenticates via Managed Identity             │  │
│  │  • No credentials stored, no keys to rotate       │  │
│  └────────────────────┬──────────────────────────────┘  │
│                       │                                 │
│  ┌────────────────────▼──────┐  ┌─────────────────────┐ │
│  │  Storage Account          │  │  Your AVD Resources  │ │
│  │  • Assessment results     │  │  • VMs, Host Pools   │ │
│  │  • HTML reports           │  │  • Log Analytics     │ │
│  │  • CSV exports            │  │  • Cost Management   │ │
│  └───────────────────────────┘  └─────────────────────┘ │
└─────────────────────────────────────────────────────────┘
```

---

## Deployment (15 minutes)

### Prerequisites

- Azure subscription with Owner or Contributor + User Access Admin
- Azure CLI installed (`az --version`)
- GitHub account (for container registry)

### Step 1: Clone this repo

```bash
git clone https://github.com/your-org/avd-assessment-portal.git
cd avd-assessment-portal
```

### Step 2: Deploy Azure infrastructure

```bash
# Login to Azure
az login

# Deploy (creates: resource group, container app, storage, managed identity)
az deployment sub create \
  --location eastus \
  --template-file deploy/main.bicep \
  --parameters \
    resourceGroupName="rg-avd-assessment" \
    namePrefix="avdassess" \
    targetSubscriptionIds="sub-id-1,sub-id-2"
```

This creates:
| Resource | Purpose | Cost |
|----------|---------|------|
| Container App (consumption) | Runs the portal | ~$0 when idle, ~$0.05/run |
| Storage Account (LRS) | Stores results | ~$0.02/GB/month |
| Managed Identity | Azure API access | Free |
| Log Analytics | Container logs | ~$2.76/GB ingested |

**Total idle cost: ~$0/month** (consumption plan scales to zero)

### Step 3: Assign permissions

The managed identity needs Reader access on your AVD subscriptions:

```powershell
# Get the managed identity principal ID from deployment output
$principalId = (az deployment sub show --name avd-assessment-resources --query properties.outputs.managedIdentityPrincipalId.value -o tsv)

# Run the setup script
pwsh deploy/Setup-Permissions.ps1 \
  -PrincipalId $principalId \
  -SubscriptionIds @("your-avd-sub-1", "your-avd-sub-2") \
  -IncludeCostReader \
  -IncludeLogAnalyticsReader
```

### Step 4: Build and push the container

```bash
# Build the Docker image
docker build -t ghcr.io/your-org/avd-assessment-portal:latest .

# Push to GitHub Container Registry
echo $GITHUB_TOKEN | docker login ghcr.io -u your-username --password-stdin
docker push ghcr.io/your-org/avd-assessment-portal:latest

# Update the Container App to use your image
az containerapp update \
  --name avdassess-portal \
  --resource-group rg-avd-assessment \
  --image ghcr.io/your-org/avd-assessment-portal:latest
```

### Step 5: Open the portal

```bash
# Get the portal URL
az containerapp show --name avdassess-portal --resource-group rg-avd-assessment --query properties.configuration.ingress.fqdn -o tsv
```

Navigate to `https://avdassess-portal.<region>.azurecontainerapps.io`

---

## Usage

1. Open the portal URL in your browser
2. Enter your Tenant ID
3. Select subscriptions (auto-discovered from managed identity access)
4. Configure options (lookback days, advisor, cost analysis, PII scrubbing)
5. Click **Start Full Assessment**
6. Wait 15-45 minutes (progress updates every 3 seconds)
7. View the 17-tab HTML dashboard directly in-browser
8. Download individual CSVs or the full ZIP

---

## Security

- **No credentials stored** — uses Azure Managed Identity
- **No data leaves your tenant** — results stored in your Storage Account
- **RBAC-scoped** — the identity only has Reader access (cannot modify resources)
- **PII scrubbing** — optional anonymization for external sharing
- **Private Container App** — optionally restrict to VNet with internal ingress

### Restricting access

To limit portal access to your corporate network:

```bash
# Add VNet integration (optional)
az containerapp env update \
  --name avdassess-env \
  --resource-group rg-avd-assessment \
  --internal-only true
```

---

## Repository Structure

```
avd-assessment-portal/
├── deploy/
│   ├── main.bicep              ← Subscription-level deployment
│   ├── resources.bicep         ← Container App + Storage + Identity
│   └── Setup-Permissions.ps1   ← RBAC role assignment helper
├── backend/
│   └── src/
│       └── server.ps1          ← PowerShell HTTP server (API + static files)
├── frontend/
│   ├── src/
│   │   ├── App.jsx             ← React dashboard
│   │   └── main.jsx            ← Entry point
│   ├── package.json
│   └── vite.config.js
├── scripts/
│   └── Get-Enhanced-AVD-EvidencePack.ps1  ← The assessment script (v4.1.0)
├── Dockerfile                  ← Container: PowerShell 7 + Az + Node + Frontend
├── .github/workflows/
│   └── build.yml               ← CI/CD: build container on push
└── README.md
```

---

## Updating the Assessment Script

When a new version of the evidence pack script is released:

1. Copy the updated `.ps1` file into `scripts/`
2. Commit and push — GitHub Actions rebuilds the container
3. The Container App auto-pulls the new image (or run `az containerapp update`)

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| "No subscriptions loaded" | Managed identity needs Reader on target subscriptions |
| Cost data shows $0 | Add Cost Management Reader role |
| KQL queries fail | Add Log Analytics Reader + configure LA workspace IDs |
| Container won't start | Check logs: `az containerapp logs show --name avdassess-portal -g rg-avd-assessment` |
| Portal loads but API fails | Check managed identity: `az identity show -g rg-avd-assessment -n avdassess-identity` |
