# AVD Assessment Portal

A self-hosted web portal that runs the Enhanced AVD Evidence Pack inside your Azure tenant. Deploy once, run assessments on-demand from your browser — no local installs required.

> 🔒 All data stays in your Azure tenant. Secured with Entra ID authentication.

Licensed under [MIT](LICENSE-MIT) or [Apache 2.0](LICENSE-APACHE), at your option.

---

## How It Works

```
┌─────────────────────────────────────────────────────────┐
│  Browser (Entra ID sign-in)                             │
│  ┌───────────────────────────────────────────────────┐  │
│  │  React Dashboard                                  │  │
│  │  • Select subscriptions (auto-discovered)         │  │
│  │  • Configure options & launch assessment          │  │
│  │  • View 17-tab HTML report in-browser             │  │
│  │  • Download CSVs, ZIPs, delete old runs           │  │
│  └────────────────────┬──────────────────────────────┘  │
└───────────────────────┼─────────────────────────────────┘
                        │ HTTPS
┌───────────────────────┼─────────────────────────────────┐
│  Your Azure Tenant    │                                 │
│  ┌────────────────────▼──────────────────────────────┐  │
│  │  Container App (consumption — scales to zero)     │  │
│  │  • PowerShell 7 + Az Modules                     │  │
│  │  • Runs Get-Enhanced-AVD-EvidencePack.ps1 v4.1.0 │  │
│  │  • Managed Identity (no credentials stored)       │  │
│  │  • Async execution — server stays responsive      │  │
│  └────────────────────┬──────────────────────────────┘  │
│                       │                                 │
│  ┌────────────────────▼─────┐  ┌──────────────────────┐ │
│  │  Storage Account         │  │  Your AVD Resources   │ │
│  │  • Assessment results    │  │  • VMs, Host Pools    │ │
│  │  • HTML reports          │  │  • Log Analytics      │ │
│  │  • CSV exports           │  │  • Cost Management    │ │
│  └──────────────────────────┘  └──────────────────────┘ │
└─────────────────────────────────────────────────────────┘
```

---

## Deployment (~15 minutes)

### Prerequisites

- Azure subscription with **Owner** or **Contributor + User Access Administrator**
- [Azure CLI](https://learn.microsoft.com/en-us/cli/azure/install-azure-cli) installed (`az --version`)
- PowerShell 7+ (`pwsh --version`)

### Step 1: Deploy Azure infrastructure

**Option A: One-click deploy**

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fyour-org%2Favd-assessment-portal%2Fmain%2Fazuredeploy.json)

> After clicking, choose a resource group (or create `rg-avd-assessment`), set a unique **namePrefix**, and click **Review + Create**.

**Option B: CLI deploy**

```powershell
az login
az group create -n rg-avd-assessment -l eastus
az deployment group create `
  --resource-group rg-avd-assessment `
  --template-file azuredeploy.json `
  --parameters namePrefix="avdassess"
```

This creates:

| Resource | Purpose | Cost |
|----------|---------|------|
| Container App (consumption) | Runs the portal | ~$0 when idle |
| Storage Account (LRS) | Stores assessment results | ~$0.02/GB/month |
| User-Assigned Managed Identity | Azure API access | Free |
| Log Analytics Workspace | Container logs | ~$2.76/GB ingested |

**Idle cost: ~$0/month** (consumption plan scales to zero)

### Step 2: Configure the container image

The pre-built portal image is published to GitHub Container Registry. Configure your Container App to pull it:

```powershell
az containerapp update `
  --name avdassess-portal `
  --resource-group rg-avd-assessment `
  --image ghcr.io/gallothefourth-rg/avd-assessment-portal:latest
```

> **Note:** The image is public — no registry credentials needed.

### Step 3: Grant permissions to target subscriptions

The managed identity needs Reader access on the subscriptions containing your AVD resources:

```powershell
$principalId = (az identity show -g rg-avd-assessment -n avdassess-identity --query principalId -o tsv)

pwsh deploy/Setup-Permissions.ps1 `
  -PrincipalId $principalId `
  -SubscriptionIds @("your-avd-subscription-id-1", "your-avd-subscription-id-2") `
  -IncludeCostReader `
  -IncludeLogAnalyticsReader
```

Roles assigned:
- **Reader** — enumerate VMs, host pools, session hosts
- **Cost Management Reader** — query actual billed costs
- **Log Analytics Reader** — run KQL queries for user metrics

### Step 4: Enable Entra ID authentication (recommended)

```powershell
pwsh deploy/Setup-Auth.ps1
```

This creates an Entra ID App Registration and enables built-in authentication on the Container App. Users must sign in with their organizational account before accessing the portal.

To restrict access to specific users/groups:
1. Azure Portal → Entra ID → Enterprise Applications → "AVD Assessment Portal"
2. Properties → **Assignment Required** = Yes
3. Users and groups → Add allowed users/groups

### Step 5: Open the portal

```powershell
$url = (az containerapp show -n avdassess-portal -g rg-avd-assessment --query "properties.configuration.ingress.fqdn" -o tsv)
Start-Process "https://$url"
```

---

## Usage

1. **Sign in** with your Entra ID credentials
2. **Tenant ID** auto-detects from the managed identity
3. **Select subscriptions** — auto-discovered from RBAC assignments
4. **Configure options** — lookback days, advisor, cost analysis, PII scrubbing
5. **Start assessment** — runs async; you can monitor progress in real-time
6. **View results** — 17-tab HTML dashboard renders in-browser
7. **Download** individual CSVs or the full ZIP
8. **Manage runs** — view past assessments, delete old runs to free storage

---

## Updating (for end users)

When a new version is released:

```powershell
az containerapp update -n avdassess-portal -g rg-avd-assessment `
  --image ghcr.io/gallothefourth-rg/avd-assessment-portal:latest `
  --set-env-vars "BUILD_ID=$(Get-Date -Format 'yyyyMMddHHmmss')"
```

---

## Publishing New Versions (maintainer only)

```powershell
# Copy latest script into scripts/ (gitignored — never committed)
Copy-Item "path\to\Get-Enhanced-AVD-EvidencePack.ps1" -Destination "scripts\" -Force

# Build and push to GitHub Container Registry
docker build -t ghcr.io/gallothefourth-rg/avd-assessment-portal:latest .
echo $env:GITHUB_TOKEN | docker login ghcr.io -u gallothefourth-rg --password-stdin
docker push ghcr.io/gallothefourth-rg/avd-assessment-portal:latest
```

---

## Security

- **No credentials stored** — Managed Identity for all Azure API access
- **No data leaves your tenant** — results stored in your own Storage Account
- **RBAC-scoped** — identity has only Reader access (cannot modify resources)
- **Entra ID authentication** — only signed-in users can access the portal
- **PII scrubbing** — optional anonymization for external sharing

---

## Repository Structure

```
avd-assessment-portal/
├── deploy/
│   ├── main.bicep                ← Subscription-level deployment (CLI)
│   ├── resources.bicep           ← Container App + ACR + Storage + Identity
│   ├── Setup-Permissions.ps1     ← RBAC role assignments for target subs
│   └── Setup-Auth.ps1            ← Entra ID Easy Auth configuration
├── backend/
│   └── src/
│       ├── server.ps1            ← PowerShell HTTP server (API + static)
│       └── run-assessment.ps1    ← Async assessment runner
├── frontend/
│   ├── src/
│   │   ├── App.jsx               ← React dashboard
│   │   └── main.jsx              ← Entry point
│   ├── package.json
│   └── vite.config.js
├── azuredeploy.json              ← ARM template (Deploy to Azure button)
├── Dockerfile                    ← Fetches script from private repo at build
├── startup.ps1                   ← Container startup wrapper
├── CHANGELOG.md                  ← Release history
├── LICENSE-MIT
├── LICENSE-APACHE
└── README.md
```

> **Note:** The assessment script (`Get-Enhanced-AVD-EvidencePack.ps1`) is hosted in a [separate private repository](https://github.com/gallothefourth-rg/enhanced-avd-evidence-pack) and fetched at container build time. It is not included in this repo.

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| "No subscriptions loaded" | Run `Setup-Permissions.ps1` for target subscriptions |
| Cost data shows $0 | Add `-IncludeCostReader` flag to permissions setup |
| KQL queries fail | Add `-IncludeLogAnalyticsReader` and configure LA workspace IDs |
| 401 after enabling auth | Ensure `--enable-id-token-issuance true` on the app registration |
| Container won't start | Check logs: `az containerapp logs show -n avdassess-portal -g rg-avd-assessment --type console` |
| Assessment fails "Connect-AzAccount" | The evidence pack script should detect managed identity — update to latest `scripts/` version |
| Portal loads, API returns errors | Check managed identity: `az identity show -g rg-avd-assessment -n avdassess-identity` |
| Need to force a fresh deployment | Change any env var: `az containerapp update -n avdassess-portal -g rg-avd-assessment --set-env-vars "BUILD_ID=$(Get-Date -Format 'yyyyMMddHHmmss')"` |

---

## Architecture Decisions

- **Split repo model** — the portal is open source; the assessment script stays private. The Dockerfile fetches it at build time via GitHub API, so the script is never in the public repo or Git history
- **PowerShell HTTP listener** instead of a framework — keeps the container simple (no extra runtime) and runs the `.ps1` assessment script natively
- **Async via saved Az profile** — `Save-AzContext`/`Import-AzContext` lets the background process share the managed identity token without process boundary issues
- **Consumption plan** — scales to zero when idle, meaning near-zero cost between assessments
- **ACR build** — builds the container in Azure, no local Docker installation required
- **Easy Auth** — authentication handled at the platform level (Container Apps), not in application code
