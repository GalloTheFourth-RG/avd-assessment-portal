# Changelog

All notable changes to the AVD Assessment Portal are documented here.

---

## [2.0.0] — 2026-02-13

### Portal v2.0 — Authentication, Async, Delete Runs

**Entra ID Authentication**
- Integrated Azure Container Apps Easy Auth with Microsoft identity provider
- Users must sign in with organizational Entra ID credentials before accessing the portal
- `Setup-Auth.ps1` automates app registration, client secret, and Easy Auth configuration
- Support for restricting access to specific users/groups via Enterprise App assignments
- Server validates `X-MS-CLIENT-PRINCIPAL-ID` header on all API endpoints (except health)

**Async Assessment Execution**
- Assessments now run as background processes using `Save-AzContext`/`Import-AzContext`
- Server stays responsive during assessments — polling, status checks, and static files all work
- `run-assessment.ps1` loads saved Az profile, runs the evidence pack, uploads results
- Automatic re-login after assessment completes (evidence pack script resets Az context)

**Run Management**
- Delete past assessment runs from the UI (removes all blobs from storage)
- Confirmation prompt before deletion
- Runs list updates immediately after deletion

**Managed Identity Resilience**
- `Ensure-AzLogin` function re-establishes managed identity on every API call
- Evidence pack script detects `ManagedService` account type and skips interactive login
- Saved Az profile (`Save-AzContext`) persisted to `/tmp/az-profile.json` for child processes

**Infrastructure**
- Updated Bicep templates for ACR-based deployment (no GitHub Container Registry dependency)
- `Setup-Permissions.ps1` for multi-subscription RBAC assignment
- `Setup-Auth.ps1` for Entra ID Easy Auth configuration
- Added `.txt`, `.csv`, `.zip` MIME types for result file downloads

---

## [1.0.0] — 2026-02-13

### Portal v1.0 — Initial Release

**Core Platform**
- Self-hosted web portal running in Azure Container App (consumption plan, scales to zero)
- PowerShell 7 HTTP listener serving React frontend and REST API
- Runs Get-Enhanced-AVD-EvidencePack.ps1 v4.1.0 on-demand
- User-assigned managed identity for Azure API access (no stored credentials)

**Frontend (React + Vite)**
- Dark-themed dashboard with sidebar navigation
- Auto-detect tenant ID from managed identity
- Subscription discovery with multi-select
- Full configuration panel: lookback days, Log Analytics, Advisor, Reservations, PII scrubbing
- Real-time assessment progress with elapsed timer
- In-browser HTML report viewer (17-tab dashboard)
- File list with individual download links
- Past runs history

**Backend API**
- `GET /api/health` — health check with Azure connection status
- `GET /api/subscriptions` — list accessible subscriptions from managed identity
- `POST /api/assess` — start assessment with configuration
- `GET /api/assess/{runId}` — poll assessment status
- `GET /api/results/{runId}` — list result files
- `GET /api/results/{runId}/{filename}` — download/view individual files
- `GET /api/runs` — list past assessment runs

**Infrastructure**
- Bicep templates for one-command Azure deployment
- Container App + Storage Account + Managed Identity + Log Analytics
- ACR-based container builds (no local Docker required)
- `Setup-Permissions.ps1` for RBAC assignment on target subscriptions

**Assessment Script v4.1.0**
- Managed identity detection — skips interactive login when running under `ManagedService`
- Full evidence pack: VM right-sizing, cost analysis, security posture, user experience metrics
- W365 comparison with workload-aware plan matching
- PII scrubbing, Login Time Analysis, Connection Success Rate
- 34 output files including HTML dashboard, CSVs, and ZIP archive
