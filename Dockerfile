# ============================================================================
# AVD Assessment Portal — Container Image
# PowerShell 7 + Az Modules + Frontend
#
# BUILD OPTIONS:
#   Option A — Fetch script from private repo (for CI/automated builds):
#     az acr build --build-arg GITHUB_TOKEN=ghp_xxx ...
#
#   Option B — Local script (for your own builds):
#     Copy script to scripts/ folder first, then build normally.
#     The .gitignore prevents scripts/ from being committed.
# ============================================================================

FROM mcr.microsoft.com/azure-powershell:latest

# Install Node.js for frontend build
RUN apt-get update && apt-get install -y curl && \
    curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# Install additional Az modules
RUN pwsh -Command " \
    Install-Module -Name Az.DesktopVirtualization -Force -AllowClobber -Scope AllUsers; \
    Install-Module -Name Az.Monitor -Force -AllowClobber -Scope AllUsers; \
    Install-Module -Name Az.ResourceGraph -Force -AllowClobber -Scope AllUsers; \
    Install-Module -Name Az.OperationalInsights -Force -AllowClobber -Scope AllUsers; \
    Install-Module -Name Az.Reservations -Force -AllowClobber -Scope AllUsers; \
    Install-Module -Name Az.Advisor -Force -AllowClobber -Scope AllUsers; \
    Install-Module -Name ThreadJob -Force -AllowClobber -Scope AllUsers; \
    "

WORKDIR /app

# Copy backend
COPY backend/ ./backend/

# Copy scripts/ if it exists locally (Option B — local build)
# This COPY won't fail if the directory is empty or missing from context
COPY script[s]/ ./scripts/

# Fetch from private repo if script not present (Option A — CI build)
ARG GITHUB_TOKEN=""
ARG SCRIPT_REPO=gallothefourth-rg/enhanced-avd-evidence-pack
ARG SCRIPT_BRANCH=main
ARG SCRIPT_PATH=Get-Enhanced-AVD-EvidencePack.ps1

RUN mkdir -p ./scripts && \
    if [ ! -f "./scripts/Get-Enhanced-AVD-EvidencePack.ps1" ] && [ -n "${GITHUB_TOKEN}" ]; then \
      echo "Fetching script from ${SCRIPT_REPO}..." && \
      curl -fsSL \
        -H "Authorization: token ${GITHUB_TOKEN}" \
        -H "Accept: application/vnd.github.v3.raw" \
        "https://api.github.com/repos/${SCRIPT_REPO}/contents/${SCRIPT_PATH}?ref=${SCRIPT_BRANCH}" \
        -o ./scripts/Get-Enhanced-AVD-EvidencePack.ps1 && \
      echo "Script downloaded: $(wc -l < ./scripts/Get-Enhanced-AVD-EvidencePack.ps1) lines"; \
    elif [ -f "./scripts/Get-Enhanced-AVD-EvidencePack.ps1" ]; then \
      echo "Using local script: $(wc -l < ./scripts/Get-Enhanced-AVD-EvidencePack.ps1) lines"; \
    else \
      echo "ERROR: No script found. Either copy to scripts/ or provide GITHUB_TOKEN build arg." && exit 1; \
    fi

# Copy and build frontend
COPY frontend/ ./frontend/
RUN cd frontend && npm install && npm run build

EXPOSE 3000

# Copy startup wrapper
COPY startup.ps1 /app/startup.ps1

# Health check
HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=5 \
    CMD pwsh -Command "try { (Invoke-WebRequest -Uri http://localhost:3000/api/health -TimeoutSec 3).StatusCode -eq 200 } catch { exit 1 }"

CMD ["pwsh", "-File", "/app/startup.ps1"]
