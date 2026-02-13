# ============================================================================
# AVD Assessment Portal — Container Image
# PowerShell 7 + Az Modules + Frontend
#
# The assessment script is fetched from a private GitHub repo at build time.
# Build with: az acr build --build-arg GITHUB_TOKEN=ghp_xxx ...
# ============================================================================

FROM mcr.microsoft.com/azure-powershell:latest

# Install Node.js for frontend build
RUN apt-get update && apt-get install -y curl && \
    curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# Install additional Az modules needed by the assessment script
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

# Fetch assessment script from private repo at build time
ARG GITHUB_TOKEN
ARG SCRIPT_REPO=gallothefourth-rg/enhanced-avd-evidence-pack
ARG SCRIPT_BRANCH=main
ARG SCRIPT_PATH=Get-Enhanced-AVD-EvidencePack.ps1

RUN mkdir -p ./scripts && \
    curl -fsSL \
      -H "Authorization: token ${GITHUB_TOKEN}" \
      -H "Accept: application/vnd.github.v3.raw" \
      "https://api.github.com/repos/${SCRIPT_REPO}/contents/${SCRIPT_PATH}?ref=${SCRIPT_BRANCH}" \
      -o ./scripts/Get-Enhanced-AVD-EvidencePack.ps1 && \
    echo "Script downloaded: $(wc -l < ./scripts/Get-Enhanced-AVD-EvidencePack.ps1) lines"

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
