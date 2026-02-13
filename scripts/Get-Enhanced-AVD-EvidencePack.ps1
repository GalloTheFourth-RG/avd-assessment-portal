<#
.SYNOPSIS
    Enhanced Azure Virtual Desktop (AVD) Evidence Collection with AVD-Aware Right-Sizing and Analysis
    Version: 4.1.0

.DESCRIPTION
    ⚠️  DISCLAIMER: This script is provided AS-IS with NO WARRANTY and NO SUPPORT.
    Use at your own risk. Test thoroughly before using in production environments.
    Always review recommendations before implementing changes.
    
    Comprehensive Azure Virtual Desktop analysis tool that collects operational telemetry and provides
    intelligent recommendations based on AVD-specific capacity metrics.
    
    KEY FEATURES:
    ═══════════════════════════════════════════════════════════════════════════════
    
    🎯 AVD-AWARE RIGHT-SIZING (Not Just CPU%)
       • Analyzes sessions-per-vCPU (Microsoft recommends 4-6 for optimal performance)
       • Calculates memory-per-user (minimum 2GB, optimal 4GB)
       • Prevents unnecessary upsizing when CPU is high but capacity is available
       • Identifies workload issues vs. sizing issues
       • Supports 75+ VM SKUs: v4, v5, v6 generations, AMD variants (ads, as)
    
    💰 COST INTELLIGENCE
       • Actual Cost: Queries Azure Cost Management API (AmortizedCost) for real billed costs
       • RI/SP Aware: Detects Reserved Instance and Savings Plan amortization, RG-level fallback
       • Infrastructure Costs: Networking, storage accounts, AVD service, Log Analytics per RG
       • Per-User Cost: Transforms fleet cost into per-seat business metric for W365 comparison
       • PAYG Fallback: Estimates from comprehensive pricing engine if no Cost Reader role
       • Orphaned Resources: Detects unattached disks, NICs, and public IPs with monthly waste
       • W365 Cloud PC Readiness: User-count-aware comparison (pooled = users, not VMs)
    
    ☁️ W365 CLOUD PC READINESS
       • Per-host-pool fit scoring (0-100) with Strong/Consider/Keep recommendations
       • User-count-aware cost comparison: unique users for pooled, concurrent for Frontline
       • Feature gap analysis: 11 capabilities compared (multi-session, autoscale, GPU, etc.)
       • Per-user cost table with AVD vs W365 verdict per pool
       • TCO comparison including storage, networking, and management overhead
       • Usage-based SKU right-sizing from actual workload metrics
       • Pilot pool recommendation with readiness scoring
       • 12-month cost projection with migration timeline
    
    🏗️ ZONE RESILIENCY & ALLOCATION RESILIENCE
       • Evaluates high-availability configuration across availability zones
       • SKU diversity analysis — flags single-SKU-family risk for allocation failures
       • Recommends alternative compatible SKUs from same workload class
       • Calculates resiliency scores per host pool
    
    🔒 SECURITY POSTURE SCORING
       • Trusted Launch, Secure Boot, vTPM, Host Encryption coverage per host pool
       • Security score (0-100, graded A-F) with hardening guidance
       • Ephemeral OS disk and private endpoint checks
       • Priority Matrix flags pools with grade D/F
    
    📈 USER EXPERIENCE SCORING
       • Composite UX Score (0-100): profile load time + RTT + disconnect rate + errors
       • Per-host-pool scoring with sub-component breakdowns
       • Login Time Analysis: Avg/P50/P95/Max per pool with Excellent/Good/Fair/Poor ratings
       • Connection Success Rate: Per-pool failure detection with Investigate thresholds
       • Profile Health Analysis: P95 load times with Good/Warning/Critical severity
       • Disconnect Reason Analysis: categorized by AVD CodeSymbolic values
    
    🌍 NETWORK READINESS
       • Geographic latency analysis — gateway-to-host distance with expected RTT baselines
       • RDP Shortpath detection — surfaces UDP vs TCP relay usage
       • Subnet capacity, NSG coverage, VNet DNS, peering health, private endpoints
       • Network findings with specific remediation steps
    
    🖼️ GOLDEN IMAGE ASSESSMENT
       • Image source classification (Gallery / Marketplace / ManagedImage / Custom)
       • OS generation and build freshness detection (24H2, 23H2, EOL detection)
       • Azure Compute Gallery version age and replication verification
       • Golden Image Maturity Score (0-100, graded A-F)
       • Context-aware "Notes from the Field" best practice guidance
    
    📊 SESSION DENSITY & HOST HEALTH
       • Optimizes user distribution based on Microsoft AVD best practices
       • Session host health monitoring with step-by-step remediation per finding
       • Drain mode detection with effective capacity calculations
       • Profile load time analysis, connection errors, disconnect rates
    
    ⚡ OPERATIONAL FEATURES
       • Dry Run Mode: Preview runtime and cost before collecting data
       • Quick Summary: 2-minute health check without metrics
       • Resume Capability: Continue from failed runs using automatic checkpoints
       • PII Scrubbing: Anonymize usernames, VM names, IPs for safe report sharing
       • Branded Reports: Custom company name, logo, analyst name on HTML dashboard
       • Priority Matrix with Quick Wins / Plan / Consider quadrants
       • 17-tab interactive HTML dashboard with sortable tables

.PARAMETER TenantId
    [Required] Azure AD Tenant ID where AVD resources are deployed.
    Format: GUID (e.g., "12345678-1234-1234-1234-123456789abc")
    
    To find your Tenant ID:
    • Azure Portal → Azure Active Directory → Properties → Tenant ID
    • Or run: (Get-AzContext).Tenant.Id

.PARAMETER SubscriptionIds
    [Required] Array of Azure Subscription IDs to analyze.
    Format: Array of GUIDs
    
    Example (Single):    @("12345678-1234-1234-1234-123456789abc")
    Example (Multiple):  @("sub-guid-1", "sub-guid-2", "sub-guid-3")
    
    To list your subscriptions:
    Get-AzSubscription | Select-Object Name, Id

.PARAMETER LogAnalyticsWorkspaceResourceIds
    [Optional] Array of Log Analytics workspace resource IDs for detailed session analytics.
    
    Format: Full resource ID path
    Example: @("/subscriptions/{sub-id}/resourceGroups/{rg}/providers/Microsoft.OperationalInsights/workspaces/{workspace}")
    
    To find workspace IDs:
    Get-AzOperationalInsightsWorkspace | Select-Object Name, ResourceId

.PARAMETER SkipAzureMonitorMetrics
    [Optional] Skip Azure Monitor metrics collection (CPU, memory).
    
    When to use:
    ✅ Configuration-only audits
    ✅ Zone resiliency checks
    ✅ Quick pre-meeting reviews
    
    Impact:
    ⚠️ Right-sizing recommendations will be limited or unavailable
    ⚠️ Can't calculate sessions-per-vCPU or memory-per-user
    ✅ Reduces runtime from hours to minutes
    
    Example: -SkipAzureMonitorMetrics

.PARAMETER SkipLogAnalyticsQueries
    [Optional] Skip Log Analytics queries for session data.
    
    Impact:
    • Faster runtime (saves ~2-5 minutes)
    • Less detailed session analytics
    • Still collects session host status from Azure API
    
    Example: -SkipLogAnalyticsQueries

.PARAMETER MetricsLookbackDays
    [Optional] Number of days to look back for metrics data.
    Default: 7 days
    Range: 1-30 days
    
    Recommendations:
    • 7 days (default) - Good for typical analysis, balances accuracy and speed
    • 14 days - Better trend analysis, recommended for capacity planning
    • 30 days - Full monthly view, best for seasonal workloads
    • 3 days - Quick spot check
    
    Impact on Runtime:
    • 7 days: ~15 min for 500 VMs (PowerShell 7)
    • 14 days: ~20 min for 500 VMs
    • 30 days: ~30 min for 500 VMs
    
    Example: -MetricsLookbackDays 14

.PARAMETER MetricsTimeGrainMinutes
    [Optional] Time granularity for metrics collection.
    Default: 15 minutes
    Options: 5, 15, 30, 60 minutes
    
    Trade-offs:
    • 5 min: Most detailed, 4x longer runtime
    • 15 min (default): Good balance
    • 60 min: Faster but may miss short spikes
    
    Example: -MetricsTimeGrainMinutes 30

.PARAMETER IncludeIncidentWindowQueries
    [Optional] Collect metrics for a specific incident time window.
    Use with IncidentWindowStart and IncidentWindowEnd.
    
    Perfect for:
    • Post-incident analysis
    • Comparing baseline vs. incident performance
    • Root cause investigation
    
    Example: -IncludeIncidentWindowQueries -IncidentWindowStart (Get-Date "2026-02-10 14:00") -IncidentWindowEnd (Get-Date "2026-02-10 16:00")

.PARAMETER IncidentWindowStart
    [Optional] Start time for incident window analysis.
    Default: 14 days ago
    Format: DateTime object
    
    Example: (Get-Date "2026-02-10 14:00:00")

.PARAMETER IncidentWindowEnd
    [Optional] End time for incident window analysis.
    Default: Current time
    Format: DateTime object
    
    Example: (Get-Date "2026-02-10 16:00:00")

.PARAMETER CreateZip
    [Optional] Create a ZIP archive of all output files.
    
    Useful for:
    • Sharing results with team members
    • Archiving for compliance
    • Uploading to ticket systems
    • Email attachments
    
    Example: -CreateZip

.PARAMETER CpuRightSizingThreshold
    [Optional] Average CPU percentage below which VMs are candidates for downsizing.
    Default: 40%
    Range: 30-50%
    
    Recommendations:
    • 30% - Aggressive cost optimization
    • 40% (default) - Balanced approach
    • 50% - Conservative, more headroom
    
    Note: AVD-aware logic ALSO checks session density before downsizing
    
    Example: -CpuRightSizingThreshold 35

.PARAMETER CpuOverloadThreshold
    [Optional] Peak CPU percentage above which VMs need investigation.
    Default: 80%
    Range: 70-90%
    
    Important: AVD-aware logic checks session density BEFORE recommending upsize!
    • High CPU + High session density (>6/vCPU) = Genuinely need more capacity
    • High CPU + Low session density (<4/vCPU) = Workload issue, NOT sizing issue
    
    Example: -CpuOverloadThreshold 75

.PARAMETER MemoryLowThreshold
    [Optional] Available memory (bytes) threshold for memory pressure detection.
    Default: 1073741824 (1 GB)
    
    Note: AVD-aware logic uses memory-per-session instead of total memory
    • Minimum: 2 GB per user
    • Optimal: 4 GB per user
    
    Example: -MemoryLowThreshold 2147483648  # 2 GB
    
.PARAMETER MinimumZonesForResiliency
    [Optional] Minimum number of availability zones for HA recommendation.
    Default: 2
    
    Microsoft Best Practice: Distribute VMs across 2+ zones for high availability
    
    Example: -MinimumZonesForResiliency 3

.PARAMETER SessionDensityTarget
    [Optional] Target session density for optimal utilization.
    Default: 0.7 (70%)
    Range: 0.5-0.9
    
    Used for capacity planning calculations.
    
    Example: -SessionDensityTarget 0.8

.PARAMETER SkipDisclaimer
    [Optional] Skip the interactive disclaimer prompt.
    
    Use for:
    • Automation scenarios
    • Scheduled tasks
    • CI/CD pipelines
    
    Example: -SkipDisclaimer

.PARAMETER GenerateHtmlReport
    [Optional] Generate an interactive HTML dashboard in addition to CSV files.
    
    Produces a 17-tab single-file HTML report:
    • Executive Summary with Priority Matrix (Quick Wins / Plan / Consider)
    • Right-Sizing with AVD-aware recommendations and SKU diversity
    • Cost Breakdown with actual billing + PAYG estimates per VM
    • Host Health with remediation steps for every finding
    • Network Readiness — RDP Shortpath, subnet capacity, NSG, DNS, peering
    • Storage Optimization — ephemeral disk candidates, disk type findings
    • Golden Image Assessment — maturity score, Notes from the Field
    • Connections & Logins — profile load, RTT, disconnect reason analysis
    • Zone Resiliency — zone distribution and resiliency scores
    • W365 Readiness — per-pool fit scores and cost comparison
    • Security Posture — Trusted Launch/Secure Boot/vTPM scores per pool
    • UX Scores — composite user experience scoring with sub-components
    • Orphaned Resources — waste detection with cleanup commands
    • Reservations — RI coverage (if -IncludeReservationAnalysis)
    • Azure Advisor — recommendations (if -IncludeAzureAdvisor)
    • Incident Analysis — baseline vs incident comparison (if -IncludeIncidentWindowQueries)
    
    Example: -GenerateHtmlReport

.PARAMETER IncludeAzureAdvisor
    [Optional] Include Azure Advisor recommendations in the analysis.
    
    Adds ~1 minute per subscription.
    Provides Microsoft's built-in recommendations for cost, performance, security.
    
    Example: -IncludeAzureAdvisor

.PARAMETER DryRun
    [NEW in v2.2.0] Preview mode - shows what would be analyzed WITHOUT collecting data.
    
    Displays:
    • Exact VM count per subscription
    • Estimated total runtime
    • Number of API calls
    • Output size estimate
    • What features will run
    
    Perfect for:
    ✅ Planning overnight runs
    ✅ Validating parameters before running
    ✅ Estimating costs before execution
    ✅ Budget approval presentations
    
    Runtime: ~30 seconds
    
    Example: -DryRun

.PARAMETER QuickSummary
    [NEW in v2.2.0] Fast 2-3 minute configuration health check without metrics collection.
    
    Shows:
    • VM SKU distribution
    • Zone usage percentages
    • Storage types
    • Immediate configuration issues
    • Host pool settings
    
    Skips:
    • Azure Monitor metrics (the slow part!)
    • Right-sizing recommendations
    • Performance analysis
    
    Perfect for:
    ✅ Daily health checks before meetings
    ✅ Incident response (is anything obviously broken?)
    ✅ Configuration compliance audits
    ✅ Quick "what do we have" reports
    
    Runtime: 2-3 minutes regardless of environment size
    
    Example: -QuickSummary

.PARAMETER ResumeFrom
    [NEW in v2.2.0] Resume from a previous incomplete run using automatic checkpoints.
    
    How it works:
    1. Script saves checkpoints after major steps
    2. If run fails, checkpoints are preserved in .checkpoints folder
    3. Resume loads checkpoints and skips completed steps
    
    Format: Folder path of incomplete run
    
    Time Savings:
    • 100 VMs: Save ~5 min
    • 500 VMs: Save ~15 min
    • 1700 VMs: Save ~45 min
    
    Perfect for:
    ✅ Network disconnects
    ✅ Throttling errors
    ✅ Power failures
    ✅ Manual stops (oops, I have a meeting!)
    
    Example: -ResumeFrom "Enhanced-AVD-EvidencePack-20260211-120000"

.PARAMETER SkipActualCosts
    [NEW in v4.0.0] Skip Azure Cost Management API queries.
    
    By default, the script automatically queries Azure Cost Management for actual
    billed costs (last 30 days) per VM using AmortizedCost, which includes RI and
    Savings Plan amortization. Cost queries are scoped to AVD resource groups.
    
    The script gracefully handles permission issues:
    • If Cost Management Reader role is missing, it warns and falls back to PAYG estimates
    • Per-subscription — one sub failing doesn't affect others
    • For RI-covered VMs showing $0 compute, runs RG-level query to capture amortization
    
    Use -SkipActualCosts to:
    • Speed up runs when you don't need cost data
    • Avoid Cost Management API calls in restricted environments
    • Force PAYG estimates only
    
    Example: -SkipActualCosts

.PARAMETER ScrubPII
    [NEW in v4.1.0] Anonymize personally identifiable information in HTML report and CSV exports.
    
    Replaces usernames, VM names, host pool names, subscription IDs, resource groups,
    and IP addresses with consistent anonymous identifiers (e.g., User-A1B2, Host-3F7C).
    Uses per-run salt with SHA256 hashing — same entity maps to same ID within a run
    (cross-referenceable across tabs) but different runs produce different IDs.
    
    Use -ScrubPII when:
    • Sharing reports with external consultants or vendors
    • Uploading to SharePoint or ticketing systems
    • Including in customer deliverables
    • Compliance requires data minimization
    
    Example: -ScrubPII

.PARAMETER CompanyName
    [Optional] Company name displayed in the HTML report header/title bar.
    
    Example: -CompanyName "Contoso Ltd"

.PARAMETER LogoPath
    [Optional] Path to a company logo image (PNG/JPG) for the HTML report header.
    The image is base64-encoded and embedded directly in the HTML file.
    
    Example: -LogoPath "C:\logos\contoso-logo.png"

.PARAMETER AnalystName
    [Optional] Analyst name displayed on the HTML report cover page.
    
    Example: -AnalystName "John Smith"

.PARAMETER IncludeReservationAnalysis
    [Optional] Analyze Reserved Instance coverage and savings opportunities.
    
    Requires Az.Reservations module and Reservations Reader role at tenant level.
    Compares existing RIs against AVD VM fleet, identifies uncovered VMs,
    calculates potential savings for 1-year and 3-year terms.
    Existing reservations table is filtered to show only AVD-relevant SKUs.
    
    Example: -IncludeReservationAnalysis

.OUTPUTS
    Creates a timestamped folder: Enhanced-AVD-EvidencePack-YYYYMMDD-HHMMSS
    
    RAW DATA (CSV Files):
    ═══════════════════════════════════════════════════════════════════════════════
    AVD-HostPools.csv               Host pool configuration and capacity
    AVD-SessionHosts.csv            Session host details with status
    AVD-VMs.csv                     VM configuration and zone placement
    AVD-ScaleSets.csv               Virtual Machine Scale Set configuration
    AVD-ScaleSet-Instances.csv      VMSS instance details
    AVD-ScalingPlans.csv            Autoscale configuration
    AVD-ScalingPlanAssignments.csv  Scaling plan to host pool assignments
    AVD-ScalingPlanSchedules.csv    Scaling plan schedule details
    AVD-VM-Metrics-Baseline.csv     CPU and memory time-series (unless -SkipAzureMonitorMetrics)
    AVD-LogAnalytics-Results.csv    KQL query results from Log Analytics
    
    ANALYSIS (CSV Files):
    ═══════════════════════════════════════════════════════════════════════════════
    ENHANCED-VM-RightSizing-Recommendations.csv
        ⭐ AVD-aware sizing with SessionsPerVCPU, MemoryPerSessionGB
    ENHANCED-Zone-Resiliency-Analysis.csv
        HA scores and zone distribution per host pool
    ENHANCED-Cost-Analysis.csv
        Monthly/annual savings estimates by host pool
    ENHANCED-SessionHost-Health.csv
        Health status with step-by-step remediation guidance
    ENHANCED-Storage-Optimization.csv
        Disk type findings (Premium-on-Pooled, non-ephemeral, Standard HDD)
    ENHANCED-AccelNet-Analysis.csv
        Accelerated Networking eligibility and gaps
    ENHANCED-Image-Analysis.csv
        Image groups with source, OS info, version, findings
    ENHANCED-Gallery-Image-Versions.csv
        Gallery image age, replication status
    ENHANCED-HostPool-Image-Consistency.csv
        Per-pool image consistency (Consistent / Mixed / Version Drift)
    ENHANCED-Network-Readiness.csv
        RDP Shortpath, NSG, DNS, peering, private endpoint findings
    ENHANCED-Subnet-Analysis.csv
        Subnet capacity, CIDR, IP usage, NSG/UDR status
    ENHANCED-VNet-Analysis.csv
        VNet DNS configuration, peering status
    ENHANCED-SKU-Diversity-Analysis.csv
        SKU concentration risk and alternative SKU recommendations
    ENHANCED-CrossRegion-Analysis.csv
        Gateway-to-host latency paths with distance calculations
    ENHANCED-Disconnect-Reasons.csv
        Disconnect categories (Network/Timeout/Server/Auth/etc.) with counts
    ENHANCED-Disconnects-ByHost.csv
        Per-host disconnect breakdown by category
    ENHANCED-W365-Readiness.csv
        Per-host-pool W365 fit score, recommendation, and cost comparison
    ENHANCED-Actual-Costs.csv (unless -SkipActualCosts)
        Per-VM actual billed costs from Azure Cost Management (last 30 days)
    ENHANCED-HostPool-Costs.csv (unless -SkipActualCosts)
        Per-host-pool actual cost totals with compute vs storage breakdown
    ENHANCED-Security-Posture.csv
        Per-host-pool security score: Trusted Launch, Secure Boot, vTPM, encryption
    ENHANCED-Orphaned-Resources.csv
        Unattached disks, NICs, public IPs in AVD resource groups with est. monthly cost
    ENHANCED-Profile-Health.csv
        Per-session-host P95 profile load times with severity classification
    ENHANCED-UX-Scores.csv
        Per-host-pool composite UX Score: profile load + RTT + disconnects + errors
    ENHANCED-Infrastructure-Costs.csv (unless -SkipActualCosts)
        Networking, storage, AVD service, Log Analytics costs per AVD resource group
    ENHANCED-Executive-Summary.txt
        High-level findings across all analysis areas
    ENHANCED-Summary.json
        Machine-readable summary for automation/dashboards
    README.txt
        Important disclaimers and data summary
    
    OPTIONAL FILES:
    ═══════════════════════════════════════════════════════════════════════════════
    ENHANCED-Analysis-Report.html (if -GenerateHtmlReport)
        17-tab interactive dashboard: Executive Summary, Right-Sizing, Cost,
        Host Health, Network, Storage, Images, Connections, Zone Resiliency,
        W365 Readiness, Security Posture, UX Scores, Orphaned Resources,
        Reservations, Advisor, Incident Analysis. Includes Priority Matrix.
    
    ENHANCED-Azure-Advisor-Recommendations.csv (if -IncludeAzureAdvisor)
        Microsoft Advisor recommendations by category and impact
    
    ENHANCED-Reservation-Analysis.csv (if -IncludeReservationAnalysis)
        RI coverage, over-provisioned RIs, and savings opportunities
    
    ENHANCED-Incident-Comparative-Analysis.csv (if -IncludeIncidentWindowQueries)
        Baseline vs. incident performance comparison
    
    Enhanced-AVD-EvidencePack-{timestamp}.zip (if -CreateZip)
        ZIP archive of all output files

.EXAMPLE
    # Basic analysis - Single subscription
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("abcd1234-1234-1234-1234-123456789abc")
    
    Runtime: ~15 minutes for 500 VMs
    Output: CSV files with AVD-aware recommendations

.EXAMPLE
    # DRY RUN FIRST - Preview before running (RECOMMENDED!)
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1", "sub-2") `
        -DryRun
    
    Runtime: ~30 seconds
    Output: Displays estimated runtime, API calls, output size
    Perfect for: Planning overnight runs, validating parameters

.EXAMPLE
    # Quick daily health check - 2 minutes!
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1") `
        -QuickSummary
    
    Runtime: 2-3 minutes
    Output: Configuration summary with immediate findings
    Perfect for: Daily health checks, incident response

.EXAMPLE
    # Full analysis with HTML report and Azure Advisor
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1", "sub-2") `
        -GenerateHtmlReport `
        -IncludeAzureAdvisor `
        -CreateZip
    
    Runtime: ~20 minutes for 500 VMs
    Output: CSVs + visual HTML report + Advisor recommendations + ZIP archive

.EXAMPLE
    # Resume from failed run - Saves hours!
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1", "sub-2") `
        -ResumeFrom "Enhanced-AVD-EvidencePack-20260211-120000"
    
    Runtime: ~10 minutes (loads checkpoints, skips completed steps)
    Output: Completes the analysis from where it left off
    Perfect for: Network disconnects, throttling errors, power failures

.EXAMPLE
    # Fast run - Configuration only, no metrics
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1") `
        -SkipAzureMonitorMetrics `
        -SkipLogAnalyticsQueries
    
    Runtime: ~5 minutes
    Output: Configuration CSVs, limited right-sizing recommendations
    Perfect for: Zone resiliency checks, configuration audits

.EXAMPLE
    # Extended lookback for capacity planning
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1") `
        -MetricsLookbackDays 30 `
        -GenerateHtmlReport
    
    Runtime: ~30 min for 500 VMs (30-day lookback)
    Output: 30-day trend analysis
    Perfect for: Monthly reviews, capacity planning, seasonal analysis

.EXAMPLE
    # Post-incident analysis
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1") `
        -IncludeIncidentWindowQueries `
        -IncidentWindowStart (Get-Date "2026-02-10 14:00") `
        -IncidentWindowEnd (Get-Date "2026-02-10 16:00") `
        -GenerateHtmlReport
    
    Output: Baseline vs. incident performance comparison
    Perfect for: Root cause analysis, post-mortem reports

.EXAMPLE
    # Automation-friendly - For scheduled tasks
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1", "sub-2") `
        -SkipDisclaimer `
        -GenerateHtmlReport `
        -CreateZip
    
    # Then email the ZIP file to stakeholders
    Perfect for: Weekly automated reports, compliance audits

.EXAMPLE
    # Large environment (1000+ VMs) - Use PowerShell 7 for speed!
    pwsh  # Launch PowerShell 7
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1", "sub-2", "sub-3", "sub-4") `
        -GenerateHtmlReport
    
    Runtime: PowerShell 7+ with parallel processing
    • 1700 VMs: ~45 min
    Perfect for: Enterprise environments

.EXAMPLE
    # RECOMMENDED: Full assessment with all analysis engines
    # Best for: First-time deep dive, quarterly reviews, customer engagements
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1") `
        -LogAnalyticsWorkspaceResourceIds @("/subscriptions/{sub-id}/resourceGroups/{rg}/providers/Microsoft.OperationalInsights/workspaces/{ws}") `
        -MetricsLookbackDays 14 `
        -GenerateHtmlReport `
        -IncludeAzureAdvisor `
        -CreateZip `
        -SkipDisclaimer
    
    Enables: Actual costs (default), security posture, UX scoring, profile health,
    orphaned resources, network readiness, golden images, W365 readiness, disconnect
    analysis, cross-region analysis, SKU diversity — all 17 HTML dashboard tabs.
    Runtime: ~20 min for 500 VMs, ~60 min for 1700 VMs (14-day lookback)
    Output: Full ZIP with CSVs + 17-tab HTML dashboard

.EXAMPLE
    # Anonymized report for customer delivery
    .\Get-Enhanced-AVD-EvidencePack.ps1 `
        -TenantId "12345678-1234-1234-1234-123456789abc" `
        -SubscriptionIds @("sub-1") `
        -GenerateHtmlReport `
        -ScrubPII `
        -CompanyName "Contoso Ltd" `
        -AnalystName "Jane Smith" `
        -IncludeReservationAnalysis `
        -CreateZip
    
    Output: PII-scrubbed HTML + CSVs with branded header. Safe for SharePoint upload.
    All usernames → User-A1B2, VM names → Host-3F7C, IPs → 10.0.x.x format.

.NOTES
    Name: Get-Enhanced-AVD-EvidencePack.ps1
    Version: 4.1.0
    Author: Enhanced AVD Analysis Tool
    Release Date: 2026-02-13
    
    REQUIREMENTS:
    • PowerShell 7.2+ (REQUIRED — PowerShell 5.1 is no longer supported)
    • Az PowerShell modules: Az.Accounts, Az.Compute, Az.DesktopVirtualization, Az.Monitor, Az.ResourceGraph
    • Azure RBAC: Reader role on subscriptions (minimum)
    • Optional: Cost Management Reader for actual billed costs
    • Optional: Az.Reservations + Reservations Reader for RI analysis
    
    PERFORMANCE:
    • Small (100 VMs): ~5 min
    • Medium (500 VMs): ~15 min
    • Large (1700 VMs): ~45 min
    (Metrics collection uses batched API calls — 1 call per VM instead of 12)
    
    TIPS:
    ✅ Always run -DryRun first to preview runtime
    ✅ Use -QuickSummary for daily health checks
    ✅ Use PowerShell 7 for 8-10x faster metrics collection
    ✅ Use -ResumeFrom if a run fails (saves hours!)
    ✅ Use -ScrubPII when sharing reports externally
    ✅ Review HTML report for visual summary, CSVs for detailed analysis
    ✅ Negative "EstimatedMonthlySavings" = Cost increases (upsizing for performance)
    
    AVD-AWARE INTELLIGENCE:
    🎯 Sessions per vCPU (optimal: 4-6) - More CPUs ≠ Better per-user performance!
    🎯 Memory per session (optimal: 4GB) - Critical for user experience
    🎯 Prevents unnecessary upsizing when session density is low
    🎯 Identifies workload issues vs. sizing issues
    🎯 Recommends E-series when memory-per-user is the constraint
    
    CHANGELOG:
    ─────────────────────────────────────────────────────────────────────
    v4.1.0 (2026-02-13)
      ──── Cost Intelligence Overhaul + PII Scrubbing + W365 User-Count Accuracy ────
      NEW: PII Scrubbing (-ScrubPII)
        • Anonymizes usernames, VM names, host pool names, sub IDs, RGs, IPs
        • Consistent per-run hashing (cross-referenceable across tabs/CSVs)
        • Applied to HTML inline (19 touchpoints) and CSV post-processing
        • Orange banner in HTML when PII scrubbing is active
      NEW: Login Time Analysis
        • Per-pool Avg/P50/P95/Max login time from WVDConnections KQL
        • Ratings: Excellent (≤15s) / Good (≤30s) / Fair (≤60s) / Poor (>60s)
      NEW: Connection Success Rate
        • Per-pool success/failure percentage with Healthy/Monitor/Investigate badges
        • Alert banner for pools with >5% failure rate
        • Unique user count per pool from dcount(UserName)
      NEW: Drain Mode Awareness
        • Detects AllowNewSession=false hosts, calculates effective capacity
        • Surfaces in Connection tab and per-user cost calculations
      NEW: Per-User Cost Calculation
        • AVD cost per user per pool (total monthly ÷ unique users)
        • Side-by-side AVD vs W365 per-user verdict
      NEW: W365 Feature Gap Analysis
        • 11-feature comparison table: multi-session, RemoteApp, GPU, autoscale, etc.
        • Environment-aware detection (flags gaps relevant to your pools)
        • Impact ratings: High/Medium/Low
      NEW: Infrastructure Cost Collection
        • Queries networking, storage accounts, AVD service, Log Analytics, Key Vault
        • Scoped to AVD resource groups (won't pull unrelated costs)
        • Infrastructure cost breakdown table in Cost Breakdown tab
        • Full AVD Monthly = Compute + Infrastructure
      IMPROVED: Cost Management — RI/Reservation handling
        • Cost queries now scoped to AVD resource groups (excludes non-AVD VMs)
        • Detects $0 compute on RI-covered VMs, runs RG-level AmortizedCost query
        • RG-level query captures RI amortization that per-ResourceId misses
        • Removed broken PAYG sanity check that incorrectly flagged RI-discounted VMs
        • Rich console diagnostics: match rate, unmatched VMs, subscription gaps
      IMPROVED: W365 cost comparison now uses actual user counts
        • Pooled pools: unique users × W365 price (not VM count)
        • Frontline: concurrent users × Frontline price
        • Personal pools: assigned users × W365 price
        • New "Users" and "W365 Licenses" columns in HTML table
      IMPROVED: W365 plan matching now workload-aware
        • Pooled Desktop: matches on per-user resources (RAM ÷ users), not total VM spec
        • RemoteApp: lighter per-user sizing (2 vCPU / 2-4 GB) → matches Frontline plans
        • Personal: matches on full VM spec (1:1, correct)
        • Usage-based plan shown alongside spec-matched in HTML table
        • Prevents inflated W365 pricing (e.g., $184 for E8ads_v6 → $38 per-user)
      IMPROVED: Per-user cost calculation uses Log Analytics unique users
        • Pooled pools: falls back to LA dcount(UserName) instead of VM count
        • 4-tier priority: Assigned → LA unique users → Active sessions → VM count
        • User count source displayed in HTML table (Assigned / Log Analytics / Active)
        • Fuzzy host pool name matching for KQL data (handles short vs full name)
      IMPROVED: Reservation analysis filtered to AVD-relevant SKUs
        • Existing Reservations table shows only RIs matching AVD VM sizes
        • New "AVD VMs" and "Coverage" columns per reservation
      IMPROVED: Resource ID matching for Cost Management
        • Normalized lowercase comparison + name-from-ID fallback
        • Handles casing differences between Get-AzVM and Cost Management API
      FIXED: GPU detection regex false positives (Standard_D4s_v5 matched as GPU)
      FIXED: CSV Import-Csv StrictMode error on single-row CSVs
      FIXED: $kqlSessionDuration missing here-string close
      FIXED: networkFindings property name (Finding → Detail)
      FIXED: $fitFactors used before initialization in W365 analysis
      FIXED: $avdResourceGroups used before initialization in cost collection
      FIXED: Nested $_ collision in reservation SKU filtering
    ─────────────────────────────────────────────────────────────────────
    v4.0.0 (2026-02-12)
      ──── Major Release: Full AVD Assessment Platform (11 new analysis engines) ────
      NEW: W365 Cloud PC Readiness Analysis
        • Per-host-pool fit scoring (0-100) with Strong/Consider/Keep recommendations
        • Cost comparison: AVD IaaS vs W365 Enterprise plans
        • Hard blocker detection: GPU, multi-session, oversized VMs
      NEW: Actual Cost Intelligence (on by default, -SkipActualCosts to disable)
        • Azure Cost Management API — real billed costs per VM (last 30 days)
        • Graceful permission fallback to PAYG estimates
        • Compute vs storage split; VMSS cost matching (split across instances)
        • Per-host-pool cost breakdown in HTML dashboard
      NEW: Security Posture Scoring
        • Trusted Launch, Secure Boot, vTPM, Host Encryption per host pool
        • Security score (0-100, graded A-F) with hardening guidance
        • Priority Matrix flags pools with grade D/F
      NEW: User Experience Score
        • Composite metric: profile load + RTT + disconnect rate + errors
        • Per-host-pool UX Score (0-100, graded A-F) with sub-component breakdown
      NEW: Orphaned Resource Detection
        • Unattached disks, NICs, public IPs scoped to AVD resource groups
        • Estimated monthly waste per resource with cleanup guidance
      NEW: Profile Health Analysis
        • P95 profile load times surfaced from Log Analytics KQL data
        • Severity classification: Good (<30s) / Warning (30-59s) / Critical (60s+)
      NEW: Disconnect Reason Analysis
        • Categorizes disconnects using real AVD CodeSymbolic values
        • Per-host disconnect breakdown with remediation guidance
      NEW: Network Readiness Analysis
        • RDP Shortpath detection — surfaces UDP vs TCP relay usage from KQL data
        • Subnet capacity analysis — CIDR sizing, IP usage %, available IPs
        • NSG coverage check — flags session hosts with no NSG at NIC or subnet level
        • VNet DNS configuration — detects custom vs Azure default DNS
        • VNet peering health — flags disconnected peerings
        • Private endpoint check — identifies host pools without private endpoints
        • UDR (Route Table) presence tracking per subnet
      NEW: Golden Image Assessment
        • Image source classification (Gallery / Marketplace / ManagedImage / Custom)
        • OS generation and build freshness detection (24H2, 23H2, etc.)
        • Per-host-pool image consistency check (Mixed Sources / Mixed SKUs / Version Drift)
        • Azure Compute Gallery version age check (flags images > 90 days)
        • Gallery replication verification across deployment regions
        • Golden Image Maturity Score (0-100, graded A-F)
        • "Notes from the Field" — context-aware best practice guidance
      NEW: SKU Diversity & Allocation Resilience
        • SKU family/series parsing, concentration risk scoring
        • Alternative SKU recommendations from same workload class
        • Pooled pools flag High risk even with 1 VM
      NEW: Cross-Region Connection Analysis
        • Haversine distance calculation, expected RTT baselines
        • Routing anomaly detection for gateway-to-host paths
      IMPROVED: HTML Dashboard — now 17 tabs (was 11)
        • Security Posture, UX Scores, Orphaned Resources tabs added
        • Cost Breakdown shows actual billing + host pool breakdown
        • Overview cards: UX Score, Security Score, Orphaned Resources
      IMPROVED: Session Host Health — remediation steps for every finding
      IMPROVED: Comprehensive pricing engine covers all VM SKU families
      FIXED: StrictMode .Count errors on VMSS VMs missing network properties
      FIXED: StrictMode null property access on VNet/subnet/NIC objects
      FIXED: Orphaned resource scan scoped to AVD resource groups only
      FIXED: VMSS cost matching (scale set level, split across instances)
      FIXED: Profile health host pool matching with safer fallback chain
    ─────────────────────────────────────────────────────────────────────
    v3.1.0 (2026-02-11)
      NEW: Cross-region connection analysis with geographic latency mapping
      NEW: SKU Diversity & Allocation Resilience analysis
      FIXED: KQL schema corrections for WVDConnectionNetworkData
    ─────────────────────────────────────────────────────────────────────
    v3.0.0 (2026-02-10)
      NEW: Comprehensive HTML dashboard with 11 analysis tabs
      NEW: AVD-aware right-sizing with sessions-per-vCPU analysis
      NEW: Zone resiliency scoring per host pool
      NEW: Session host health monitoring (drain mode, heartbeat, availability)
      NEW: Storage optimization (ephemeral disks, premium-on-pooled)
      NEW: Accelerated Networking gap detection
      NEW: Image staleness and version drift detection
      NEW: Priority Matrix with Quick Wins / Plan / Consider quadrants
      NEW: Cost analysis with monthly/annual savings estimates
      NEW: Reservation analysis (when -IncludeReservationAnalysis specified)
    ─────────────────────────────────────────────────────────────────────
    
    SUPPORT:
    • Documentation: See README.md and docs/ folder

.LINK
    Microsoft AVD Best Practices: https://learn.microsoft.com/azure/virtual-desktop/

.LINK
    Azure VM Sizing: https://learn.microsoft.com/azure/virtual-machines/sizes

#>
[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string]$TenantId,

  [Parameter(Mandatory)]
  [string[]]$SubscriptionIds,

  [string[]]$LogAnalyticsWorkspaceResourceIds = @(),

  [switch]$SkipAzureMonitorMetrics,
  [switch]$SkipLogAnalyticsQueries,

  [int]$MetricsLookbackDays = 7,
  [int]$MetricsTimeGrainMinutes = 15,

  [switch]$IncludeIncidentWindowQueries,
  [datetime]$IncidentWindowStart = (Get-Date).AddDays(-14),
  [datetime]$IncidentWindowEnd   = (Get-Date),

  [switch]$CreateZip,
  
  # Enhanced parameters
  [int]$CpuRightSizingThreshold = 40,        # VMs below this avg CPU are candidates for downsizing
  [int]$CpuOverloadThreshold = 80,           # VMs above this peak CPU need upsizing
  [int]$MemoryLowThreshold = 1073741824,     # 1 GB - VMs below this free memory need upsizing
  [int]$MinimumZonesForResiliency = 2,       # Minimum zones for HA recommendation
  [decimal]$SessionDensityTarget = 0.7,      # Target 70% session density for optimal utilization
  
  # New features (v2.1)
  [switch]$SkipDisclaimer,                   # Skip interactive disclaimer for automation
  [switch]$GenerateHtmlReport,               # Generate HTML report in addition to CSV
  [switch]$IncludeAzureAdvisor,              # Include Azure Advisor recommendations
  
  # Phase 1 features (v2.2)
  [switch]$DryRun,                           # Show what would be collected without collecting
  [switch]$QuickSummary,                     # Fast health check without metrics collection
  [string]$ResumeFrom = "",                  # Resume from a previous incomplete run
  
  # Branding (v3.0)
  [string]$CompanyName = "",                 # Company name for branded HTML report header
  [string]$LogoPath = "",                    # Path to logo image (PNG/JPG) for HTML report header
  [string]$AnalystName = "",                 # Analyst name shown on report cover
  
  # Reservation Analysis (v3.0)
  [switch]$IncludeReservationAnalysis,       # Analyze RI coverage and savings opportunities
  
  # Cost Intelligence (v3.2)
  [switch]$SkipActualCosts,                  # Skip Azure Cost Management API query (use PAYG estimates only)
  
  # Privacy (v4.1)
  [switch]$ScrubPII                          # Anonymize usernames, VM names, IPs, and subscription IDs in HTML report and CSVs
)

$WarningPreference = 'SilentlyContinue'
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# =========================================================
# PowerShell 7 Requirement
# =========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
  Write-Host ""
  Write-Host "ERROR: PowerShell 7.2+ is required." -ForegroundColor Red
  Write-Host ""
  Write-Host "You are running PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "Install PowerShell 7:" -ForegroundColor Cyan
  Write-Host "  winget install Microsoft.PowerShell" -ForegroundColor White
  Write-Host "  or: https://aka.ms/powershell-release?tag=stable" -ForegroundColor White
  Write-Host ""
  Write-Host "Then run this script from pwsh.exe (not powershell.exe)" -ForegroundColor Cyan
  exit 1
}

# =========================================================
# Prerequisite Validation
# =========================================================
Write-Host "Validating prerequisites..." -ForegroundColor Cyan

$requiredModules = @(
    @{Name='Az.Accounts'; MinVersion='2.0.0'},
    @{Name='Az.Compute'; MinVersion='4.0.0'},
    @{Name='Az.DesktopVirtualization'; MinVersion='2.0.0'},
    @{Name='Az.Monitor'; MinVersion='2.0.0'},
    @{Name='Az.OperationalInsights'; MinVersion='2.0.0'},
    @{Name='Az.Resources'; MinVersion='4.0.0'}
)

$missingModules = @()
foreach ($module in $requiredModules) {
    $installed = Get-Module -ListAvailable -Name $module.Name | 
        Where-Object { $_.Version -ge [version]$module.MinVersion } |
        Select-Object -First 1
    
    if (-not $installed) {
        $missingModules += $module.Name
        Write-Host "  ✗ Missing: $($module.Name) (>= $($module.MinVersion))" -ForegroundColor Red
    } else {
        Write-Host "  ✓ Found: $($module.Name) v$($installed.Version)" -ForegroundColor Green
    }
}

if (@($missingModules).Count -gt 0) {
    Write-Host "`nERROR: Missing required modules" -ForegroundColor Red
    Write-Host "Install with: Install-Module -Name $($missingModules -join ', ')" -ForegroundColor Yellow
    exit 1
}

Write-Host "✓ All prerequisites validated`n" -ForegroundColor Green

# Optional modules
$hasAzReservations = $false
if ($IncludeReservationAnalysis) {
  $azResModule = Get-Module -ListAvailable -Name 'Az.Reservations' | Select-Object -First 1
  if ($azResModule) {
    $hasAzReservations = $true
    Write-Host "  ✓ Optional: Az.Reservations v$($azResModule.Version)" -ForegroundColor Green
  } else {
    Write-Host "  ⚠ Optional: Az.Reservations not installed — will analyze RI opportunities but cannot check existing reservations" -ForegroundColor Yellow
    Write-Host "    Install with: Install-Module -Name Az.Reservations" -ForegroundColor Gray
  }
}

# =========================================================
# Resume Mode Check
# =========================================================
$resumeData = $null
if ($ResumeFrom) {
  Write-Host ""
  Write-Host "╔═══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
  Write-Host "║                         RESUME MODE                                   ║" -ForegroundColor Cyan
  Write-Host "╚═══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
  Write-Host ""
  
  if (-not (Test-Path $ResumeFrom)) {
    Write-Host "✗ ERROR: Resume folder not found: $ResumeFrom" -ForegroundColor Red
    Write-Host ""
    Write-Host "Make sure the path is correct and try again." -ForegroundColor Yellow
    exit 1
  }
  
  Write-Host "Checking resume folder: $ResumeFrom" -ForegroundColor Cyan
  
  # Check for checkpoint file
  $checkpointFile = Join-Path $ResumeFrom ".checkpoint.json"
  if (Test-Path $checkpointFile) {
    try {
      $resumeData = Get-Content $checkpointFile -Raw | ConvertFrom-Json
      Write-Host "✓ Found checkpoint from: $($resumeData.Timestamp)" -ForegroundColor Green
      Write-Host "  Last completed step: $($resumeData.LastCompletedStep)" -ForegroundColor Gray
      Write-Host ""
    }
    catch {
      Write-Host "⚠ Could not read checkpoint file, will attempt partial resume" -ForegroundColor Yellow
      Write-Host ""
    }
  }
  else {
    Write-Host "⚠ No checkpoint file found, will attempt to resume from existing CSVs" -ForegroundColor Yellow
    Write-Host ""
  }
}

# =========================================================
# Helpers
# =========================================================

# --- PII Scrubbing (v4.1) ---
# Consistent hashing: same input always produces same anonymous ID within a run
$script:piiSalt = [guid]::NewGuid().ToString().Substring(0, 8)
$script:piiCache = @{}

function Scrub-Value {
  param(
    [string]$Value,
    [string]$Prefix = "Anon",
    [int]$Length = 4
  )
  if (-not $ScrubPII) { return $Value }
  if ([string]::IsNullOrEmpty($Value)) { return $Value }
  $key = "${Prefix}:${Value}"
  if ($script:piiCache.ContainsKey($key)) { return $script:piiCache[$key] }
  $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash(
    [System.Text.Encoding]::UTF8.GetBytes("${Value}:${script:piiSalt}")
  )
  $short = [BitConverter]::ToString($hash[0..($Length/2)]).Replace('-','').Substring(0, $Length).ToUpper()
  $result = "${Prefix}-${short}"
  $script:piiCache[$key] = $result
  return $result
}

function Scrub-Username {
  param([string]$Value)
  if (-not $ScrubPII) { return $Value }
  if ([string]::IsNullOrEmpty($Value)) { return $Value }
  return Scrub-Value -Value $Value -Prefix "User" -Length 4
}

function Scrub-VMName {
  param([string]$Value)
  if (-not $ScrubPII) { return $Value }
  if ([string]::IsNullOrEmpty($Value)) { return $Value }
  return Scrub-Value -Value $Value -Prefix "Host" -Length 6
}

function Scrub-HostPoolName {
  param([string]$Value)
  if (-not $ScrubPII) { return $Value }
  if ([string]::IsNullOrEmpty($Value)) { return $Value }
  return Scrub-Value -Value $Value -Prefix "Pool" -Length 4
}

function Scrub-SubscriptionId {
  param([string]$Value)
  if (-not $ScrubPII) { return $Value }
  if ([string]::IsNullOrEmpty($Value)) { return $Value }
  if ($Value.Length -ge 4) { return "****-****-****-" + $Value.Substring($Value.Length - 4) }
  return "****"
}

function Scrub-ResourceGroup {
  param([string]$Value)
  if (-not $ScrubPII) { return $Value }
  if ([string]::IsNullOrEmpty($Value)) { return $Value }
  return Scrub-Value -Value $Value -Prefix "RG" -Length 4
}

function Scrub-IP {
  param([string]$Value)
  if (-not $ScrubPII) { return $Value }
  if ([string]::IsNullOrEmpty($Value)) { return $Value }
  # Keep subnet, mask host: 10.0.1.45 → 10.0.1.x
  if ($Value -match '^(\d+\.\d+\.\d+)\.\d+$') { return "$($matches[1]).x" }
  return "x.x.x.x"
}

function SafeArray {
  param([object]$Obj)
  if ($null -eq $Obj) { return @() }
  return @($Obj)
}

function SafeCount {
  param([object]$Obj)
  if ($null -eq $Obj) { return 0 }
  return @($Obj).Count
}

function SafeProp {
  param([object]$Obj, [string]$Name)
  if ($null -eq $Obj) { return $null }
  if ($Obj.PSObject.Properties.Name -contains $Name) { return $Obj.$Name }
  return $null
}

function SafeMeasure {
  param([object]$MeasureResult, [string]$Property)
  if ($null -eq $MeasureResult) { return $null }
  if ($MeasureResult.PSObject.Properties.Name -contains $Property) { return $MeasureResult.$Property }
  return $null
}

function TryGet-ArmId {
  param([object]$Obj)
  if (-not $Obj) { return $null }
  foreach ($p in @("Id","ResourceId","ArmId")) {
    if ($Obj.PSObject.Properties.Name -contains $p) {
      $v = $Obj.$p
      if ($v -is [string] -and $v.Trim().Length -gt 0) { return $v }
    }
  }
  return $null
}

function Get-NameFromArmId {
  param([string]$ArmId)
  if (-not $ArmId) { return $null }
  return ($ArmId -split "/")[-1]
}

function Get-RgFromArmId {
  param([string]$ArmId)
  if (-not $ArmId) { return $null }
  if ($ArmId -match "/resourceGroups/([^/]+)/") { return $matches[1] }
  return $null
}

function Get-SubFromArmId {
  param([string]$ArmId)
  if (-not $ArmId) { return $null }
  if ($ArmId -match "/subscriptions/([^/]+)/") { return $matches[1] }
  return $null
}

function Invoke-WithRetry {
  param(
    [Parameter(Mandatory)]
    [scriptblock]$ScriptBlock,
    
    [int]$MaxRetries = 3,
    [int]$BaseDelaySeconds = 30,
    [string]$OperationName = "Operation"
  )
  
  $attempt = 0
  while ($attempt -lt $MaxRetries) {
    try {
      return & $ScriptBlock
    }
    catch {
      $attempt++
      $errorMessage = $_.Exception.Message
      
      # Check if it's a throttling error
      $isThrottling = $errorMessage -match "Too Many Requests" -or 
                      $errorMessage -match "429" -or
                      $errorMessage -match "throttl" -or
                      $errorMessage -match "rate limit"
      
      if ($isThrottling -and $attempt -lt $MaxRetries) {
        $delay = $BaseDelaySeconds * $attempt
        Write-Host "  ⚠ Throttled - retrying in $delay seconds (attempt $attempt of $MaxRetries)..." -ForegroundColor Yellow
        Start-Sleep -Seconds $delay
      }
      else {
        # Not throttling or max retries reached - throw the error
        throw
      }
    }
  }
}

function Write-ProgressSection {
  param(
    [Parameter(Mandatory)]
    [string]$Section,
    
    [Parameter(Mandatory)]
    [ValidateSet('Start', 'Progress', 'Complete', 'Skip', 'Error')]
    [string]$Status,
    
    [string]$Message = "",
    [int]$Current = 0,
    [int]$Total = 0,
    [int]$EstimatedMinutes = 0
  )
  
  $timestamp = Get-Date -Format "HH:mm:ss"
  
  switch ($Status) {
    'Start' {
      Write-Host ""
      Write-Host "[$timestamp] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
      Write-Host "[$timestamp] 🔄 $Section" -ForegroundColor Cyan
      if ($EstimatedMinutes -gt 0) {
        Write-Host "[$timestamp]    Estimated time: ~$EstimatedMinutes minutes" -ForegroundColor Gray
      }
      if ($Message) {
        Write-Host "[$timestamp]    $Message" -ForegroundColor Gray
      }
      Write-Host "[$timestamp] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    }
    'Progress' {
      if ($Total -gt 0) {
        $percent = [math]::Round(($Current / $Total) * 100, 1)
        $bar = "█" * [math]::Floor($percent / 5) + "░" * (20 - [math]::Floor($percent / 5))
        Write-Host "[$timestamp]    Progress: $bar $percent% ($Current / $Total)" -ForegroundColor Yellow
        if ($Message) {
          Write-Host "[$timestamp]    $Message" -ForegroundColor Gray
        }
      }
      else {
        Write-Host "[$timestamp]    $Message" -ForegroundColor Gray
      }
    }
    'Complete' {
      Write-Host "[$timestamp] ✓ $Section - Complete!" -ForegroundColor Green
      if ($Message) {
        Write-Host "[$timestamp]    $Message" -ForegroundColor Gray
      }
    }
    'Skip' {
      Write-Host "[$timestamp] ⊘ $Section - Skipped" -ForegroundColor Yellow
      if ($Message) {
        Write-Host "[$timestamp]    $Message" -ForegroundColor Gray
      }
    }
    'Error' {
      Write-Host "[$timestamp] ✗ $Section - Error" -ForegroundColor Red
      if ($Message) {
        Write-Host "[$timestamp]    $Message" -ForegroundColor Red
      }
    }
  }
}

# =========================================================
# Phase 1 Feature Helper Functions (v2.2)
# =========================================================
function Save-Checkpoint {
  param(
    [Parameter(Mandatory)]
    [string]$CheckpointName,
    
    [Parameter(Mandatory)]
    [string]$OutputFolder,
    
    [Parameter(Mandatory)]
    [hashtable]$Data
  )
  
  $checkpointPath = Join-Path $OutputFolder ".checkpoints"
  if (-not (Test-Path $checkpointPath)) {
    New-Item -ItemType Directory -Path $checkpointPath -Force | Out-Null
  }
  
  $checkpointFile = Join-Path $checkpointPath "$CheckpointName.json"
  $Data | ConvertTo-Json -Depth 10 | Out-File $checkpointFile -Encoding utf8
  Write-Host "  💾 Checkpoint saved: $CheckpointName" -ForegroundColor Gray
}

function Test-Checkpoint {
  param(
    [Parameter(Mandatory)]
    [string]$CheckpointName,
    
    [Parameter(Mandatory)]
    [string]$OutputFolder
  )
  
  $checkpointFile = Join-Path $OutputFolder ".checkpoints\$CheckpointName.json"
  return (Test-Path $checkpointFile)
}

function Load-Checkpoint {
  param(
    [Parameter(Mandatory)]
    [string]$CheckpointName,
    
    [Parameter(Mandatory)]
    [string]$OutputFolder
  )
  
  $checkpointFile = Join-Path $OutputFolder ".checkpoints\$CheckpointName.json"
  if (Test-Path $checkpointFile) {
    $json = Get-Content $checkpointFile -Raw | ConvertFrom-Json
    Write-Host "  ♻️  Loaded checkpoint: $CheckpointName" -ForegroundColor Green
    return $json
  }
  return $null
}

function Resolve-ArmIdentity {
  param(
    [object]$Obj,
    [string]$FallbackResourceType
  )

  $id = TryGet-ArmId $Obj

  if (-not $id -and $FallbackResourceType) {
    $name = SafeProp $Obj "Name"
    if ($name) {
      $res = Get-AzResource -ResourceType $FallbackResourceType -Name $name -ErrorAction SilentlyContinue | Select-Object -First 1
      if ($res) { $id = $res.ResourceId }
    }
  }

  if (-not $id) { return $null }

  [PSCustomObject]@{
    Id            = $id
    Name          = Get-NameFromArmId $id
    ResourceGroup = Get-RgFromArmId $id
    SubscriptionId= Get-SubFromArmId $id
  }
}

function Normalize-SessionHostToVmName {
  param([string]$SessionHostNameField)

  if (-not $SessionHostNameField) { return $null }

  $leaf = ($SessionHostNameField -split "/")[-1]
  if (-not $leaf) { return $null }

  if ($leaf -match "^(?<short>[^\.]+)\.") { return $matches["short"] }
  return $leaf
}

function Get-ArmIdSafe {
    param([object]$Obj)
    if (-not $Obj) { return $null }
    foreach ($p in @("Id","ResourceId","ArmId")) {
        if ($Obj.PSObject.Properties.Name -contains $p) {
            $v = $Obj.$p
            if ($v) { return $v }
        }
    }
    return $null
}

# =========================================================
# Enhanced Pricing Helper (Estimated)
# =========================================================
function Get-EstimatedVmCostPerHour {
    param([string]$VmSize, [string]$Region)
    
    # 1. Check the RI pricing table first — it has accurate PAYG rates for 40+ SKUs
    if ($riPricingTable -and $riPricingTable.ContainsKey($VmSize)) {
        return $riPricingTable[$VmSize].PAYG
    }
    
    # 2. Fallback: Calculate from family + core count using per-vCPU rates
    $perVcpuRates = @{
        "D" = 0.048; "E" = 0.063; "F" = 0.040; "B" = 0.021
        "L" = 0.062; "M" = 0.133; "N" = 0.180; "H" = 0.090; "A" = 0.050
    }
    
    # Inline SKU parser (avoids dependency on Get-SkuFamilyInfo which may not be defined yet)
    $normalized = $VmSize -replace '^Standard_', '' -replace '^Basic_', ''
    if ($normalized -match '^([A-Z]+)(\d+)([a-z]*)_?(v\d+)?$') {
        $family = $Matches[1]
        $cores = [int]$Matches[2]
        $suffix = $Matches[3]
        
        $baseRate = if ($perVcpuRates.ContainsKey($family)) { $perVcpuRates[$family] } else { 0.048 }
        $hourly = $baseRate * $cores
        
        # AMD variants (suffix contains 'a') are ~15% cheaper
        if ($suffix -match 'a') { $hourly *= 0.85 }
        
        return [math]::Round($hourly, 4)
    }
    
    return $null
}

# =========================================================
# Reserved Instance Pricing (Estimated - East US baseline)
# =========================================================
# Format: SKU = @{ PAYG = hourly; RI1Y = hourly-equivalent; RI3Y = hourly-equivalent }
# RI hourly-equivalent = (annual cost / 8760 hours)
# Source: Azure retail pricing, rounded. Actual rates vary by region and EA/CSP.
$riPricingTable = @{
    # D-series v3
    "Standard_D2s_v3"   = @{ PAYG = 0.096;  RI1Y = 0.060;  RI3Y = 0.038 }
    "Standard_D4s_v3"   = @{ PAYG = 0.192;  RI1Y = 0.121;  RI3Y = 0.077 }
    "Standard_D8s_v3"   = @{ PAYG = 0.384;  RI1Y = 0.242;  RI3Y = 0.154 }
    "Standard_D16s_v3"  = @{ PAYG = 0.768;  RI1Y = 0.483;  RI3Y = 0.307 }
    # D-series v4
    "Standard_D2s_v4"   = @{ PAYG = 0.096;  RI1Y = 0.060;  RI3Y = 0.038 }
    "Standard_D4s_v4"   = @{ PAYG = 0.192;  RI1Y = 0.121;  RI3Y = 0.077 }
    "Standard_D8s_v4"   = @{ PAYG = 0.384;  RI1Y = 0.242;  RI3Y = 0.154 }
    "Standard_D16s_v4"  = @{ PAYG = 0.768;  RI1Y = 0.483;  RI3Y = 0.307 }
    # D-series v5
    "Standard_D2s_v5"   = @{ PAYG = 0.096;  RI1Y = 0.060;  RI3Y = 0.039 }
    "Standard_D4s_v5"   = @{ PAYG = 0.192;  RI1Y = 0.121;  RI3Y = 0.078 }
    "Standard_D8s_v5"   = @{ PAYG = 0.384;  RI1Y = 0.242;  RI3Y = 0.155 }
    "Standard_D16s_v5"  = @{ PAYG = 0.768;  RI1Y = 0.483;  RI3Y = 0.310 }
    # D-series v5 AMD
    "Standard_D2ads_v5"  = @{ PAYG = 0.082;  RI1Y = 0.052;  RI3Y = 0.033 }
    "Standard_D4ads_v5"  = @{ PAYG = 0.164;  RI1Y = 0.103;  RI3Y = 0.066 }
    "Standard_D8ads_v5"  = @{ PAYG = 0.329;  RI1Y = 0.207;  RI3Y = 0.132 }
    "Standard_D16ads_v5" = @{ PAYG = 0.658;  RI1Y = 0.414;  RI3Y = 0.263 }
    # D-series v6
    "Standard_D2s_v6"   = @{ PAYG = 0.096;  RI1Y = 0.061;  RI3Y = 0.039 }
    "Standard_D4s_v6"   = @{ PAYG = 0.192;  RI1Y = 0.121;  RI3Y = 0.078 }
    "Standard_D8s_v6"   = @{ PAYG = 0.384;  RI1Y = 0.242;  RI3Y = 0.156 }
    "Standard_D16s_v6"  = @{ PAYG = 0.768;  RI1Y = 0.484;  RI3Y = 0.311 }
    # E-series v3
    "Standard_E2s_v3"   = @{ PAYG = 0.126;  RI1Y = 0.080;  RI3Y = 0.051 }
    "Standard_E4s_v3"   = @{ PAYG = 0.252;  RI1Y = 0.159;  RI3Y = 0.101 }
    "Standard_E8s_v3"   = @{ PAYG = 0.504;  RI1Y = 0.318;  RI3Y = 0.202 }
    "Standard_E16s_v3"  = @{ PAYG = 1.008;  RI1Y = 0.635;  RI3Y = 0.403 }
    # E-series v4
    "Standard_E2s_v4"   = @{ PAYG = 0.126;  RI1Y = 0.080;  RI3Y = 0.051 }
    "Standard_E4s_v4"   = @{ PAYG = 0.252;  RI1Y = 0.159;  RI3Y = 0.101 }
    "Standard_E8s_v4"   = @{ PAYG = 0.504;  RI1Y = 0.318;  RI3Y = 0.202 }
    "Standard_E16s_v4"  = @{ PAYG = 0.768;  RI1Y = 0.484;  RI3Y = 0.307 }
    # E-series v5
    "Standard_E2s_v5"   = @{ PAYG = 0.126;  RI1Y = 0.080;  RI3Y = 0.051 }
    "Standard_E4s_v5"   = @{ PAYG = 0.252;  RI1Y = 0.159;  RI3Y = 0.101 }
    "Standard_E8s_v5"   = @{ PAYG = 0.504;  RI1Y = 0.318;  RI3Y = 0.203 }
    "Standard_E16s_v5"  = @{ PAYG = 1.008;  RI1Y = 0.635;  RI3Y = 0.404 }
    # E-series v5 AMD
    "Standard_E2ads_v5"  = @{ PAYG = 0.110;  RI1Y = 0.069;  RI3Y = 0.044 }
    "Standard_E4ads_v5"  = @{ PAYG = 0.220;  RI1Y = 0.139;  RI3Y = 0.088 }
    "Standard_E8ads_v5"  = @{ PAYG = 0.440;  RI1Y = 0.277;  RI3Y = 0.176 }
    "Standard_E16ads_v5" = @{ PAYG = 0.880;  RI1Y = 0.554;  RI3Y = 0.352 }
    # E-series v6 AMD (customer's VMs)
    "Standard_E2ads_v6"  = @{ PAYG = 0.110;  RI1Y = 0.070;  RI3Y = 0.044 }
    "Standard_E4ads_v6"  = @{ PAYG = 0.220;  RI1Y = 0.139;  RI3Y = 0.089 }
    "Standard_E8ads_v6"  = @{ PAYG = 0.440;  RI1Y = 0.278;  RI3Y = 0.177 }
    "Standard_E16ads_v6" = @{ PAYG = 0.880;  RI1Y = 0.555;  RI3Y = 0.353 }
    # B-series (burstable)
    "Standard_B2s"    = @{ PAYG = 0.042;  RI1Y = 0.026;  RI3Y = 0.017 }
    "Standard_B4ms"   = @{ PAYG = 0.166;  RI1Y = 0.105;  RI3Y = 0.067 }
    "Standard_B8ms"   = @{ PAYG = 0.333;  RI1Y = 0.210;  RI3Y = 0.133 }
    "Standard_B16ms"  = @{ PAYG = 0.666;  RI1Y = 0.420;  RI3Y = 0.266 }
}

function Get-RightSizedVmRecommendation {
    param(
        [string]$CurrentSize,
        [double]$AvgCpu,
        [double]$PeakCpu,
        [double]$AvgMemoryUsedGB,
        [double]$PeakMemoryUsedGB,
        [int]$AvgSessions,
        [int]$PeakSessions,
        [int]$LookbackDays = 0,
        [double]$AvgFreeMem = 0,
        [double]$MinFreeMem = 0,
        [string]$HostPoolType = "",
        [string]$AppGroupType = "",
        [string]$LoadBalancer = ""
    )
    
    # VM Size families with vCPU and Memory
    # Includes: v4, v5, v6 series and variants (s, as, ads, d, ds)
    $vmSizes = @{
        # D-series v4
        "Standard_D2s_v4"  = @{ vCPU = 2;  MemoryGB = 8  }
        "Standard_D4s_v4"  = @{ vCPU = 4;  MemoryGB = 16 }
        "Standard_D8s_v4"  = @{ vCPU = 8;  MemoryGB = 32 }
        "Standard_D16s_v4" = @{ vCPU = 16; MemoryGB = 64 }
        "Standard_D32s_v4" = @{ vCPU = 32; MemoryGB = 128 }
        
        # D-series v5
        "Standard_D2s_v5"  = @{ vCPU = 2;  MemoryGB = 8  }
        "Standard_D4s_v5"  = @{ vCPU = 4;  MemoryGB = 16 }
        "Standard_D8s_v5"  = @{ vCPU = 8;  MemoryGB = 32 }
        "Standard_D16s_v5" = @{ vCPU = 16; MemoryGB = 64 }
        "Standard_D32s_v5" = @{ vCPU = 32; MemoryGB = 128 }
        "Standard_D48s_v5" = @{ vCPU = 48; MemoryGB = 192 }
        "Standard_D64s_v5" = @{ vCPU = 64; MemoryGB = 256 }
        
        # D-series v5 AMD (ads = AMD + SSD)
        "Standard_D2ads_v5"  = @{ vCPU = 2;  MemoryGB = 8  }
        "Standard_D4ads_v5"  = @{ vCPU = 4;  MemoryGB = 16 }
        "Standard_D8ads_v5"  = @{ vCPU = 8;  MemoryGB = 32 }
        "Standard_D16ads_v5" = @{ vCPU = 16; MemoryGB = 64 }
        "Standard_D32ads_v5" = @{ vCPU = 32; MemoryGB = 128 }
        
        # D-series v6 (latest generation)
        "Standard_D2s_v6"  = @{ vCPU = 2;  MemoryGB = 8  }
        "Standard_D4s_v6"  = @{ vCPU = 4;  MemoryGB = 16 }
        "Standard_D8s_v6"  = @{ vCPU = 8;  MemoryGB = 32 }
        "Standard_D16s_v6" = @{ vCPU = 16; MemoryGB = 64 }
        "Standard_D32s_v6" = @{ vCPU = 32; MemoryGB = 128 }
        
        # E-series v4
        "Standard_E2s_v4"  = @{ vCPU = 2;  MemoryGB = 16 }
        "Standard_E4s_v4"  = @{ vCPU = 4;  MemoryGB = 32 }
        "Standard_E8s_v4"  = @{ vCPU = 8;  MemoryGB = 64 }
        "Standard_E16s_v4" = @{ vCPU = 16; MemoryGB = 128 }
        "Standard_E32s_v4" = @{ vCPU = 32; MemoryGB = 256 }
        
        # E-series v5
        "Standard_E2s_v5"  = @{ vCPU = 2;  MemoryGB = 16 }
        "Standard_E4s_v5"  = @{ vCPU = 4;  MemoryGB = 32 }
        "Standard_E8s_v5"  = @{ vCPU = 8;  MemoryGB = 64 }
        "Standard_E16s_v5" = @{ vCPU = 16; MemoryGB = 128 }
        "Standard_E32s_v5" = @{ vCPU = 32; MemoryGB = 256 }
        "Standard_E48s_v5" = @{ vCPU = 48; MemoryGB = 384 }
        "Standard_E64s_v5" = @{ vCPU = 64; MemoryGB = 512 }
        
        # E-series v5 AMD (ads = AMD + SSD)
        "Standard_E2ads_v5"  = @{ vCPU = 2;  MemoryGB = 16 }
        "Standard_E4ads_v5"  = @{ vCPU = 4;  MemoryGB = 32 }
        "Standard_E8ads_v5"  = @{ vCPU = 8;  MemoryGB = 64 }
        "Standard_E16ads_v5" = @{ vCPU = 16; MemoryGB = 128 }
        "Standard_E32ads_v5" = @{ vCPU = 32; MemoryGB = 256 }
        
        # E-series v6 (latest generation)
        "Standard_E2s_v6"  = @{ vCPU = 2;  MemoryGB = 16 }
        "Standard_E4s_v6"  = @{ vCPU = 4;  MemoryGB = 32 }
        "Standard_E8s_v6"  = @{ vCPU = 8;  MemoryGB = 64 }
        "Standard_E16s_v6" = @{ vCPU = 16; MemoryGB = 128 }
        "Standard_E32s_v6" = @{ vCPU = 32; MemoryGB = 256 }
        
        # E-series v6 AMD (ads = AMD + SSD) - Customer's VMs!
        "Standard_E2ads_v6"  = @{ vCPU = 2;  MemoryGB = 16 }
        "Standard_E4ads_v6"  = @{ vCPU = 4;  MemoryGB = 32 }
        "Standard_E8ads_v6"  = @{ vCPU = 8;  MemoryGB = 64 }
        "Standard_E16ads_v6" = @{ vCPU = 16; MemoryGB = 128 }
        "Standard_E32ads_v6" = @{ vCPU = 32; MemoryGB = 256 }
        "Standard_E48ads_v6" = @{ vCPU = 48; MemoryGB = 384 }
        "Standard_E64ads_v6" = @{ vCPU = 64; MemoryGB = 512 }
        
        # B-series (Burstable - common for small AVD deployments)
        "Standard_B2s"   = @{ vCPU = 2;  MemoryGB = 4  }
        "Standard_B4ms"  = @{ vCPU = 4;  MemoryGB = 16 }
        "Standard_B8ms"  = @{ vCPU = 8;  MemoryGB = 32 }
        "Standard_B12ms" = @{ vCPU = 12; MemoryGB = 48 }
        "Standard_B16ms" = @{ vCPU = 16; MemoryGB = 64 }
    }
    
    $currentSpecs = $vmSizes[$CurrentSize]
    if (-not $currentSpecs) {
        return [PSCustomObject]@{
            Recommendation = "Unknown"
            Reason = "Current VM size not in recommendation database"
            Confidence = "Low"
            EvidenceScore = 0
            EvidenceSignals = "Unknown-SKU"
            CurrentvCPU = $null
            CurrentMemoryGB = $null
            SessionsPerVCPU = $null
            MemoryPerSessionGB = $null
        }
    }
    
    $recommendation = $null
    $reason = @()
    
    # ========================================================
    # EVIDENCE SCORE (v3.0.0)
    # ========================================================
    # Score 0-100 based on what data actually informed the recommendation.
    # Each signal gets points based on its diagnostic value for right-sizing.
    #
    # Signal                     Points  Why it matters
    # -------------------------  ------  ------------------------------------------
    # CPU metrics available        20    Base utilization signal
    # Memory metrics available     15    Catches memory-bound workloads
    # Session data available       30    Critical for AVD — CPU alone is misleading
    # Lookback >= 7 days           15    Filters out single-day anomalies
    # Avg + Peak CPU divergence    10    High divergence = bursty; needs peak headroom
    # Known VM SKU                 10    Required for family-aware recommendations
    #
    # Label mapping:
    #   85-100  →  High      (CPU + memory + sessions + sufficient history)
    #   50-84   →  Medium    (missing sessions or short lookback)
    #   0-49    →  Low       (missing most signals; recommendation is a guess)
    
    $evidenceScore = 0
    $evidenceSignals = @()
    
    # CPU metrics
    if ($AvgCpu -gt 0 -or $PeakCpu -gt 0) {
      $evidenceScore += 20
      $evidenceSignals += "CPU"
    }
    
    # Memory metrics
    if ($AvgFreeMem -gt 0 -or $MinFreeMem -gt 0) {
      $evidenceScore += 15
      $evidenceSignals += "Memory"
    }
    
    # Session data (most valuable for AVD)
    if ($PeakSessions -gt 0) {
      $evidenceScore += 30
      $evidenceSignals += "Sessions"
    }
    
    # Lookback period coverage (caller passes lookback days via metric timestamps)
    # If we have CPU data, assume we have the configured lookback
    if (($AvgCpu -gt 0 -or $PeakCpu -gt 0) -and $LookbackDays -ge 7) {
      $evidenceScore += 15
      $evidenceSignals += "${LookbackDays}d-lookback"
    }
    
    # CPU burstiness detection (avg vs peak divergence > 30% = bursty workload)
    if ($PeakCpu -gt 0 -and $AvgCpu -gt 0 -and ($PeakCpu - $AvgCpu) -gt 30) {
      $evidenceScore += 10
      $evidenceSignals += "Burst-pattern"
    }
    elseif ($PeakCpu -gt 0 -and $AvgCpu -gt 0) {
      $evidenceScore += 10
      $evidenceSignals += "Steady-pattern"
    }
    
    # Known SKU
    if ($currentSpecs) {
      $evidenceScore += 10
      $evidenceSignals += "Known-SKU"
    }
    
    $confidence = if ($evidenceScore -ge 85) { "High" }
                  elseif ($evidenceScore -ge 50) { "Medium" }
                  else { "Low" }
    
    # ========================================================
    # AVD-SPECIFIC METRICS
    # ========================================================
    
    # Calculate session density (key AVD metric!)
    $sessionsPerVCpu = if ($PeakSessions -gt 0 -and $currentSpecs.vCPU -gt 0) {
        [math]::Round($PeakSessions / $currentSpecs.vCPU, 2)
    } else { 0 }
    
    # Calculate memory per session (critical for user experience!)
    $memoryPerSession = if ($PeakSessions -gt 0) {
        [math]::Round($currentSpecs.MemoryGB / $PeakSessions, 2)
    } else { 0 }
    
    # Workload-aware thresholds based on app group type
    # RemoteApp (RailApplications): Users run 1-2 published apps, lighter footprint per session
    #   - Higher sessions/vCPU is safe because each session consumes less CPU/RAM
    #   - Microsoft guidance supports 8-12 sessions/vCPU for light RemoteApp workloads
    # Desktop: Full desktop experience, heavier per-session resource usage
    #   - Microsoft recommendation: 4-6 sessions/vCPU for multi-session desktop
    # Personal: 1:1 mapping, session density is always 1 (or 0 if idle)
    $isRemoteApp = ($AppGroupType -eq "RailApplications")
    $isPersonal = ($HostPoolType -eq "Personal")
    
    if ($isRemoteApp) {
      $optimalSessionsPerVCpu = 8   # RemoteApp sweet spot
      $maxSessionsPerVCpu = 12      # RemoteApp maximum before degradation
      $minMemoryPerSession = 1.5    # RemoteApp apps need less RAM per session
      $optimalMemoryPerSession = 2  # Comfortable for published apps
      $workloadLabel = "RemoteApp"
    } elseif ($isPersonal) {
      $optimalSessionsPerVCpu = 1   # Personal is always 1:1
      $maxSessionsPerVCpu = 1
      $minMemoryPerSession = 4      # Full desktop needs more
      $optimalMemoryPerSession = 8
      $workloadLabel = "Personal Desktop"
    } else {
      # Pooled Desktop (default)
      $optimalSessionsPerVCpu = 5   # Desktop sweet spot
      $maxSessionsPerVCpu = 8       # Microsoft maximum for desktop
      $minMemoryPerSession = 2      # Absolute minimum
      $optimalMemoryPerSession = 4  # Recommended
      $workloadLabel = "Pooled Desktop"
    }
    
    # ========================================================
    # DECISION LOGIC - AVD-AWARE
    # ========================================================
    
    # Workload-aware density thresholds for Scenarios 3 and 5
    $lowDensityThreshold = if ($isRemoteApp) { 3 } elseif ($isPersonal) { 1 } else { 4 }
    $downsizeDensityThreshold = if ($isRemoteApp) { 5 } else { 3 }
    
    # CRITICAL: Check if we're hitting capacity limits FIRST
    $hitSessionDensityLimit = $sessionsPerVCpu -gt $maxSessionsPerVCpu
    $hitMemoryPerSessionLimit = $memoryPerSession -lt $minMemoryPerSession -and $PeakSessions -gt 0
    
    # --- UPSIZING SCENARIOS ---
    
    # Scenario 1: High session density + high CPU = Need more capacity
    if ($hitSessionDensityLimit -or ($sessionsPerVCpu -gt 6 -and $PeakCpu -gt 70)) {
        # Extract base family (D or E) and variant (ads, as, s, etc)
        $family = if ($CurrentSize -match "^(Standard_[DE]\d+)(ads|as|s)?") { 
            $matches[1] + $(if ($matches[2]) { $matches[2] } else { "" })
        }
        $largerSizes = $vmSizes.GetEnumerator() | Where-Object { 
            $_.Key -like "$family*" -and 
            $_.Value.vCPU -gt $currentSpecs.vCPU 
        } | Sort-Object { $_.Value.vCPU }
        
        if ($largerSizes) {
            $recommendation = $largerSizes[0].Key
            $reason += "High session density ($sessionsPerVCpu sessions/vCPU) with elevated CPU ($([math]::Round($PeakCpu,1))%)"
            $reason += "$workloadLabel Best Practice: Target $optimalSessionsPerVCpu sessions per vCPU (max $maxSessionsPerVCpu)"
        }
    }
    
    # Scenario 2: Memory pressure per session (not just total memory!)
    elseif ($hitMemoryPerSessionLimit) {
        # Need more memory - consider E-series or larger size
        if ($CurrentSize -match "D\d+(ads|as|s)?") {
            # Switch to E-series for better memory ratio (preserve variant and version)
            $recommendation = $CurrentSize -replace "D", "E"
            if (-not $vmSizes.ContainsKey($recommendation)) {
                # E-series equivalent doesn't exist, upsize in D-series
                $family = if ($CurrentSize -match "^(Standard_[DE]\d+)(ads|as|s)?") { 
                    $matches[1] + $(if ($matches[2]) { $matches[2] } else { "" })
                }
                $largerSizes = $vmSizes.GetEnumerator() | Where-Object { 
                    $_.Key -like "$family*" -and 
                    $_.Value.vCPU -gt $currentSpecs.vCPU 
                } | Sort-Object { $_.Value.vCPU }
                if ($largerSizes) {
                    $recommendation = $largerSizes[0].Key
                }
            }
            $reason += "Insufficient memory per session ($([math]::Round($memoryPerSession,1))GB per user with $PeakSessions peak sessions)"
            $reason += "$workloadLabel Recommendation: Minimum ${minMemoryPerSession}GB per user, optimal ${optimalMemoryPerSession}GB per user"
        }
        else {
            # Already E-series, upsize
            $family = if ($CurrentSize -match "^(Standard_[DE]\d+)") { $matches[1] }
            $largerSizes = $vmSizes.GetEnumerator() | Where-Object { 
                $_.Key -like "$family*" -and 
                $_.Value.vCPU -gt $currentSpecs.vCPU 
            } | Sort-Object { $_.Value.vCPU }
            if ($largerSizes) {
                $recommendation = $largerSizes[0].Key
            }
            $reason += "Memory pressure: Only $([math]::Round($memoryPerSession,1))GB per session (need minimum 2GB)"
        }
    }
    
    # Scenario 3: High CPU but LOW session density = Workload issue, not sizing issue
    # For RemoteApp: low density threshold is lower since each session is lighter
    elseif ($PeakCpu -gt $CpuOverloadThreshold -and $sessionsPerVCpu -lt $lowDensityThreshold -and $PeakSessions -gt 0) {
        $recommendation = "Keep Current"
        $reason += "High CPU ($([math]::Round($PeakCpu,1))%) but low session density ($sessionsPerVCpu sessions/vCPU) for $workloadLabel"
        $reason += "⚠️ Investigate: Apps may be inefficient or startup storms occurring"
        if ($isRemoteApp) {
          $reason += "Consider: Check published app resource usage, tune app launch behavior"
        } else {
          $reason += "Consider: App optimization, FSLogix tuning, or stagger user logons"
        }
    }
    
    # Scenario 4: High CPU but no session data = Legacy check (backward compatible)
    elseif ($PeakCpu -gt $CpuOverloadThreshold -and $PeakSessions -eq 0) {
        $family = if ($CurrentSize -match "^(Standard_[DE]\d+)") { $matches[1] }
        $largerSizes = $vmSizes.GetEnumerator() | Where-Object { 
            $_.Key -like "$family*" -and 
            $_.Value.vCPU -gt $currentSpecs.vCPU 
        } | Sort-Object { $_.Value.vCPU }
        
        if ($largerSizes) {
            $recommendation = $largerSizes[0].Key
            $reason += "High CPU peaks ($([math]::Round($PeakCpu,1))%)"
            $reason += "Note: No session data available - recommendation based on CPU only"
        }
    }
    
    # --- DOWNSIZING SCENARIOS ---
    
    # Scenario 5: Low utilization AND low session density = Can safely downsize
    # RemoteApp pools: lower density threshold for downsize (sessions are lighter)
    elseif ($AvgCpu -lt $CpuRightSizingThreshold -and $PeakCpu -lt 60 -and 
            ($sessionsPerVCpu -lt $downsizeDensityThreshold -or $PeakSessions -eq 0)) {
        
        $family = if ($CurrentSize -match "^(Standard_[DE]\d+)") { $matches[1] }
        $smallerSizes = $vmSizes.GetEnumerator() | Where-Object { 
            $_.Key -like "$family*" -and 
            $_.Value.vCPU -lt $currentSpecs.vCPU 
        } | Sort-Object { $_.Value.vCPU } -Descending
        
        if ($smallerSizes) {
            # Verify downsize won't cause session density or memory issues
            $proposedSize = $smallerSizes[0]
            $proposedSessionsPerVCpu = if ($PeakSessions -gt 0) {
                $PeakSessions / $proposedSize.Value.vCPU
            } else { 0 }
            $proposedMemoryPerSession = if ($PeakSessions -gt 0) {
                $proposedSize.Value.MemoryGB / $PeakSessions
            } else { 0 }
            
            # Only downsize if it won't create capacity issues
            if ($proposedSessionsPerVCpu -le $optimalSessionsPerVCpu -and 
                ($proposedMemoryPerSession -ge $optimalMemoryPerSession -or $PeakSessions -eq 0)) {
                
                $recommendation = $smallerSizes[0].Key
                $reason += "Low CPU utilization (Avg: $([math]::Round($AvgCpu,1))%, Peak: $([math]::Round($PeakCpu,1))%)"
                if ($PeakSessions -gt 0) {
                    $reason += "Low session density ($sessionsPerVCpu sessions/vCPU allows downsizing)"
                    $reason += "Post-downsize: $([math]::Round($proposedSessionsPerVCpu,1)) sessions/vCPU, $([math]::Round($proposedMemoryPerSession,1))GB per user"
                }
            }
            else {
                # Can't downsize - would create capacity issues
                $recommendation = "Keep Current"
                $reason += "Low CPU but downsizing would create capacity constraints"
                $reason += "Post-downsize would have: $([math]::Round($proposedSessionsPerVCpu,1)) sessions/vCPU (too high)"
            }
        }
    }
    
    # --- KEEP CURRENT ---
    
    if (-not $recommendation) {
        $recommendation = "Keep Current"
        if (@($reason).Count -eq 0) {
            $reason += "VM is appropriately sized for current $workloadLabel workload"
            if ($PeakSessions -gt 0) {
                $reason += "Session density: $sessionsPerVCpu sessions/vCPU (optimal: $optimalSessionsPerVCpu for $workloadLabel)"
                $reason += "Memory per user: $([math]::Round($memoryPerSession,1))GB (optimal: ${optimalMemoryPerSession}GB)"
            }
        }
    }
    
    return [PSCustomObject]@{
        Recommendation = $recommendation
        Reason = ($reason -join "; ")
        Confidence = $confidence
        EvidenceScore = $evidenceScore
        EvidenceSignals = ($evidenceSignals -join ", ")
        CurrentvCPU = $currentSpecs.vCPU
        CurrentMemoryGB = $currentSpecs.MemoryGB
        SessionsPerVCPU = if ($PeakSessions -gt 0) { $sessionsPerVCpu } else { $null }
        MemoryPerSessionGB = if ($PeakSessions -gt 0) { $memoryPerSession } else { $null }
    }
}

# =========================================================
# Zone Resiliency Analysis
# =========================================================
function Get-ZoneResiliencyScore {
    param(
        [string]$HostPoolName,
        [array]$VMs
    )
    
    $zoneDistribution = $VMs | Where-Object { $_.Zones } | 
        Group-Object Zones | 
        Select-Object Name, Count
    
    $uniqueZones = ($zoneDistribution | Measure-Object).Count
    $totalVMs = ($VMs | Measure-Object).Count
    $zoneEnabledVMs = ($VMs | Where-Object { $_.Zones } | Measure-Object).Count
    
    $score = 0
    $findings = @()
    $recommendations = @()
    
    # Scoring criteria
    if ($zoneEnabledVMs -eq 0) {
        $score = 0
        $findings += "No VMs are zone-enabled"
        $recommendations += "Deploy VMs across multiple availability zones for improved resiliency"
        $recommendations += "Aim for at least 2-3 zones with balanced VM distribution"
    } elseif ($zoneEnabledVMs -lt $totalVMs) {
        $score = 25
        $findings += "$zoneEnabledVMs of $totalVMs VMs are zone-enabled"
        $recommendations += "Enable zones for remaining $($totalVMs - $zoneEnabledVMs) VMs"
    } elseif ($uniqueZones -eq 1) {
        $score = 40
        $findings += "All VMs in single zone - no zone redundancy"
        $recommendations += "Distribute VMs across at least 2 availability zones"
    } elseif ($uniqueZones -eq 2) {
        $score = 75
        $findings += "VMs distributed across 2 zones"
        $recommendations += "Consider adding a third zone for enhanced resiliency"
        
        # Check balance
        $maxMeasure = $zoneDistribution | Measure-Object Count -Maximum
        $minMeasure = $zoneDistribution | Measure-Object Count -Minimum
        $maxVMsInZone = if ((SafeMeasure $maxMeasure 'Maximum')) { (SafeMeasure $maxMeasure 'Maximum') } else { 1 }
        $minVMsInZone = if ((SafeMeasure $minMeasure 'Minimum')) { (SafeMeasure $minMeasure 'Minimum') } else { 0 }
        $balance = if ($maxVMsInZone -gt 0) { [math]::Round(($minVMsInZone / $maxVMsInZone) * 100, 0) } else { 0 }
        
        if ($balance -lt 70) {
            $score -= 10
            $findings += "Unbalanced distribution across zones ($balance% balance)"
            $recommendations += "Rebalance VMs for more equal distribution across zones"
        }
    } else {
        $score = 100
        $findings += "VMs distributed across $uniqueZones zones"
        
        # Check balance
        $maxMeasure = $zoneDistribution | Measure-Object Count -Maximum
        $minMeasure = $zoneDistribution | Measure-Object Count -Minimum
        $maxVMsInZone = if ((SafeMeasure $maxMeasure 'Maximum')) { (SafeMeasure $maxMeasure 'Maximum') } else { 1 }
        $minVMsInZone = if ((SafeMeasure $minMeasure 'Minimum')) { (SafeMeasure $minMeasure 'Minimum') } else { 0 }
        $balance = if ($maxVMsInZone -gt 0) { [math]::Round(($minVMsInZone / $maxVMsInZone) * 100, 0) } else { 0 }
        
        if ($balance -lt 70) {
            $score -= 15
            $findings += "Unbalanced distribution across zones ($balance% balance)"
            $recommendations += "Rebalance VMs for more equal distribution across zones"
        } else {
            $findings += "Well-balanced distribution across zones"
        }
    }
    
    return [PSCustomObject]@{
        HostPoolName = $HostPoolName
        ResiliencyScore = $score
        TotalVMs = $totalVMs
        ZoneEnabledVMs = $zoneEnabledVMs
        UniqueZones = $uniqueZones
        ZoneDistribution = ($zoneDistribution | ForEach-Object { "Zone $($_.Name): $($_.Count) VMs" }) -join "; "
        Findings = ($findings -join "; ")
        Recommendations = ($recommendations -join "; ")
    }
}

# =========================================================
# DISCLAIMER
# =========================================================
Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                                                                       ║" -ForegroundColor Cyan
Write-Host "║          Enhanced AVD Evidence Pack - v4.0.0                          ║" -ForegroundColor Cyan
Write-Host "║          Azure Virtual Desktop Analysis & Optimization                ║" -ForegroundColor Cyan
Write-Host "║                                                                       ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

if (-not $SkipDisclaimer) {
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Yellow
    Write-Host "                    ⚠️  DISCLAIMER - READ CAREFULLY" -ForegroundColor Yellow
    Write-Host "========================================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "This script is provided AS-IS with:" -ForegroundColor White
    Write-Host "  • NO WARRANTY of any kind" -ForegroundColor White
    Write-Host "  • NO SUPPORT - use at your own risk" -ForegroundColor White
    Write-Host "  • NO LIABILITY for any damages or issues" -ForegroundColor White
    Write-Host ""
    Write-Host "You are responsible for:" -ForegroundColor White
    Write-Host "  • Testing in non-production environments first" -ForegroundColor White
    Write-Host "  • Reviewing all outputs before taking action" -ForegroundColor White
    Write-Host "  • Validating all recommendations against your requirements" -ForegroundColor White
    Write-Host ""
    Write-Host "Cost estimates are approximations only - validate against actual billing" -ForegroundColor White
    Write-Host ""
    Write-Host "By continuing, you accept these terms and conditions." -ForegroundColor Yellow
    Write-Host "========================================================================" -ForegroundColor Yellow
    Write-Host ""

    # Wait for user acknowledgment
    Write-Host "Press any key to continue or Ctrl+C to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Write-Host ""
} else {
    Write-Host "⚠️  Running with -SkipDisclaimer: You accept all terms and conditions" -ForegroundColor Yellow
    Write-Host ""
}

# =========================================================
# Authentication
# =========================================================
# Check if we're already logged in with a managed identity (e.g., running in Azure Container App)
$existingContext = Get-AzContext -ErrorAction SilentlyContinue
$isManagedIdentity = $existingContext -and $existingContext.Account.Type -eq 'ManagedService'

if ($isManagedIdentity) {
    Write-Host "  ✓ Using existing Managed Identity connection" -ForegroundColor Green
    Write-Host "    Tenant: $($existingContext.Tenant.Id)" -ForegroundColor Gray
} else {
    Disable-AzContextAutosave -Scope Process | Out-Null
    Clear-AzContext -Scope Process -Force -ErrorAction SilentlyContinue | Out-Null
    Connect-AzAccount -TenantId $TenantId | Out-Null
}

# =========================================================
# DRY RUN MODE
# =========================================================
if ($DryRun) {
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║                         DRY RUN MODE                                  ║" -ForegroundColor Yellow
    Write-Host "║                   No data will be collected                           ║" -ForegroundColor Yellow
    Write-Host "╚═══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    Write-Host ""
    
    Write-Host "Estimating environment size..." -ForegroundColor Cyan
    $totalVMs = 0
    $totalHostPools = 0
    foreach ($subId in $SubscriptionIds) {
        try {
            Set-AzContext -SubscriptionId $subId | Out-Null
            $vmsInSub = (Get-AzVM -ErrorAction SilentlyContinue | Measure-Object).Count
            $hpsInSub = (Get-AzWvdHostPool -ErrorAction SilentlyContinue | Measure-Object).Count
            $totalVMs += $vmsInSub
            $totalHostPools += $hpsInSub
            Write-Host "  Subscription $($subId): ~$vmsInSub VMs, ~$hpsInSub host pools" -ForegroundColor Gray
        }
        catch {
            Write-Host "  ⚠ Could not query subscription $subId" -ForegroundColor Yellow
        }
    }
    
    # Calculate estimates — batched API calls: 1 per VM instead of 12
    $estimatedMetricsMinutes = [math]::Ceiling($totalVMs / 50)
    
    $estimatedTotalMinutes = 3 + $estimatedMetricsMinutes + 5
    if ($SkipAzureMonitorMetrics) {
        $estimatedTotalMinutes = 8
    }
    
    $estimatedApiCalls = $totalVMs  # 1 batched call per VM
    $estimatedOutputSizeMB = [math]::Ceiling($totalVMs * 0.05)
    
    Write-Host ""
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "ANALYSIS PREVIEW" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Environment to Analyze:" -ForegroundColor Yellow
    Write-Host "  • Subscriptions: $(SafeCount $SubscriptionIds)" -ForegroundColor White
    Write-Host "  • Estimated Host Pools: ~$totalHostPools" -ForegroundColor White
    Write-Host "  • Estimated VMs: ~$totalVMs" -ForegroundColor White
    Write-Host "  • Metrics lookback: $MetricsLookbackDays days" -ForegroundColor White
    Write-Host "  • PowerShell: v$($PSVersionTable.PSVersion)" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Estimated Runtime:" -ForegroundColor Yellow
    Write-Host "  • Resource collection: ~3 minutes" -ForegroundColor White
    if (-not $SkipAzureMonitorMetrics) {
        Write-Host "  • Metrics collection: ~$estimatedMetricsMinutes minutes" -ForegroundColor White
    } else {
        Write-Host "  • Metrics collection: SKIPPED (-SkipAzureMonitorMetrics)" -ForegroundColor Gray
    }
    Write-Host "  • Analysis & export: ~5 minutes" -ForegroundColor White
    Write-Host "  • TOTAL RUNTIME: ~$estimatedTotalMinutes minutes" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Resource Usage:" -ForegroundColor Yellow
    Write-Host "  • Azure API calls: ~$estimatedApiCalls" -ForegroundColor White
    Write-Host "  • Output size: ~$estimatedOutputSizeMB MB" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Will Generate:" -ForegroundColor Yellow
    Write-Host "  ✓ VM right-sizing recommendations" -ForegroundColor Green
    Write-Host "  ✓ Zone resiliency analysis" -ForegroundColor Green
    Write-Host "  ✓ Cost optimization analysis" -ForegroundColor Green
    if ($IncludeAzureAdvisor) {
        Write-Host "  ✓ Azure Advisor recommendations" -ForegroundColor Green
    }
    if ($GenerateHtmlReport) {
        Write-Host "  ✓ HTML report" -ForegroundColor Green
    }
    if ($IncludeIncidentWindowQueries) {
        Write-Host "  ✓ Incident window comparison" -ForegroundColor Green
    }
    Write-Host ""
    
    if ($SkipAzureMonitorMetrics) {
        Write-Host "⚠ NOTE: Skipping metrics reduces runtime but limits recommendations" -ForegroundColor Yellow
        Write-Host ""
    }
    
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Yellow
    Write-Host "This was a DRY RUN - no data was collected." -ForegroundColor Yellow
    Write-Host "Remove -DryRun parameter to run the actual analysis." -ForegroundColor Yellow
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Yellow
    Write-Host ""
    
    exit 0
}

# =========================================================
# Output setup (with Resume support)
# =========================================================
$resuming = $false
if ($ResumeFrom) {
    if (Test-Path $ResumeFrom) {
        $outFolder = $ResumeFrom
        $resuming = $true
        Write-Host ""
        Write-Host "♻️  RESUMING from checkpoint: $ResumeFrom" -ForegroundColor Green
        Write-Host ""
    }
    else {
        Write-Error "ERROR: Resume folder not found: $ResumeFrom"
        Write-Host "Make sure the path is correct and the folder exists." -ForegroundColor Yellow
        exit 1
    }
}
else {
    $ts = Get-Date -Format "yyyyMMdd-HHmmss"
    $outFolder = "Enhanced-AVD-EvidencePack-$ts"
    New-Item -ItemType Directory -Path $outFolder | Out-Null
}

# =========================================================
# Data containers (using Generic Lists for O(1) append instead of O(n) array copy)
# =========================================================
$hostPools              = [System.Collections.Generic.List[object]]::new()
$sessionHosts           = [System.Collections.Generic.List[object]]::new()
$vms                    = [System.Collections.Generic.List[object]]::new()
$vmss                   = [System.Collections.Generic.List[object]]::new()
$vmssInstances          = [System.Collections.Generic.List[object]]::new()
$scalingPlans           = [System.Collections.Generic.List[object]]::new()
$scalingPlanAssignments = [System.Collections.Generic.List[object]]::new()
$scalingPlanSchedules   = [System.Collections.Generic.List[object]]::new()
$vmMetrics              = [System.Collections.Generic.List[object]]::new()
$vmMetricsIncident      = [System.Collections.Generic.List[object]]::new()
$laResults              = [System.Collections.Generic.List[object]]::new()

# Enhanced containers
$vmRightSizing          = [System.Collections.Generic.List[object]]::new()
$zoneResiliency         = [System.Collections.Generic.List[object]]::new()
$costAnalysis           = [System.Collections.Generic.List[object]]::new()
$incidentAnalysis       = [System.Collections.Generic.List[object]]::new()
$advisorRecommendations = [System.Collections.Generic.List[object]]::new()
$reservationAnalysis    = [System.Collections.Generic.List[object]]::new()
$existingReservations   = [System.Collections.Generic.List[object]]::new()

# =========================================================
# Resume: Load Checkpoints
# =========================================================
if ($resuming) {
    Write-Host "Loading saved data from checkpoints..." -ForegroundColor Cyan
    
    # Load collection checkpoint
    if (Test-Checkpoint -CheckpointName "collection" -OutputFolder $outFolder) {
        $saved = Load-Checkpoint -CheckpointName "collection" -OutputFolder $outFolder
        foreach ($item in @($saved.hostPools)) { $hostPools.Add($item) }
        foreach ($item in @($saved.sessionHosts)) { $sessionHosts.Add($item) }
        foreach ($item in @($saved.vms)) { $vms.Add($item) }
        foreach ($item in @($saved.vmss)) { $vmss.Add($item) }
        foreach ($item in @($saved.vmssInstances)) { $vmssInstances.Add($item) }
        foreach ($item in @($saved.scalingPlans)) { $scalingPlans.Add($item) }
        foreach ($item in @($saved.scalingPlanAssignments)) { $scalingPlanAssignments.Add($item) }
        foreach ($item in @($saved.scalingPlanSchedules)) { $scalingPlanSchedules.Add($item) }
        Write-Host "  ✓ Loaded collection data: $(SafeCount $vms) VMs" -ForegroundColor Green
    }
    
    # Load metrics checkpoint
    if (Test-Checkpoint -CheckpointName "metrics" -OutputFolder $outFolder) {
        $saved = Load-Checkpoint -CheckpointName "metrics" -OutputFolder $outFolder
        foreach ($item in @($saved.vmMetrics)) { $vmMetrics.Add($item) }
        Write-Host "  ✓ Loaded metrics data: $(SafeCount $vmMetrics) records" -ForegroundColor Green
    }
    
    Write-Host ""
}

# =========================================================
# Autoscale: ARM-first enumerator
# =========================================================
function Get-ScalingPlansArm {
  param([string]$SubscriptionId)
  Set-AzContext -SubscriptionId $SubscriptionId -TenantId $TenantId | Out-Null
  $plans = Get-AzResource -ResourceType "Microsoft.DesktopVirtualization/scalingPlans" -ExpandProperties -ErrorAction SilentlyContinue
  return (SafeArray $plans)
}

function Expand-ScalingPlanEvidence {
  param([object]$PlanResource)

  if (-not $PlanResource) { return }

  $planId = $PlanResource.ResourceId
  $subId  = Get-SubFromArmId $planId
  $rg     = $PlanResource.ResourceGroupName
  $name   = $PlanResource.Name
  $loc    = $PlanResource.Location

  $props = $PlanResource.Properties

  $script:scalingPlans.Add([PSCustomObject]@{
    SubscriptionId = $subId
    ResourceGroup  = $rg
    ScalingPlanName= $name
    Location       = $loc
    TimeZone       = $props.timeZone
    HostPoolType   = $props.hostPoolType
    Description    = $props.description
    FriendlyName   = $props.friendlyName
    Id             = $planId
  })

  foreach ($hpr in SafeArray $props.hostPoolReferences) {
    $hpArmId = $hpr.hostPoolArmPath
    $script:scalingPlanAssignments.Add([PSCustomObject]@{
      SubscriptionId      = $subId
      ResourceGroup       = $rg
      ScalingPlanName     = $name
      ScalingPlanId       = $planId
      HostPoolArmId       = $hpArmId
      HostPoolName        = (Get-NameFromArmId $hpArmId)
      IsEnabled           = $hpr.scalingPlanEnabled
    })
  }

  foreach ($sch in SafeArray $props.schedules) {
    $script:scalingPlanSchedules.Add([PSCustomObject]@{
      SubscriptionId   = $subId
      ResourceGroup    = $rg
      ScalingPlanName  = $name
      ScalingPlanId    = $planId
      ScheduleName     = $sch.name
      DaysOfWeek       = (($sch.daysOfWeek) -join ",")
      RampUpStartTime  = $sch.rampUpStartTime
      PeakStartTime    = $sch.peakStartTime
      RampDownStartTime= $sch.rampDownStartTime
      OffPeakStartTime = $sch.offPeakStartTime
      RampUpCapacity   = $sch.rampUpCapacityThresholdPct
      RampDownCapacity = $sch.rampDownCapacityThresholdPct
      RampDownForceLogoff = $sch.rampDownForceLogoffUsers
      RampDownLogoffTimeoutMinutes = $sch.rampDownWaitTimeMinutes
      RampDownNotificationMessage  = $sch.rampDownNotificationMessage
    })
  }
}

# =========================================================
# Metrics (Optimized: single batched API call per VM)
# =========================================================
# Azure Monitor limits: 12,000 requests per hour per subscription (per provider).
# By batching all metrics + aggregations into ONE call per VM, we reduce API load
# by ~12x compared to the original 6-metric × 2-aggregation approach.
# ThrottleLimit is kept conservative for production safety.
function Collect-VmMetrics {
  param(
    [string]$VmId,
    [int]$LookbackDays,
    [int]$TimeGrainMinutes
  )

  $results = @()
  if (-not $VmId) { return $results }

  $start = (Get-Date).AddDays(-$LookbackDays)
  $end   = Get-Date
  $grain = New-TimeSpan -Minutes $TimeGrainMinutes

  # Batch: only collect metrics that are actually used by analysis engines
  # CPU (Avg/Max) → right-sizing, UX scoring, capacity analysis
  # Memory (Avg/Max) → right-sizing, memory-per-session
  $metricNames = @("Percentage CPU", "Available Memory Bytes")
  $aggregations = @("Average", "Maximum")

  $attempt = 0
  $maxRetries = 3
  $success = $false

  while (-not $success -and $attempt -lt $maxRetries) {
    try {
      $metricObjects = Get-AzMetric `
        -ResourceId $VmId `
        -MetricName $metricNames `
        -Aggregation $aggregations `
        -StartTime $start `
        -EndTime $end `
        -TimeGrain $grain `
        -ErrorAction Stop `
        -WarningAction SilentlyContinue

      foreach ($metric in SafeArray $metricObjects) {
        $mName = $metric.Name.Value
        foreach ($ts in SafeArray $metric.Timeseries) {
          foreach ($pt in SafeArray $ts.Data) {
            foreach ($agg in $aggregations) {
              $value = if ($agg -eq "Average") { $pt.Average } else { $pt.Maximum }
              if ($null -ne $value) {
                $results += [PSCustomObject]@{
                  VmId        = $VmId
                  Metric      = $mName
                  Aggregation = $agg
                  TimeStamp   = $pt.TimeStamp
                  Value       = $value
                }
              }
            }
          }
        }
      }
      $success = $true
    }
    catch {
      $attempt++
      $errorMsg = $_.Exception.Message
      $isThrottling = $errorMsg -match "Too Many Requests|429|throttl|rate limit"

      if ($isThrottling -and $attempt -lt $maxRetries) {
        $delay = 30 * $attempt  # 30s, 60s, 90s
        Write-Host "  ⚠ Throttled on metrics batch - retrying in ${delay}s (attempt $attempt/$maxRetries)..." -ForegroundColor Yellow
        Start-Sleep -Seconds $delay
      }
      elseif (-not $isThrottling) {
        # Non-throttle error - don't retry
        break
      }
    }
  }

  return $results
}

# =========================================================
# Log Analytics
# =========================================================
function Invoke-LaQuery {
    param(
        [string]$WorkspaceResourceId,
        [string]$Label,
        [string]$Query,
        [datetime]$StartTime,
        [datetime]$EndTime
    )

    $parts = $WorkspaceResourceId -split '/'
    $resourceGroupName = $parts[4]
    $workspaceName     = $parts[8]

    try {
        $workspace = Get-AzOperationalInsightsWorkspace `
            -ResourceGroupName $resourceGroupName `
            -Name $workspaceName `
            -ErrorAction Stop
    }
    catch {
        return [PSCustomObject]@{
            WorkspaceResourceId = $WorkspaceResourceId
            Label               = $Label
            QueryName           = "Meta"
            Status              = "WorkspaceNotFound"
            RowCount            = 0
        }
    }

    $duration = New-TimeSpan -Start $StartTime -End $EndTime

    try {
        $result = Invoke-AzOperationalInsightsQuery `
            -WorkspaceId $workspace.CustomerId `
            -Query $Query `
            -Timespan $duration `
            -ErrorAction Stop
    }
    catch {
        return [PSCustomObject]@{
            WorkspaceResourceId = $WorkspaceResourceId
            Label               = $Label
            QueryName           = "Meta"
            Status              = "QueryFailed"
            Error               = $_.Exception.Message
            RowCount            = 0
        }
    }

    if (-not $result.Results -or @($result.Results).Count -eq 0) {
        return [PSCustomObject]@{
            WorkspaceResourceId = $WorkspaceResourceId
            Label               = $Label
            QueryName           = "Meta"
            Status              = "NoRowsReturned"
            RowCount            = 0
        }
    }

    $output = @()
    foreach ($row in $result.Results) {
        $o = [PSCustomObject]@{
            WorkspaceResourceId = $WorkspaceResourceId
            Label               = $Label
            QueryName           = "AVD"
        }

        foreach ($p in $row.PSObject.Properties) {
            Add-Member -InputObject $o -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
        }

        $output += $o
    }
    
    return $output
}

# =========================================================
# KQL Queries
# =========================================================
$kqlTableDiscovery = @"
search *
| where TimeGenerated > ago(7d)
| summarize Count = count() by Type
| order by Count desc
"@

$kqlWvdConnections = @"
WVDConnections
| summarize Connections = count() by UserName, ClientOS
| order by Connections desc
"@

$kqlShortpathUsage = @"
WVDConnections
| extend TransportType = tostring(UdpUse)
| summarize Connections = count() by TransportType
"@

$kqlPeakConcurrency = @"
WVDConnections
| extend TimeSlot = bin(TimeGenerated, 15m)
| summarize ConcurrentSessions = dcount(CorrelationId) by TimeSlot
| summarize PeakConcurrentSessions = max(ConcurrentSessions)
"@

$kqlAutoscaleActivity = @"
union isfuzzy=true
    (WVDAutoscaleEvaluationPooled
    | summarize EvaluationCount = count() by Result),
    (print Result="NoTable", EvaluationCount=0 | where 1==0)
"@

$kqlAutoscaleDetailedActivity = @"
union isfuzzy=true
    (WVDAutoscaleEvaluationPooled
    | extend HostPoolPath = tostring(Properties.hostPoolArmPath),
             PoolName = tostring(split(Properties.hostPoolArmPath, '/')[-1]),
             ResultType = tostring(ResultType),
             ActiveHosts = toint(Properties.activeSessionHostCount),
             TotalHosts = toint(Properties.totalSessionHostCount),
             SessionsOnActive = toint(Properties.activeSessionCount),
             ConfigCapacityThreshold = toint(Properties.capacityThreshold)
    | summarize
        Evaluations = count(),
        Succeeded = countif(ResultType == "Succeeded" or Result == "Succeeded"),
        Failed = countif(ResultType == "Failed" or Result == "Failed"),
        AvgActiveHosts = round(avg(ActiveHosts), 1),
        MaxActiveHosts = max(ActiveHosts),
        AvgTotalHosts = round(avg(TotalHosts), 1),
        AvgSessions = round(avg(SessionsOnActive), 1)
        by PoolName
    | where isnotempty(PoolName)),
    (print PoolName="NoTable", Evaluations=0, Succeeded=0, Failed=0, AvgActiveHosts=0.0, MaxActiveHosts=0, AvgTotalHosts=0.0, AvgSessions=0.0 | where 1==0)
"@

$kqlSessionDuration = @"
WVDConnections
| where State == "Connected"
| project CorrelationId, UserName, ConnectTime = TimeGenerated
| join kind=inner (
    WVDConnections
    | where State == "Completed"
    | project CorrelationId, EndTime = TimeGenerated
) on CorrelationId
| extend SessionDurationMinutes = datetime_diff('minute', EndTime, ConnectTime)
| where SessionDurationMinutes >= 0
| summarize AvgDuration = round(avg(SessionDurationMinutes), 1), MaxDuration = max(SessionDurationMinutes) by UserName
"@

$kqlLoginTime = @"
WVDConnections
| where State == "Connected"
| extend HostPool = tostring(split(_ResourceId, '/')[-1])
| project CorrelationId, HostPool, UserName, ConnectTime = TimeGenerated
| join kind=inner (
    WVDConnections
    | where State == "Started"
    | project CorrelationId, StartTime = TimeGenerated
) on CorrelationId
| extend LoginDurationSec = datetime_diff('second', ConnectTime, StartTime)
| where LoginDurationSec >= 0 and LoginDurationSec < 600
| summarize AvgLoginSec = round(avg(LoginDurationSec), 1), P50LoginSec = round(percentile(LoginDurationSec, 50), 1), P95LoginSec = round(percentile(LoginDurationSec, 95), 1), MaxLoginSec = max(LoginDurationSec), TotalConnections = count() by HostPool
"@

$kqlConnectionSuccessRate = @"
WVDConnections
| extend HostPool = tostring(split(_ResourceId, '/')[-1])
| summarize
    TotalAttempts = count(),
    Succeeded = countif(State == "Connected" or State == "Completed"),
    Failed = countif(State == "Failed" or isnotempty(FailureReason)),
    UniqueUsers = dcount(UserName),
    PeakConcurrent = max(dcount(UserName))
    by HostPool
| extend SuccessRate = round(todouble(Succeeded) / todouble(TotalAttempts) * 100, 1),
         FailureRate = round(todouble(Failed) / todouble(TotalAttempts) * 100, 1)
| order by FailureRate desc
"@

# =========================================================
# NEW KQL Queries (v3.0.0 - rewritten v4.0.0 with correct schemas)
# =========================================================

# Connection Duration / Profile Load Performance
# WVDConnections doesn't have SessionStartTime/SessionEstablishedTime directly.
# We self-join on CorrelationId: Started->Connected = connection establishment time.
$kqlProfileLoadPerformance = @"
WVDConnections
| where State == "Started"
| project CorrelationId, StartTime = TimeGenerated
| join kind=inner (
    WVDConnections
    | where State == "Connected"
    | project CorrelationId, SessionHostName, ConnectTime = TimeGenerated
) on CorrelationId
| extend ConnectionTimeSec = datetime_diff('second', ConnectTime, StartTime)
| where ConnectionTimeSec >= 0 and ConnectionTimeSec < 600
| extend HostName = iif(SessionHostName contains "/", tostring(split(SessionHostName, '/')[1]), SessionHostName)
| summarize
    AvgProfileLoadSec = round(avg(ConnectionTimeSec), 1),
    P50ProfileLoadSec = round(percentile(ConnectionTimeSec, 50), 1),
    P95ProfileLoadSec = round(percentile(ConnectionTimeSec, 95), 1),
    MaxProfileLoadSec = round(max(ConnectionTimeSec), 1),
    TotalSessions = count(),
    SlowLogins_Over30s = countif(ConnectionTimeSec > 30),
    VerySlowLogins_Over60s = countif(ConnectionTimeSec > 60)
    by SessionHostName = HostName
| extend SlowLoginPct = round(todouble(SlowLogins_Over30s) / TotalSessions * 100, 1)
| order by P95ProfileLoadSec desc
"@

# Connection Quality (RTT and Bandwidth) by Client OS
# RTT/Bandwidth live in WVDConnectionNetworkData, joined to WVDConnections for ClientOS
$kqlConnectionQuality = @"
WVDConnectionNetworkData
| where isnotnull(EstRoundTripTimeInMs) and EstRoundTripTimeInMs > 0
| join kind=inner (
    WVDConnections
    | where State == "Connected"
    | project CorrelationId, ClientOS
) on CorrelationId
| summarize
    AvgRTTms = round(avg(EstRoundTripTimeInMs), 1),
    P50RTTms = round(percentile(EstRoundTripTimeInMs, 50), 1),
    P95RTTms = round(percentile(EstRoundTripTimeInMs, 95), 1),
    MaxRTTms = round(max(EstRoundTripTimeInMs), 1),
    AvgBandwidthKBps = round(avg(EstAvailableBandwidthKBps), 0),
    MinBandwidthKBps = round(min(EstAvailableBandwidthKBps), 0),
    Connections = dcount(CorrelationId),
    HighLatency_Over150ms = countif(EstRoundTripTimeInMs > 150),
    PoorLatency_Over250ms = countif(EstRoundTripTimeInMs > 250)
    by ClientOS
| extend HighLatencyPct = round(todouble(HighLatency_Over150ms) / Connections * 100, 1)
| order by P95RTTms desc
"@

# Connection Quality by Gateway Region
$kqlConnectionQualityByRegion = @"
WVDConnectionNetworkData
| where isnotnull(EstRoundTripTimeInMs) and EstRoundTripTimeInMs > 0
| join kind=inner (
    WVDConnections
    | where State == "Connected"
    | project CorrelationId, GatewayRegion
) on CorrelationId
| where isnotempty(GatewayRegion)
| summarize
    AvgRTTms = round(avg(EstRoundTripTimeInMs), 1),
    P95RTTms = round(percentile(EstRoundTripTimeInMs, 95), 1),
    AvgBandwidthKBps = round(avg(EstAvailableBandwidthKBps), 0),
    Connections = dcount(CorrelationId)
    by GatewayRegion
| order by Connections desc
"@

# Connection Errors and Failures (from WVDErrors table)
$kqlConnectionErrors = @"
WVDErrors
| summarize
    ErrorCount = count(),
    DistinctUsers = dcount(UserName),
    DistinctCorrelations = dcount(CorrelationId)
    by CodeSymbolic, Message = substring(Message, 0, 200)
| order by ErrorCount desc
| take 50
"@

# Unexpected Disconnects
# Use state transitions: Started->Connected->Completed. Short sessions = unexpected.
$kqlDisconnects = @"
let starts = WVDConnections | where State == "Connected" | project CorrelationId, SessionHostName, ConnectTime = TimeGenerated;
let ends = WVDConnections | where State == "Completed" | project CorrelationId, EndTime = TimeGenerated;
starts
| join kind=inner ends on CorrelationId
| extend SessionDurationSec = datetime_diff('second', EndTime, ConnectTime)
| extend IsUnexpected = (SessionDurationSec < 60)
| extend HostName = iif(SessionHostName contains "/", tostring(split(SessionHostName, '/')[1]), SessionHostName)
| summarize
    TotalSessions = count(),
    UnexpectedDisconnects = countif(IsUnexpected),
    AvgSessionMinutes = round(avg(SessionDurationSec / 60.0), 1)
    by SessionHostName = HostName
| extend DisconnectPct = round(todouble(UnexpectedDisconnects) / TotalSessions * 100, 1)
| where TotalSessions > 5
| order by DisconnectPct desc
"@

# Disconnect Reason Categorization
# Joins WVDConnections with WVDErrors to classify why sessions ended
# Uses actual AVD CodeSymbolic values (e.g. ConnectionBrokenMissedHeartbeatThresholdExceeded)
$kqlDisconnectReasons = @"
let sessions = WVDConnections
| where State == "Connected"
| project CorrelationId, SessionHostName, UserName;
let completions = WVDConnections
| where State == "Completed"
| project CorrelationId, CompletedTime = TimeGenerated;
let errors = WVDErrors
| summarize ErrorCode = take_any(CodeSymbolic), ErrorMsg = take_any(substring(Message, 0, 150)) by CorrelationId;
let enriched = sessions
| join kind=leftouter completions on CorrelationId
| join kind=leftouter errors on CorrelationId;
let completedOrErrored = enriched
| where isnotnull(CompletedTime) or isnotempty(ErrorCode);
let totalCompleted = toscalar(completedOrErrored | summarize dcount(CorrelationId));
completedOrErrored
| extend Category = case(
    // Network / Heartbeat - connection lost between client and host
    isnotempty(ErrorCode) and (ErrorCode has_any ("MissedHeartbeat", "ConnectionBroken", "ConnectionLost", "NL_DISCONNECT", "CM_ERR_MISSED_HEARTBEAT", "TransportClose", "ConnectionDropped", "SocketClose") or ErrorCode contains "Shortpath" or ErrorCode contains "NetworkDrop" or ErrorCode contains "PeerLeg" or ErrorCode contains "Heartbeat"), "Network Drop",
    // User-initiated - normal user logoff or client disconnect
    isnotempty(ErrorCode) and ErrorCode has_any ("ClientDisconnect", "LogoffByUser", "UserInitiated", "ConnectionFailedClientDisconnect"), "User Initiated",
    // Idle / Activity timeout - session timed out due to inactivity
    isnotempty(ErrorCode) and ErrorCode has_any ("ActivityTimeout", "SessionTimeout", "IdleTimeout", "SessionLogoff", "IdleDisconnect"), "Idle Timeout",
    // Authentication - logon, password, or credential failures
    isnotempty(ErrorCode) and ErrorCode has_any ("LogonFailed", "AuthenticationLogonFailed", "FreshCredsRequired", "PasswordExpired", "AccountLocked", "AccountExpired", "InvalidCredentials", "SavedCredentialsNotAllowed", "CredSSP", "Kerberos", "NLA", "AutoReconnectNoCookie"), "Authentication",
    // Server-side - host shutdown, reboot, or scaling action
    isnotempty(ErrorCode) and ErrorCode has_any ("ServerDisconnect", "ServerShutdown", "HostShutdown", "SessionHostShutdown", "ServerMaintenanceDisconnect"), "Server Side",
    // Resource exhaustion - out of memory, disk full
    isnotempty(ErrorCode) and ErrorCode has_any ("OutOfMemory", "DiskFull", "ResourceExhausted", "QuotaExceeded"), "Resource Exhaustion",
    // No healthy host available - capacity or health issue
    isnotempty(ErrorCode) and ErrorCode has_any ("NoHealthyRdsh", "NoHealthySession", "MaxSession", "SessionHostNotFound", "HostPoolNotFound", "NoAvailableSession"), "Licensing/Capacity",
    // Agent / Health - RD Agent or host health issues
    isnotempty(ErrorCode) and ErrorCode has_any ("AgentRegistration", "AgentHeartbeat", "Unavailable", "HostNotAvailable", "DomainJoin", "DomainTrust"), "Agent/Health",
    // Gateway / Broker - control plane issues
    isnotempty(ErrorCode) and ErrorCode has_any ("GatewayError", "BrokerError", "OrchestrationError", "ReverseConnect", "PendingReconnect"), "Gateway/Broker",
    // FSLogix / Profile - profile container issues
    isnotempty(ErrorCode) and ErrorCode has_any ("ERROR_PATH_NOT_FOUND", "ProfileDisk", "FSLogix", "VHD"), "Profile/FSLogix",
    // Catch remaining known non-critical codes
    isnotempty(ErrorCode) and ErrorCode has_any ("AutoReconnect", "Reconnect"), "Auto-Reconnect",
    // Anything else with an error code
    isnotempty(ErrorCode), strcat("Other: ", ErrorCode),
    "Normal Completion"
)
| extend HostName = iif(SessionHostName contains "/", tostring(split(SessionHostName, '/')[1]), SessionHostName)
| summarize
    SessionCount = dcount(CorrelationId),
    DistinctUsers = dcount(UserName),
    DistinctHosts = dcount(HostName),
    SampleError = take_any(ErrorMsg)
    by DisconnectCategory = Category
| extend Pct = round(100.0 * SessionCount / totalCompleted, 1)
| order by SessionCount desc
"@

# Disconnect Breakdown by Host - which hosts have the most abnormal disconnects
# Uses same real AVD CodeSymbolic matching as the reasons query
$kqlDisconnectsByHost = @"
let sessions = WVDConnections
| where State == "Connected"
| project CorrelationId, SessionHostName;
let completions = WVDConnections
| where State == "Completed"
| project CorrelationId, CompletedTime = TimeGenerated;
let errors = WVDErrors
| summarize ErrorCode = take_any(CodeSymbolic) by CorrelationId;
sessions
| join kind=leftouter completions on CorrelationId
| join kind=leftouter errors on CorrelationId
| where isnotnull(CompletedTime) or isnotempty(ErrorCode)
| extend HostName = iif(SessionHostName contains "/", tostring(split(SessionHostName, '/')[1]), SessionHostName)
| extend IsAbnormal = isnotempty(ErrorCode) and not(ErrorCode has_any ("ClientDisconnect", "LogoffByUser", "UserInitiated", "ActivityTimeout", "SessionTimeout", "IdleTimeout", "SessionLogoff", "IdleDisconnect", "AutoReconnect", "Reconnect", "SavedCredentialsNotAllowed", "AutoReconnectNoCookie"))
| summarize
    TotalSessions = dcount(CorrelationId),
    AbnormalDisconnects = dcountif(CorrelationId, IsAbnormal),
    NetworkDrops = dcountif(CorrelationId, isnotempty(ErrorCode) and (ErrorCode has_any ("MissedHeartbeat", "ConnectionBroken", "ConnectionLost", "NL_DISCONNECT", "CM_ERR_MISSED_HEARTBEAT", "TransportClose", "ConnectionDropped", "SocketClose") or ErrorCode contains "Shortpath" or ErrorCode contains "NetworkDrop" or ErrorCode contains "PeerLeg" or ErrorCode contains "Heartbeat")),
    Timeouts = dcountif(CorrelationId, isnotempty(ErrorCode) and ErrorCode has_any ("ActivityTimeout", "SessionTimeout", "IdleTimeout", "SessionLogoff", "IdleDisconnect")),
    ServerSide = dcountif(CorrelationId, isnotempty(ErrorCode) and ErrorCode has_any ("ServerDisconnect", "ServerShutdown", "HostShutdown", "SessionHostShutdown", "ServerMaintenanceDisconnect")),
    AuthFailures = dcountif(CorrelationId, isnotempty(ErrorCode) and ErrorCode has_any ("LogonFailed", "AuthenticationLogonFailed", "FreshCredsRequired", "PasswordExpired", "AccountLocked", "AccountExpired", "InvalidCredentials", "CredSSP", "Kerberos")),
    ResourceIssues = dcountif(CorrelationId, isnotempty(ErrorCode) and ErrorCode has_any ("OutOfMemory", "DiskFull", "ResourceExhausted", "QuotaExceeded")),
    OtherErrors = dcountif(CorrelationId, IsAbnormal and not(ErrorCode has_any ("MissedHeartbeat", "ConnectionBroken", "ConnectionLost", "NL_DISCONNECT", "CM_ERR_MISSED_HEARTBEAT", "TransportClose", "ConnectionDropped", "SocketClose", "ServerDisconnect", "ServerShutdown", "HostShutdown", "SessionHostShutdown", "ServerMaintenanceDisconnect", "LogonFailed", "AuthenticationLogonFailed", "FreshCredsRequired", "PasswordExpired", "AccountLocked", "AccountExpired", "InvalidCredentials", "CredSSP", "Kerberos", "OutOfMemory", "DiskFull", "ResourceExhausted", "QuotaExceeded")))
    by SessionHostName = HostName
| extend AbnormalPct = round(100.0 * AbnormalDisconnects / TotalSessions, 1)
| where TotalSessions > 5
| order by AbnormalPct desc
| take 30
"@


# Scaling Plan Effectiveness - Hourly Concurrency Pattern
$kqlHourlyConcurrency = @"
WVDConnections
| where State == "Connected"
| extend HourOfDay = hourofday(TimeGenerated), DayOfWeek = dayofweek(TimeGenerated)
| where DayOfWeek >= 1d and DayOfWeek <= 5d
| extend TimeSlot = bin(TimeGenerated, 15m)
| summarize ConcurrentSessions = dcount(CorrelationId) by TimeSlot, HourOfDay
| summarize
    AvgConcurrency = round(avg(ConcurrentSessions), 0),
    PeakConcurrency = max(ConcurrentSessions),
    P95Concurrency = round(percentile(ConcurrentSessions, 95), 0)
    by HourOfDay
| order by HourOfDay asc
"@

# Cross-Region Connection Analysis
# Joins network data with connection data to show Gateway Region → Session Host path with RTT
$kqlCrossRegionConnections = @"
WVDConnectionNetworkData
| where isnotnull(EstRoundTripTimeInMs) and EstRoundTripTimeInMs > 0
| join kind=inner (
    WVDConnections
    | where State == "Connected"
    | project CorrelationId, GatewayRegion, SessionHostName, UserName
) on CorrelationId
| where isnotempty(GatewayRegion) and isnotempty(SessionHostName)
| extend HostName = iif(SessionHostName contains "/", tostring(split(SessionHostName, '/')[1]), SessionHostName)
| summarize
    AvgRTTms = round(avg(EstRoundTripTimeInMs), 1),
    P50RTTms = round(percentile(EstRoundTripTimeInMs, 50), 1),
    P95RTTms = round(percentile(EstRoundTripTimeInMs, 95), 1),
    MaxRTTms = round(max(EstRoundTripTimeInMs), 1),
    AvgBandwidthKBps = round(avg(EstAvailableBandwidthKBps), 0),
    MinBandwidthKBps = round(min(EstAvailableBandwidthKBps), 0),
    Connections = dcount(CorrelationId),
    DistinctUsers = dcount(UserName)
    by GatewayRegion, SessionHostName = HostName
| order by P95RTTms desc
"@

# =========================================================
# Collection
# =========================================================
if ($resuming -and (Test-Checkpoint -CheckpointName "collection" -OutputFolder $outFolder)) {
  Write-ProgressSection -Section "Step 1: Collecting AVD Resources" -Status Skip -Message "Resuming: Using saved resource data from previous run"
}
else {
  # Estimate time based on subscription count (scales with environment size)
  $subsCount = SafeCount $SubscriptionIds
  $estimatedCollectionMinutes = [math]::Max(3, $subsCount * 3)  # ~3 min per subscription, minimum 3
  
  Write-ProgressSection -Section "Step 1: Collecting AVD Resources" -Status Start -EstimatedMinutes $estimatedCollectionMinutes -Message "Host pools, session hosts, VMs (time scales with environment size)"

# NIC cache: batch-fetch per resource group to avoid per-VM API calls (v3.0.0)
$nicCacheByRg = @{}

$subsProcessed = 0
foreach ($subId in $SubscriptionIds) {
  $subsProcessed++
  Set-AzContext -SubscriptionId $subId -TenantId $TenantId | Out-Null
  Write-ProgressSection -Section "Step 1: Collecting AVD Resources" -Status Progress -Current $subsProcessed -Total (SafeCount $SubscriptionIds) -Message "Subscription: $subId"

  # Host Pools
  $hpObjs = Get-AzWvdHostPool -ErrorAction SilentlyContinue
  foreach ($hpObj in SafeArray $hpObjs) {
    $hp = Resolve-ArmIdentity $hpObj "Microsoft.DesktopVirtualization/hostPools"
    if (-not $hp) { continue }

    $hpResourceGroup = $hp.ResourceGroup ?? (SafeProp $hpObj "ResourceGroupName")

    $hostPools.Add([PSCustomObject]@{
      SubscriptionId   = $subId
      ResourceGroup    = $hpResourceGroup
      HostPoolName     = $hp.Name
      HostPoolType     = (SafeProp $hpObj "HostPoolType")
      LoadBalancer     = (SafeProp $hpObj "LoadBalancerType")
      MaxSessions      = (SafeProp $hpObj "MaxSessionLimit")
      StartVMOnConnect = (SafeProp $hpObj "StartVMOnConnect")
      PreferredAppGroupType = (SafeProp $hpObj "PreferredAppGroupType")
      Location         = (SafeProp $hpObj "Location")
      ValidationEnv    = (SafeProp $hpObj "ValidationEnvironment")
      Id               = Get-ArmIdSafe $hp
    })

    $hpRg = $hpResourceGroup
    if (-not $hpRg) { continue }

    # Session Hosts
    $shObjs = Get-AzWvdSessionHost -ResourceGroupName $hpRg -HostPoolName $hp.Name -ErrorAction SilentlyContinue
    foreach ($shObj in SafeArray $shObjs) {
      $vmName = Normalize-SessionHostToVmName (SafeProp $shObj "Name")
      if (-not $vmName) { continue }

      $sessionHosts.Add([PSCustomObject]@{
        SubscriptionId  = $subId
        ResourceGroup   = $hpRg
        HostPoolName    = $hp.Name
        SessionHostName = $vmName
        SessionHostArmName = (SafeProp $shObj "Name")
        Status          = (SafeProp $shObj "Status")
        AllowNewSession = (SafeProp $shObj "AllowNewSession")
        ActiveSessions  = (SafeProp $shObj "Session")
        AssignedUser    = (SafeProp $shObj "AssignedUser")
        UpdateState     = (SafeProp $shObj "UpdateState")
        LastHeartBeat   = (SafeProp $shObj "LastHeartBeat")
      })

      # Backing VM
      $vm = $null
      try {
        $vm = Get-AzVM -ResourceGroupName $hpRg -Name $vmName -ErrorAction SilentlyContinue
        $vmStatus = Get-AzVM -ResourceGroupName $hpRg -Name $vmName -Status -ErrorAction SilentlyContinue
      } catch {}

      if (-not $vm) {
        $vmRes = Get-AzResource -ResourceType "Microsoft.Compute/virtualMachines" -Name $vmName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($vmRes) {
          try {
            $vm       = Get-AzVM -ResourceGroupName $vmRes.ResourceGroupName -Name $vmName -ErrorAction SilentlyContinue
            $vmStatus = Get-AzVM -ResourceGroupName $vmRes.ResourceGroupName -Name $vmName -Status -ErrorAction SilentlyContinue
          } catch {}
          $hpRgForVm = $vmRes.ResourceGroupName
        }
      } else {
        $hpRgForVm = $hpRg
      }

      if ($vm) {
        $osDiskType = $null
        try { $osDiskType = $vm.StorageProfile.OsDisk.ManagedDisk.StorageAccountType } catch {}

        $zones = $null
        try { $zones = ($vm.Zones -join ",") } catch {}

        $power = $null
        try { $power = (($vmStatus.Statuses | Where-Object { $_.Code -like "PowerState/*" }) | Select-Object -First 1).DisplayStatus } catch {}

        # Image reference (v3.0.0, enhanced v4.0.0)
        $imagePublisher = $null; $imageOffer = $null; $imageSku = $null; $imageVersion = $null
        $imageId = $null; $imageSource = "Unknown"
        try {
          $imgRef = $vm.StorageProfile.ImageReference
          $imagePublisher = $imgRef.Publisher
          $imageOffer = $imgRef.Offer
          $imageSku = $imgRef.Sku
          $imageVersion = $imgRef.ExactVersion
          if (-not $imageVersion) { $imageVersion = $imgRef.Version }
          $imageId = $imgRef.Id
          
          # Determine image source type
          if ($imageId -and $imageId -match '/galleries/') {
            $imageSource = "Gallery"
          } elseif ($imageId -and $imageId -match '/images/') {
            $imageSource = "ManagedImage"
          } elseif ($imagePublisher -and $imagePublisher -ne "") {
            $imageSource = "Marketplace"
          } elseif ($imageId) {
            $imageSource = "Custom"
          }
        } catch {}

        # OS disk type details (v3.0.0)
        $osDiskEphemeral = $false
        try { $osDiskEphemeral = ($vm.StorageProfile.OsDisk.DiffDiskSettings.Option -eq "Local") } catch {}

        # Security posture (v4.0.0)
        $securityType = $null; $secureBoot = $null; $vtpm = $null; $diskEncryption = $null
        try { $securityType = $vm.SecurityProfile.SecurityType } catch {}
        try { $secureBoot = $vm.SecurityProfile.UefiSettings.SecureBootEnabled } catch {}
        try { $vtpm = $vm.SecurityProfile.UefiSettings.VTpmEnabled } catch {}
        try { $diskEncryption = $vm.StorageProfile.OsDisk.ManagedDisk.SecurityProfile.DiskEncryptionSet.Id } catch {}
        # Check host encryption
        $hostEncryption = $false
        try { $hostEncryption = $vm.SecurityProfile.EncryptionAtHost } catch {}
        if (-not $hostEncryption) { try { $hostEncryption = $vm.AdditionalCapabilities.EncryptionAtHost } catch {} }

        # Accelerated Networking + Network topology check (v3.0.0/v4.0.0) - batch fetched per resource group
        $accelNetEnabled = $null
        $nicSubnetId = $null
        $nicNsgId = $null
        $nicPrivateIp = $null

        # Identity / Entra join detection (v4.1.0)
        $identityType = $null
        $hasAadExtension = $false
        try { $identityType = $vm.Identity.Type } catch {}  # SystemAssigned, UserAssigned, SystemAssignedUserAssigned, or null
        try {
          # Check for AADLoginForWindows or MicrosoftEntraID extension
          $vmExtensions = @($vm.Extensions | Where-Object { $_.VirtualMachineExtensionType -match 'AADLoginForWindows|AADLogin|MicrosoftEntraID' })
          if ($vmExtensions.Count -gt 0) { $hasAadExtension = $true }
        } catch {}

        try {
          $nicId = $vm.NetworkProfile.NetworkInterfaces[0].Id
          if ($nicId) {
            $nicRg = ($nicId -split '/')[4]
            # Batch-fetch NICs per resource group (cached)
            if (-not $nicCacheByRg.ContainsKey($nicRg)) {
              try {
                $rgNics = Get-AzNetworkInterface -ResourceGroupName $nicRg -ErrorAction SilentlyContinue
                $nicCacheByRg[$nicRg] = @{}
                foreach ($n in @($rgNics)) {
                  if ($n.Id) {
                    $nicCacheByRg[$nicRg][$n.Id] = @{
                      AccelNet  = $n.EnableAcceleratedNetworking
                      SubnetId  = if ($n.IpConfigurations -and $n.IpConfigurations[0].Subnet) { $n.IpConfigurations[0].Subnet.Id } else { $null }
                      NsgId     = if ($n.NetworkSecurityGroup) { $n.NetworkSecurityGroup.Id } else { $null }
                      PrivateIp = if ($n.IpConfigurations) { $n.IpConfigurations[0].PrivateIpAddress } else { $null }
                    }
                  }
                }
              } catch {
                $nicCacheByRg[$nicRg] = @{}
              }
            }
            $nicData = $nicCacheByRg[$nicRg][$nicId]
            if ($nicData) {
              $accelNetEnabled = $nicData.AccelNet
              $nicSubnetId = $nicData.SubnetId
              $nicNsgId = $nicData.NsgId
              $nicPrivateIp = $nicData.PrivateIp
            }
          }
        } catch {}

        $vms.Add([PSCustomObject]@{
          SubscriptionId  = $subId
          ResourceGroup   = $hpRgForVm
          HostPoolName    = $hp.Name
          SessionHostName = $vmName
          VMName          = $vm.Name
          VMId            = Get-ArmIdSafe $vm
          VMSize          = $vm.HardwareProfile.VmSize
          Region          = $vm.Location
          Zones           = $zones
          OSDiskType      = $osDiskType
          OSDiskEphemeral = $osDiskEphemeral
          DataDiskCount   = (SafeCount $vm.StorageProfile.DataDisks)
          PowerState      = $power
          ImagePublisher  = $imagePublisher
          ImageOffer      = $imageOffer
          ImageSku        = $imageSku
          ImageVersion    = $imageVersion
          ImageId         = $imageId
          ImageSource     = $imageSource
          AccelNetEnabled = $accelNetEnabled
          SubnetId        = $nicSubnetId
          NsgId           = $nicNsgId
          PrivateIp       = $nicPrivateIp
          SecurityType    = $securityType
          SecureBoot      = $secureBoot
          VTpm            = $vtpm
          HostEncryption  = $hostEncryption
          IdentityType    = $identityType
          HasAadExtension = $hasAadExtension
        })
      }
    }
  }

  # Autoscale
  foreach ($planRes in Get-ScalingPlansArm -SubscriptionId $subId) {
    Expand-ScalingPlanEvidence -PlanResource $planRes
  }

  # -------------------------
  # Scale Sets (VMSS)
  # -------------------------
  Write-Host "Collecting scale sets..." -ForegroundColor Cyan
  $vmssResources = Get-AzVmss -ErrorAction SilentlyContinue
  
  foreach ($vmssObj in SafeArray $vmssResources) {
    $vmssRg = $vmssObj.ResourceGroupName
    $vmssName = $vmssObj.Name
    
    # Determine if this VMSS is used for AVD by checking if any instances are session hosts
    $vmssInstObjs = Get-AzVmssVM -ResourceGroupName $vmssRg -VMScaleSetName $vmssName -ErrorAction SilentlyContinue
    $isAvdVmss = $false
    
    foreach ($inst in SafeArray $vmssInstObjs) {
      $instName = $inst.Name
      # Check if this instance name matches any session host
      $matchingSessionHost = $sessionHosts | Where-Object { 
        $_.SessionHostName -like "*$instName*" -or $instName -like "*$($_.SessionHostName)*"
      } | Select-Object -First 1
      
      if ($matchingSessionHost) {
        $isAvdVmss = $true
        break
      }
    }
    
    # Collect VMSS details
    $zones = $null
    try { $zones = ($vmssObj.Zones -join ",") } catch {}
    
    $osDiskType = $null
    try { $osDiskType = $vmssObj.VirtualMachineProfile.StorageProfile.OsDisk.ManagedDisk.StorageAccountType } catch {}
    
    $vmss.Add([PSCustomObject]@{
      SubscriptionId = $subId
      ResourceGroup = $vmssRg
      VMSSName = $vmssName
      VMSize = $vmssObj.Sku.Name
      Capacity = $vmssObj.Sku.Capacity
      Region = $vmssObj.Location
      Zones = $zones
      OSDiskType = $osDiskType
      UpgradePolicy = $vmssObj.UpgradePolicy.Mode
      Overprovision = $vmssObj.Overprovision
      IsAVD = $isAvdVmss
      Id = $vmssObj.Id
    })
    
    # Collect VMSS instances
    foreach ($inst in SafeArray $vmssInstObjs) {
      $instName = $inst.Name
      $instId = $inst.InstanceId
      
      # Try to match to session host
      $matchingSessionHost = $sessionHosts | Where-Object { 
        $_.SessionHostName -like "*$instName*" -or $instName -like "*$($_.SessionHostName)*"
      } | Select-Object -First 1
      
      $instZones = $null
      try { $instZones = ($inst.Zones -join ",") } catch {}
      
      $instPower = $null
      try {
        $instView = Get-AzVmssVM -ResourceGroupName $vmssRg -VMScaleSetName $vmssName -InstanceId $instId -InstanceView -ErrorAction SilentlyContinue
        $instPower = (($instView.Statuses | Where-Object { $_.Code -like "PowerState/*" }) | Select-Object -First 1).DisplayStatus
      } catch {}
      
      $vmssInstances.Add([PSCustomObject]@{
        SubscriptionId = $subId
        ResourceGroup = $vmssRg
        VMSSName = $vmssName
        InstanceId = $instId
        InstanceName = $instName
        VMSize = $vmssObj.Sku.Name
        Zones = $instZones
        PowerState = $instPower
        HostPoolName = if ($matchingSessionHost) { $matchingSessionHost.HostPoolName } else { $null }
        SessionHostName = if ($matchingSessionHost) { $matchingSessionHost.SessionHostName } else { $null }
        IsAVD = [bool]$matchingSessionHost
        ResourceId = $inst.Id
      })
      
      # Add to VMs collection if it's an AVD instance (for metrics collection)
      if ($matchingSessionHost) {
        # VMSS-level image and disk properties (v3.0.0)
        $vmssImagePublisher = $null; $vmssImageOffer = $null; $vmssImageSku = $null; $vmssImageVersion = $null
        $vmssImageId = $null; $vmssImageSource = "Unknown"
        try {
          $vmssImgRef = $vmssObj.VirtualMachineProfile.StorageProfile.ImageReference
          $vmssImagePublisher = $vmssImgRef.Publisher
          $vmssImageOffer = $vmssImgRef.Offer
          $vmssImageSku = $vmssImgRef.Sku
          $vmssImageVersion = $vmssImgRef.ExactVersion
          if (-not $vmssImageVersion) { $vmssImageVersion = $vmssImgRef.Version }
          $vmssImageId = $vmssImgRef.Id
          if ($vmssImageId -and $vmssImageId -match '/galleries/') { $vmssImageSource = "Gallery" }
          elseif ($vmssImageId -and $vmssImageId -match '/images/') { $vmssImageSource = "ManagedImage" }
          elseif ($vmssImagePublisher) { $vmssImageSource = "Marketplace" }
          elseif ($vmssImageId) { $vmssImageSource = "Custom" }
        } catch {}
        
        $vmssEphemeral = $false
        try { $vmssEphemeral = ($vmssObj.VirtualMachineProfile.StorageProfile.OsDisk.DiffDiskSettings.Option -eq "Local") } catch {}
        
        # VMSS AccelNet from network profile
        $vmssAccelNet = $null
        try { $vmssAccelNet = $vmssObj.VirtualMachineProfile.NetworkProfile.NetworkInterfaceConfigurations[0].EnableAcceleratedNetworking } catch {}
        
        # VMSS Security properties (v4.0.0)
        $vmssSecType = $null; $vmssSecBoot = $null; $vmssVtpm = $null
        try { $vmssSecType = $vmssObj.VirtualMachineProfile.SecurityProfile.SecurityType } catch {}
        try { $vmssSecBoot = $vmssObj.VirtualMachineProfile.SecurityProfile.UefiSettings.SecureBootEnabled } catch {}
        try { $vmssVtpm = $vmssObj.VirtualMachineProfile.SecurityProfile.UefiSettings.VTpmEnabled } catch {}
        
        $vms.Add([PSCustomObject]@{
          SubscriptionId = $subId
          ResourceGroup = $vmssRg
          HostPoolName = $matchingSessionHost.HostPoolName
          SessionHostName = $matchingSessionHost.SessionHostName
          VMName = $instName
          VMId = $inst.Id
          VMSize = $vmssObj.Sku.Name
          Region = $vmssObj.Location
          Zones = $instZones
          OSDiskType = $osDiskType
          OSDiskEphemeral = $vmssEphemeral
          DataDiskCount = 0
          PowerState = $instPower
          ImagePublisher = $vmssImagePublisher
          ImageOffer = $vmssImageOffer
          ImageSku = $vmssImageSku
          ImageVersion = $vmssImageVersion
          ImageId         = $vmssImageId
          ImageSource     = $vmssImageSource
          AccelNetEnabled = $vmssAccelNet
          SubnetId        = $null
          NsgId           = $null
          PrivateIp       = $null
          SecurityType    = $vmssSecType
          SecureBoot      = $vmssSecBoot
          VTpm            = $vmssVtpm
          HostEncryption  = $false
          IsVMSS = $true
          VMSSName = $vmssName
        })
      }
    }
  }
}

Write-ProgressSection -Section "Step 1: Collecting AVD Resources" -Status Complete -Message "Collected: $(SafeCount $hostPools) host pools, $(SafeCount $sessionHosts) session hosts, $(SafeCount $vms) VMs"

# Note if large environment
$vmCount = SafeCount $vms
if ($vmCount -gt 500) {
  Write-Host "  ℹ️  Large environment detected ($vmCount VMs) - collection times scale with size" -ForegroundColor Cyan
}

# Save checkpoint after collection
if (-not $resuming) {
    Save-Checkpoint -CheckpointName "collection" -OutputFolder $outFolder -Data @{
        hostPools = $hostPools
        sessionHosts = $sessionHosts
        vms = $vms
        vmss = $vmss
        vmssInstances = $vmssInstances
        scalingPlans = $scalingPlans
        scalingPlanAssignments = $scalingPlanAssignments
        scalingPlanSchedules = $scalingPlanSchedules
    }
}
}  # End of collection if block

# =========================================================
# QUICK SUMMARY MODE
# =========================================================
if ($QuickSummary) {
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                      QUICK SUMMARY MODE                               ║" -ForegroundColor Cyan
    Write-Host "║             Fast health check (no metrics collection)                ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    $vmCount = SafeCount $vms
    $hpCount = SafeCount $hostPools
    $shCount = SafeCount $sessionHosts
    
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "ENVIRONMENT SUMMARY" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Resources:" -ForegroundColor Yellow
    Write-Host "  • Host Pools: $hpCount" -ForegroundColor White
    Write-Host "  • Session Hosts: $shCount" -ForegroundColor White
    Write-Host "  • VMs: $vmCount" -ForegroundColor White
    
    if ((SafeCount $vmss) -gt 0) {
        Write-Host "  • Scale Sets: $(SafeCount $vmss) ($(SafeCount $vmssInstances) instances)" -ForegroundColor White
    }
    Write-Host ""
    
    # VM SKU distribution
    Write-Host "VM SKU Distribution:" -ForegroundColor Yellow
    $skuGroups = $vms | Group-Object VMSize | Sort-Object Count -Descending | Select-Object -First 5
    foreach ($sku in $skuGroups) {
        $pct = [math]::Round(($sku.Count / $vmCount) * 100, 1)
        Write-Host "  • $($sku.Name): $($sku.Count) ($pct%)" -ForegroundColor White
    }
    Write-Host ""
    
    # Zone distribution
    Write-Host "Zone Distribution:" -ForegroundColor Yellow
    $zonesEnabled = (SafeArray ($vms | Where-Object { $_.Zones })).Count
    $zonesPct = if ($vmCount -gt 0) { [math]::Round(($zonesEnabled / $vmCount) * 100, 0) } else { 0 }
    
    if ($zonesEnabled -eq 0) {
        Write-Host "  ⚠ NO VMs using availability zones (single point of failure!)" -ForegroundColor Red
    }
    elseif ($zonesPct -lt 50) {
        Write-Host "  ⚠ Only $zonesEnabled VMs ($zonesPct%) using zones" -ForegroundColor Yellow
    }
    else {
        Write-Host "  ✓ $zonesEnabled VMs ($zonesPct%) using zones" -ForegroundColor Green
    }
    Write-Host ""
    
    # Disk types
    Write-Host "Storage:" -ForegroundColor Yellow
    $premiumDisks = (SafeArray ($vms | Where-Object { $_.OSDiskType -like "*Premium*" })).Count
    $standardDisks = $vmCount - $premiumDisks
    Write-Host "  • Premium SSD: $premiumDisks VMs" -ForegroundColor White
    Write-Host "  • Standard: $standardDisks VMs" -ForegroundColor White
    if ($premiumDisks -gt ($vmCount * 0.5)) {
        Write-Host "  ⚠ Over 50% using Premium SSD (may be over-provisioned)" -ForegroundColor Yellow
    }
    Write-Host ""
    
    # Immediate findings
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "IMMEDIATE FINDINGS (Configuration Only)" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host ""
    
    $issues = 0
    $warnings = 0
    
    # Check host pools for zone distribution
    $hpNoZones = 0
    foreach ($hp in $hostPools) {
        $hpVMs = $vms | Where-Object { $_.HostPoolName -eq $hp.HostPoolName }
        $hpWithZones = (SafeArray ($hpVMs | Where-Object { $_.Zones })).Count
        if ($hpWithZones -eq 0 -and (SafeCount $hpVMs) -gt 0) {
            $hpNoZones++
        }
    }
    
    if ($hpNoZones -gt 0) {
        Write-Host "  ✗ $hpNoZones host pools have NO zone distribution (high risk)" -ForegroundColor Red
        $issues++
    }
    
    # Check for Start VM on Connect
    $hpNoStartVM = (SafeArray ($hostPools | Where-Object { -not $_.StartVMOnConnect })).Count
    if ($hpNoStartVM -gt 0) {
        Write-Host "  ⚠ $hpNoStartVM host pools without Start VM on Connect enabled" -ForegroundColor Yellow
        $warnings++
    }
    
    # Check VM sizes
    $oversizedVMs = (SafeArray ($vms | Where-Object { $_.VMSize -like "*64*" -or $_.VMSize -like "*96*" })).Count
    if ($oversizedVMs -gt 0) {
        Write-Host "  ⚠ $oversizedVMs very large VMs (64+ cores) - may be over-provisioned" -ForegroundColor Yellow
        $warnings++
    }
    
    if ($issues -eq 0 -and $warnings -eq 0) {
        Write-Host "  ✓ No immediate configuration issues found" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Yellow
    Write-Host "This was a QUICK SUMMARY - no metrics were collected." -ForegroundColor Yellow
    Write-Host "Run without -QuickSummary for detailed analysis with metrics." -ForegroundColor Yellow
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Yellow
    Write-Host ""
    
    exit 0
}

# =========================================================
# Metrics collection (OPTIMIZED: Parallel processing)
# =========================================================
if (-not $SkipAzureMonitorMetrics) {
  # Skip if resuming and we already have metrics
  if ($resuming -and (Test-Checkpoint -CheckpointName "metrics" -OutputFolder $outFolder)) {
    Write-ProgressSection -Section "Step 2: Collecting VM Metrics (Baseline)" -Status Skip -Message "Resuming: Using cached metrics from previous run"
  }
  else {
    $uniqueVmIds = $vms | Where-Object { $_.VMId } | Select-Object -ExpandProperty VMId -Unique
    $totalVMs = ($uniqueVmIds | Measure-Object).Count
  
  # Calculate estimated time — 1 API call per VM now instead of 12
  $estimatedMinutes = [math]::Ceiling($totalVMs / 50)  # ~50 VMs per minute with batched calls
  
  Write-ProgressSection -Section "Step 2: Collecting VM Metrics (Baseline)" -Status Start -EstimatedMinutes $estimatedMinutes -Message "$totalVMs VMs to analyze over last $MetricsLookbackDays days"
  
  Write-Host "  Using parallel processing (15 VMs at a time, 1 batched API call each)" -ForegroundColor Green
  Write-Host "  Estimated completion: $estimatedMinutes minutes" -ForegroundColor Gray
  Write-Host ""
  
  # Throttle design for production safety:
  # Azure Monitor allows 12,000 reads/hr/subscription. At ThrottleLimit 15 and 1 call/VM,
  # we sustain ~50-60 calls/min = ~3,000-3,600/hr — well within limits with headroom for
  # other tools and portal users sharing the same subscription quota.
  # Exponential backoff on 429s ensures we adapt if the sub is under heavy load.
  
  $parallelResults = $uniqueVmIds | ForEach-Object -Parallel {
      function SafeArray {
        param([object]$Obj)
        if ($null -eq $Obj) { return @() }
        return @($Obj)
      }
      
      $vmId = $_
      $lookback = $using:MetricsLookbackDays
      $timeGrain = $using:MetricsTimeGrainMinutes
      
      $results = @()
      $start = (Get-Date).AddDays(-$lookback)
      $end   = Get-Date
      $grain = New-TimeSpan -Minutes $timeGrain

      # Batched: CPU + Memory in one call, both aggregations
      # Disk and Network metrics were collected but never used by any analysis engine
      $metricNames = @("Percentage CPU", "Available Memory Bytes")
      $aggregations = @("Average", "Maximum")
      
      $attempt = 0
      $maxRetries = 4
      $success = $false
      
      while (-not $success -and $attempt -lt $maxRetries) {
        try {
          $metricObjects = Get-AzMetric `
            -ResourceId $vmId `
            -MetricName $metricNames `
            -Aggregation $aggregations `
            -StartTime $start `
            -EndTime $end `
            -TimeGrain $grain `
            -ErrorAction Stop `
            -WarningAction SilentlyContinue

          foreach ($metric in SafeArray $metricObjects) {
            $mName = $metric.Name.Value
            foreach ($ts in SafeArray $metric.Timeseries) {
              foreach ($pt in SafeArray $ts.Data) {
                foreach ($agg in $aggregations) {
                  $value = if ($agg -eq "Average") { $pt.Average } else { $pt.Maximum }
                  if ($null -ne $value) {
                    $results += [PSCustomObject]@{
                      VmId        = $vmId
                      Metric      = $mName
                      Aggregation = $agg
                      TimeStamp   = $pt.TimeStamp
                      Value       = $value
                    }
                  }
                }
              }
            }
          }
          $success = $true
        }
        catch {
          $attempt++
          $errMsg = $_.Exception.Message
          $isThrottled = $errMsg -match "Too Many Requests|429|throttl|rate limit"
          
          if ($isThrottled -and $attempt -lt $maxRetries) {
            # Exponential backoff: 15s, 45s, 135s — gives subscription time to recover
            $delay = 15 * [math]::Pow(3, $attempt - 1)
            Start-Sleep -Seconds $delay
          }
          elseif (-not $isThrottled) {
            break  # Non-throttle error, don't retry
          }
        }
      }
      
      return $results
    } -ThrottleLimit 15
    
    # Populate the List from parallel results
    foreach ($item in $parallelResults) { $vmMetrics.Add($item) }
    
    Write-ProgressSection -Section "Step 2: Collecting VM Metrics (Baseline)" -Status Complete -Message "Collected metrics for $totalVMs VMs"
  
  Write-Host "  Collected metrics for $totalVMs VMs" -ForegroundColor Cyan
  
  # Incident window metrics collection (OPTIMIZED)
  if ($IncludeIncidentWindowQueries) {
    Write-Host "Collecting Azure Monitor VM metrics (incident window)..." -ForegroundColor Cyan
    Write-Host "  Using parallel processing for incident window" -ForegroundColor Green
    
    $parallelIncidentResults = $uniqueVmIds | ForEach-Object -Parallel {
        function SafeArray {
          param([object]$Obj)
          if ($null -eq $Obj) { return @() }
          return @($Obj)
        }
        
        $vmId = $_
        $start = $using:IncidentWindowStart
        $end = $using:IncidentWindowEnd
        $grain = New-TimeSpan -Minutes $using:MetricsTimeGrainMinutes
        
        $results = @()
        $metricNames = @("Percentage CPU", "Available Memory Bytes")
        $aggregations = @("Average", "Maximum")
        
        $attempt = 0
        $maxRetries = 4
        $success = $false
        
        while (-not $success -and $attempt -lt $maxRetries) {
          try {
            $metricObjects = Get-AzMetric `
              -ResourceId $vmId `
              -MetricName $metricNames `
              -Aggregation $aggregations `
              -StartTime $start `
              -EndTime $end `
              -TimeGrain $grain `
              -ErrorAction Stop `
              -WarningAction SilentlyContinue
            
            foreach ($metric in SafeArray $metricObjects) {
              $mName = $metric.Name.Value
              foreach ($ts in SafeArray $metric.Timeseries) {
                foreach ($pt in SafeArray $ts.Data) {
                  foreach ($agg in $aggregations) {
                    $value = if ($agg -eq "Average") { $pt.Average } else { $pt.Maximum }
                    if ($null -ne $value) {
                      $results += [PSCustomObject]@{
                        VmId        = $vmId
                        Metric      = $mName
                        Aggregation = $agg
                        TimeStamp   = $pt.TimeStamp
                        Value       = $value
                      }
                    }
                  }
                }
              }
            }
            $success = $true
          }
          catch {
            $attempt++
            $errMsg = $_.Exception.Message
            $isThrottled = $errMsg -match "Too Many Requests|429|throttl|rate limit"
            
            if ($isThrottled -and $attempt -lt $maxRetries) {
              $delay = 15 * [math]::Pow(3, $attempt - 1)
              Start-Sleep -Seconds $delay
            }
            elseif (-not $isThrottled) {
              break
            }
          }
        }
        
        return $results
      } -ThrottleLimit 15
      
      # Populate the List from parallel results
      foreach ($item in $parallelIncidentResults) { $vmMetricsIncident.Add($item) }
  }
  }  # Close the else block for resume check
} else {
  Write-ProgressSection -Section "Step 2: Collecting VM Metrics" -Status Skip -Message "-SkipAzureMonitorMetrics parameter specified (right-sizing will be limited)"
}

# Save metrics checkpoint (after baseline and incident window collection)
if (-not $SkipAzureMonitorMetrics -and -not $resuming) {
    Save-Checkpoint -CheckpointName "metrics" -OutputFolder $outFolder -Data @{
        vmMetrics = $vmMetrics
        vmMetricsIncident = $vmMetricsIncident
    }
}

# =========================================================
# Log Analytics queries
# =========================================================
if (-not $SkipLogAnalyticsQueries -and (SafeCount $LogAnalyticsWorkspaceResourceIds) -gt 0) {
  Write-Host "Executing Log Analytics queries..." -ForegroundColor Cyan

  $now = Get-Date
  $curStart = $now.AddDays(-$MetricsLookbackDays)

  foreach ($wsId in SafeArray $LogAnalyticsWorkspaceResourceIds) {
    Write-Host "  Querying workspace $wsId" -ForegroundColor Gray
    
    # Current window queries
    $kqlQueries = @(
      @{ Label = "CurrentWindow_TableDiscovery";            Query = $kqlTableDiscovery }
      @{ Label = "CurrentWindow_WVDConnections";            Query = $kqlWvdConnections }
      @{ Label = "CurrentWindow_WVDShortpathUsage";         Query = $kqlShortpathUsage }
      @{ Label = "CurrentWindow_WVDPeakConcurrency";        Query = $kqlPeakConcurrency }
      @{ Label = "CurrentWindow_WVDAutoscaleActivity";      Query = $kqlAutoscaleActivity }
      @{ Label = "CurrentWindow_WVDAutoscaleDetailed";      Query = $kqlAutoscaleDetailedActivity }
      @{ Label = "CurrentWindow_SessionDuration";           Query = $kqlSessionDuration }
      @{ Label = "CurrentWindow_ProfileLoadPerformance";    Query = $kqlProfileLoadPerformance }
      @{ Label = "CurrentWindow_ConnectionQuality";         Query = $kqlConnectionQuality }
      @{ Label = "CurrentWindow_ConnectionQualityByRegion"; Query = $kqlConnectionQualityByRegion }
      @{ Label = "CurrentWindow_ConnectionErrors";          Query = $kqlConnectionErrors }
      @{ Label = "CurrentWindow_Disconnects";               Query = $kqlDisconnects }
      @{ Label = "CurrentWindow_DisconnectReasons";          Query = $kqlDisconnectReasons }
      @{ Label = "CurrentWindow_DisconnectsByHost";          Query = $kqlDisconnectsByHost }
      @{ Label = "CurrentWindow_HourlyConcurrency";         Query = $kqlHourlyConcurrency }
      @{ Label = "CurrentWindow_CrossRegionConnections";   Query = $kqlCrossRegionConnections }
      @{ Label = "CurrentWindow_LoginTime";                 Query = $kqlLoginTime }
      @{ Label = "CurrentWindow_ConnectionSuccessRate";     Query = $kqlConnectionSuccessRate }
    )
    
    Write-Host "    Running $($kqlQueries.Count) KQL queries..." -ForegroundColor Gray
    foreach ($kq in $kqlQueries) {
      $beforeCount = $laResults.Count
      foreach ($r in @(Invoke-LaQuery -WorkspaceResourceId $wsId -Label $kq.Label -Query $kq.Query -StartTime $curStart -EndTime $now)) {
        $laResults.Add($r)
      }
      $added = $laResults.Count - $beforeCount
      $lastRow = if ($added -gt 0) { $laResults[$laResults.Count - 1] } else { $null }
      if ($lastRow -and $lastRow.QueryName -eq "Meta") {
        $errDetail = if ($lastRow.PSObject.Properties['Error']) { ": $($lastRow.Error.Substring(0, [math]::Min(120, $lastRow.Error.Length)))" } else { "" }
        Write-Host "      $($kq.Label): ⚠ $($lastRow.Status)$errDetail" -ForegroundColor Yellow
      } elseif ($added -gt 0) {
        Write-Host "      $($kq.Label): ✓ $added rows" -ForegroundColor Green
      } else {
        Write-Host "      $($kq.Label): ⚠ 0 rows returned" -ForegroundColor Yellow
      }
    }
    
    if ($IncludeIncidentWindowQueries) {
      Write-Host "    Running incident window queries..." -ForegroundColor Gray
      $incidentQueries = @(
        @{ Label = "IncidentWindow_WVDConnections";            Query = $kqlWvdConnections }
        @{ Label = "IncidentWindow_WVDPeakConcurrency";        Query = $kqlPeakConcurrency }
        @{ Label = "IncidentWindow_ProfileLoadPerformance";    Query = $kqlProfileLoadPerformance }
        @{ Label = "IncidentWindow_ConnectionErrors";          Query = $kqlConnectionErrors }
        @{ Label = "IncidentWindow_ConnectionQuality";         Query = $kqlConnectionQuality }
      )
      foreach ($kq in $incidentQueries) {
        foreach ($r in @(Invoke-LaQuery -WorkspaceResourceId $wsId -Label $kq.Label -Query $kq.Query -StartTime $IncidentWindowStart -EndTime $IncidentWindowEnd)) {
          $laResults.Add($r)
        }
      }
    }
    
    Write-Host "    ✓ Completed queries ($($laResults.Count) total rows)" -ForegroundColor Green
  }
} else {
  if ($SkipLogAnalyticsQueries) {
    Write-ProgressSection -Section "Log Analytics Queries" -Status Skip -Message "-SkipLogAnalyticsQueries parameter specified"
  } else {
    Write-ProgressSection -Section "Log Analytics Queries" -Status Skip -Message "No Log Analytics workspace IDs provided"
  }
}

# =========================================================
# Enhanced Analysis: VM Right-Sizing
# =========================================================
$vmCount = SafeCount $vms
$estimatedRightSizingMinutes = [math]::Ceiling($vmCount / 200)  # ~200 VMs per minute
Write-ProgressSection -Section "Step 3: Right-Sizing Analysis" -Status Start -EstimatedMinutes $estimatedRightSizingMinutes -Message "Analyzing $vmCount VMs for optimization opportunities"

# OPTIMIZATION: Pre-group metrics by VM ID to avoid repeated Where-Object on large arrays
Write-Host "  [1/3] Indexing metrics by VM..." -ForegroundColor Gray

# Build a script-level lookup for VM memory from size name
# (mirrors the $vmSizes table inside Get-RightSizedVmRecommendation)
$vmMemoryLookup = @{
  "Standard_D2s_v4"=8; "Standard_D4s_v4"=16; "Standard_D8s_v4"=32; "Standard_D16s_v4"=64; "Standard_D32s_v4"=128
  "Standard_D2s_v5"=8; "Standard_D4s_v5"=16; "Standard_D8s_v5"=32; "Standard_D16s_v5"=64; "Standard_D32s_v5"=128; "Standard_D48s_v5"=192; "Standard_D64s_v5"=256
  "Standard_D2ads_v5"=8; "Standard_D4ads_v5"=16; "Standard_D8ads_v5"=32; "Standard_D16ads_v5"=64; "Standard_D32ads_v5"=128
  "Standard_D2s_v6"=8; "Standard_D4s_v6"=16; "Standard_D8s_v6"=32; "Standard_D16s_v6"=64; "Standard_D32s_v6"=128
  "Standard_E2s_v4"=16; "Standard_E4s_v4"=32; "Standard_E8s_v4"=64; "Standard_E16s_v4"=128; "Standard_E32s_v4"=256
  "Standard_E2s_v5"=16; "Standard_E4s_v5"=32; "Standard_E8s_v5"=64; "Standard_E16s_v5"=128; "Standard_E32s_v5"=256; "Standard_E48s_v5"=384; "Standard_E64s_v5"=512
  "Standard_E2ads_v5"=16; "Standard_E4ads_v5"=32; "Standard_E8ads_v5"=64; "Standard_E16ads_v5"=128; "Standard_E32ads_v5"=256
  "Standard_E2s_v6"=16; "Standard_E4s_v6"=32; "Standard_E8s_v6"=64; "Standard_E16s_v6"=128; "Standard_E32s_v6"=256
  "Standard_E2ads_v6"=16; "Standard_E4ads_v6"=32; "Standard_E8ads_v6"=64; "Standard_E16ads_v6"=128; "Standard_E32ads_v6"=256; "Standard_E48ads_v6"=384; "Standard_E64ads_v6"=512
  "Standard_B2s"=4; "Standard_B4ms"=16; "Standard_B8ms"=32; "Standard_B12ms"=48; "Standard_B16ms"=64
  "Standard_D2s_v3"=8; "Standard_D4s_v3"=16; "Standard_D8s_v3"=32; "Standard_D16s_v3"=64
  "Standard_E2s_v3"=16; "Standard_E4s_v3"=32; "Standard_E8s_v3"=64; "Standard_E16s_v3"=128
}
$metricsByVm = @{}
foreach ($metric in $vmMetrics) {
  $vmId = $metric.VmId
  if (-not $vmId) { continue }
  
  if (-not $metricsByVm.ContainsKey($vmId)) {
    $metricsByVm[$vmId] = @{
      CpuAvg = @()
      CpuMax = @()
      Mem = @()
    }
  }
  
  if ($metric.Metric -eq "Percentage CPU" -and $metric.Aggregation -eq "Average") {
    $metricsByVm[$vmId].CpuAvg += $metric
  }
  elseif ($metric.Metric -eq "Percentage CPU" -and $metric.Aggregation -eq "Maximum") {
    $metricsByVm[$vmId].CpuMax += $metric
  }
  elseif ($metric.Metric -eq "Available Memory Bytes") {
    $metricsByVm[$vmId].Mem += $metric
  }
}
Write-Host "  ✓ Indexed $($metricsByVm.Keys.Count) VMs" -ForegroundColor Green

# Pre-group session hosts by VM name
Write-Host "  [2/3] Indexing session hosts by VM..." -ForegroundColor Gray
$sessionHostsByVm = @{}
foreach ($sh in $sessionHosts) {
  $vmName = $sh.SessionHostName
  if (-not $vmName) { continue }
  
  if (-not $sessionHostsByVm.ContainsKey($vmName)) {
    $sessionHostsByVm[$vmName] = @()
  }
  $sessionHostsByVm[$vmName] += $sh
}

# Build host pool type, app group, and load balancer lookups (used by right-sizing, storage, scaling, W365)
$hpTypeLookup = @{}
$hpAppGroupLookup = @{}
$hpLoadBalancerLookup = @{}
$hpMaxSessionsLookup = @{}
foreach ($hp in $hostPools) {
  if ($hp.HostPoolName) {
    $hpTypeLookup[$hp.HostPoolName] = $hp.HostPoolType
    $hpAppGroupLookup[$hp.HostPoolName] = $hp.PreferredAppGroupType    # Desktop or RailApplications
    $hpLoadBalancerLookup[$hp.HostPoolName] = $hp.LoadBalancer          # BreadthFirst, DepthFirst, Persistent
    $hpMaxSessionsLookup[$hp.HostPoolName] = if ($hp.MaxSessions) { [int]$hp.MaxSessions } else { 0 }
  }
}

Write-Host "  [3/3] Analyzing VMs for right-sizing recommendations..." -ForegroundColor Gray
$progressCount = 0
$startTime = Get-Date

foreach ($vm in $vms) {
  $vmId = $vm.VMId
  if (-not $vmId) { continue }
  
  $progressCount++
  if ($progressCount % 100 -eq 0) {
    $elapsed = (Get-Date) - $startTime
    $rate = $progressCount / $elapsed.TotalMinutes
    $remaining = ($vmCount - $progressCount) / $rate
    Write-ProgressSection -Section "Step 3: Right-Sizing Analysis" -Status Progress -Current $progressCount -Total $vmCount -Message "Rate: $([math]::Round($rate, 0)) VMs/min | ETA: $([math]::Round($remaining, 0)) min"
  }

  # Get pre-grouped metrics
  $vmMetricsData = $metricsByVm[$vmId]
  
  if ($vmMetricsData) {
    $cpuAvgMeasure = $vmMetricsData.CpuAvg | Measure-Object Value -Average
    $cpuMaxMeasure = $vmMetricsData.CpuMax | Measure-Object Value -Maximum
    $memAvgMeasure = $vmMetricsData.Mem | Measure-Object Value -Average
    $memMinMeasure = $vmMetricsData.Mem | Measure-Object Value -Minimum
    
    $avgCpu = (SafeMeasure $cpuAvgMeasure 'Average') ?? 0
    $peakCpu = (SafeMeasure $cpuMaxMeasure 'Maximum') ?? 0
    $avgFreeMem = (SafeMeasure $memAvgMeasure 'Average') ?? 0
    $minFreeMem = (SafeMeasure $memMinMeasure 'Minimum') ?? 0
  }
  else {
    $avgCpu = 0
    $peakCpu = 0
    $avgFreeMem = 0
    $minFreeMem = 0
  }

  # Estimate used memory (lookup actual memory from VM size, fallback to 8GB)
  $totalMemGB = $vmMemoryLookup[$vm.VMSize] ?? 8
  $avgMemUsedGB = $totalMemGB - ($avgFreeMem / 1GB)
  $peakMemUsedGB = $totalMemGB - ($minFreeMem / 1GB)

  # Get pre-grouped sessions
  $vmSessions = $sessionHostsByVm[$vm.VMName]
  if ($vmSessions) {
    $sessionsAvgMeasure = $vmSessions | Measure-Object ActiveSessions -Average
    $sessionsMaxMeasure = $vmSessions | Measure-Object ActiveSessions -Maximum
    $avgSessions = if ((SafeMeasure $sessionsAvgMeasure 'Average')) { (SafeMeasure $sessionsAvgMeasure 'Average') } else { 0 }
    $peakSessions = if ((SafeMeasure $sessionsMaxMeasure 'Maximum')) { (SafeMeasure $sessionsMaxMeasure 'Maximum') } else { 0 }
  }
  else {
    $avgSessions = 0
    $peakSessions = 0
  }

  # Get recommendation
  $poolAppGroup = $hpAppGroupLookup[$vm.HostPoolName]
  $poolLoadBalancer = $hpLoadBalancerLookup[$vm.HostPoolName]
  $poolType = $hpTypeLookup[$vm.HostPoolName]
  
  $recommendation = Get-RightSizedVmRecommendation `
    -CurrentSize $vm.VMSize `
    -AvgCpu $avgCpu `
    -PeakCpu $peakCpu `
    -AvgMemoryUsedGB $avgMemUsedGB `
    -PeakMemoryUsedGB $peakMemUsedGB `
    -AvgSessions $avgSessions `
    -PeakSessions $peakSessions `
    -LookbackDays $MetricsLookbackDays `
    -AvgFreeMem $avgFreeMem `
    -MinFreeMem $minFreeMem `
    -HostPoolType $poolType `
    -AppGroupType $poolAppGroup `
    -LoadBalancer $poolLoadBalancer

  # Calculate potential savings
  $currentCost = Get-EstimatedVmCostPerHour -VmSize $vm.VMSize -Region $vm.Region
  $recommendedCost = if ($recommendation -and $recommendation.Recommendation -and $recommendation.Recommendation -ne "Keep Current" -and $recommendation.Recommendation -ne "Unknown") {
    Get-EstimatedVmCostPerHour -VmSize $recommendation.Recommendation -Region $vm.Region
  } else { $currentCost }

  $monthlySavings = if ($currentCost -and $recommendedCost) {
    ($currentCost - $recommendedCost) * 730  # Hours per month
  } else { $null }

  $vmRightSizing.Add([PSCustomObject]@{
    SubscriptionId = $vm.SubscriptionId
    ResourceGroup = $vm.ResourceGroup
    HostPoolName = $vm.HostPoolName
    HostPoolType = $poolType
    AppGroupType = if ($poolAppGroup -eq "RailApplications") { "RemoteApp" } elseif ($poolAppGroup) { $poolAppGroup } else { "" }
    VMName = $vm.VMName
    CurrentSize = $vm.VMSize
    CurrentvCPU = if ($recommendation -and $recommendation.CurrentvCPU) { $recommendation.CurrentvCPU } else { $null }
    CurrentMemoryGB = if ($recommendation -and $recommendation.CurrentMemoryGB) { $recommendation.CurrentMemoryGB } else { $null }
    AvgCPU = [math]::Round($avgCpu, 1)
    PeakCPU = [math]::Round($peakCpu, 1)
    AvgMemoryUsedGB = [math]::Round($avgMemUsedGB, 1)
    PeakMemoryUsedGB = [math]::Round($peakMemUsedGB, 1)
    AvgSessions = [math]::Round($avgSessions, 1)
    PeakSessions = $peakSessions
    SessionsPerVCPU = if ($recommendation -and $recommendation.SessionsPerVCPU) { $recommendation.SessionsPerVCPU } else { $null }
    MemoryPerSessionGB = if ($recommendation -and $recommendation.MemoryPerSessionGB) { $recommendation.MemoryPerSessionGB } else { $null }
    RecommendedSize = if ($recommendation -and $recommendation.Recommendation) { $recommendation.Recommendation } else { "Unknown" }
    Reason = if ($recommendation -and $recommendation.Reason) { $recommendation.Reason } else { "No recommendation available" }
    Confidence = if ($recommendation -and $recommendation.Confidence) { $recommendation.Confidence } else { "Low" }
    EvidenceScore = if ($recommendation -and $null -ne $recommendation.EvidenceScore) { $recommendation.EvidenceScore } else { 0 }
    EvidenceSignals = if ($recommendation -and $recommendation.EvidenceSignals) { $recommendation.EvidenceSignals } else { "None" }
    EstimatedMonthlySavings = if ($monthlySavings) { [math]::Round($monthlySavings, 2) } else { "N/A" }
  })
}

Write-ProgressSection -Section "Step 3: Right-Sizing Analysis" -Status Complete -Message "Analyzed $(SafeCount $vms) VMs | Recommendations: $(SafeCount $vmRightSizing)"

# =========================================================
# Enhanced Analysis: Zone Resiliency
# =========================================================
Write-ProgressSection -Section "Step 4: Zone Resiliency Analysis" -Status Start -EstimatedMinutes 1 -Message "Evaluating high-availability posture"

$hostPoolNames = $vms | Select-Object -ExpandProperty HostPoolName -Unique

foreach ($hpName in $hostPoolNames) {
  $hpVMs = $vms | Where-Object { $_.HostPoolName -eq $hpName }
  $resiliency = Get-ZoneResiliencyScore -HostPoolName $hpName -VMs $hpVMs
  $zoneResiliency.Add($resiliency)
}

Write-ProgressSection -Section "Step 4: Zone Resiliency Analysis" -Status Complete -Message "Evaluated $(SafeCount $hostPoolNames) host pools"

# =========================================================
# Enhanced Analysis: Session Host Health (v3.0.0)
# =========================================================
Write-Host "Analyzing session host health..." -ForegroundColor Cyan

$sessionHostHealth = [System.Collections.Generic.List[object]]::new()

foreach ($sh in $sessionHosts) {
  $drainMode = ($sh.AllowNewSession -eq $false)
  $lastHeartbeat = $sh.LastHeartBeat
  $heartbeatAge = $null
  $staleHeartbeat = $false
  
  if ($lastHeartbeat) {
    try {
      $hbDate = [datetime]$lastHeartbeat
      $heartbeatAge = [math]::Round(((Get-Date) - $hbDate).TotalHours, 1)
      $staleHeartbeat = ($heartbeatAge -gt 24)
    } catch {}
  }
  
  $stuckInDrain = ($drainMode -and $staleHeartbeat)
  
  $status = $sh.Status
  $isUnavailable = ($status -eq "Unavailable" -or $status -eq "NeedsAssistance" -or $status -eq "Shutdown")
  
  $finding = "Healthy"
  $severity = "Normal"
  $remediation = ""
  
  if ($stuckInDrain) {
    $finding = "Stuck in drain mode (drained $([math]::Round($heartbeatAge, 0))h ago, no heartbeat)"
    $severity = "High"
    $remediation = "1) Check if VM is running in Azure Portal → restart if deallocated. 2) If running, RDP to the VM and check RDAgent service (RDAgentBootLoader). 3) If unresponsive, restart the VM. 4) If still stuck after restart, remove and re-register the session host. 5) For pooled hosts, consider reimaging instead."
  }
  elseif ($drainMode) {
    $finding = "Drain mode active — not accepting new sessions"
    $severity = "Medium"
    $remediation = "Drain mode is often intentional (maintenance, patching). If not expected: 1) Check if a scaling plan set drain mode. 2) Undrain via Portal → Host Pool → Session Hosts → select host → Allow new sessions. 3) Verify no active sessions need to complete first."
  }
  elseif ($isUnavailable) {
    $finding = "Unavailable ($status)"
    $severity = "High"
    $remediation = "1) Check VM power state in Azure Portal. 2) If Shutdown/Deallocated, start the VM. 3) If running but NeedsAssistance, RDP to VM and restart the RDAgentBootLoader service. 4) Check Windows Event Viewer → Application/System logs for crash or blue screen events. 5) Verify domain join and DNS resolution. 6) If persistent, reinstall the RD Agent from https://aka.ms/AVDAgent."
  }
  elseif ($staleHeartbeat) {
    $finding = "Stale heartbeat ($([math]::Round($heartbeatAge, 0))h old) — VM may be frozen or RDAgent stopped"
    $severity = "Medium"
    $remediation = "1) Verify VM is running (not stopped/deallocated). 2) RDP to the VM and check if RDAgentBootLoader and RDAgent services are running. 3) Restart both services: Restart-Service RDAgentBootLoader -Force. 4) Check for Windows updates pending reboot. 5) Review Event Viewer for RDAgent errors under Application log, source 'RemoteDesktopServices'. 6) If services won't start, reinstall from https://aka.ms/AVDAgent."
  }
  
  $sessionHostHealth.Add([PSCustomObject]@{
    SubscriptionId  = $sh.SubscriptionId
    HostPoolName    = $sh.HostPoolName
    SessionHostName = $sh.SessionHostName
    Status          = $status
    AllowNewSession = $sh.AllowNewSession
    ActiveSessions  = $sh.ActiveSessions
    LastHeartBeat   = $lastHeartbeat
    HeartbeatAgeHrs = $heartbeatAge
    DrainMode       = $drainMode
    StuckInDrain    = $stuckInDrain
    Finding         = $finding
    Severity        = $severity
    Remediation     = $remediation
  })
}

$drainedHosts = ($sessionHostHealth | Where-Object { $_.DrainMode } | Measure-Object).Count
$stuckHosts = ($sessionHostHealth | Where-Object { $_.StuckInDrain } | Measure-Object).Count
$unavailableHosts = ($sessionHostHealth | Where-Object { $_.Severity -eq "High" } | Measure-Object).Count
Write-Host "  Session host health: $drainedHosts drained, $stuckHosts stuck, $unavailableHosts with issues" -ForegroundColor $(if ($stuckHosts -gt 0) { "Yellow" } else { "Green" })

# =========================================================
# Enhanced Analysis: OS Disk & Storage Optimization (v3.0.0)
# =========================================================
Write-Host "Analyzing OS disk and storage configuration..." -ForegroundColor Cyan

$storageFindingsList = [System.Collections.Generic.List[object]]::new()

foreach ($vm in $vms) {
  $findings = @()
  $poolType = $hpTypeLookup[$vm.HostPoolName]
  $isPooled = ($poolType -eq "Pooled")
  $appGroup = $hpAppGroupLookup[$vm.HostPoolName]
  $isRemoteAppVM = ($appGroup -eq "RailApplications")
  
  # Premium SSD on pooled hosts (overprovisioned)
  if ($isPooled -and $vm.OSDiskType -match "Premium") {
    if ($isRemoteAppVM) {
      $findings += "Premium SSD on RemoteApp host — ephemeral OS disk is strongly recommended (stateless app delivery, fastest reimage)"
    } else {
      $findings += "Premium SSD on pooled desktop — consider Standard SSD or Ephemeral for cost savings"
    }
  }
  
  # Ephemeral disk check (best practice for pooled)
  if ($isPooled -and -not $vm.OSDiskEphemeral) {
    if ($isRemoteAppVM) {
      $findings += "Non-ephemeral OS disk on RemoteApp host — ephemeral disks are ideal for stateless published app delivery (zero disk cost, faster reimage, no stale state)"
    } else {
      $findings += "Non-ephemeral OS disk on pooled desktop — ephemeral disks reduce cost and improve performance"
    }
  }
  
  # Premium SSD on personal is fine, but StandardHDD is too slow
  if ($vm.OSDiskType -eq "Standard_LRS") {
    $findings += "Standard HDD detected - upgrade to Standard SSD (Premium_LRS) for better IOPS"
  }
  
  if (@($findings).Count -eq 0) {
    $findings += "Optimal"
  }
  
  $storageFindingsList.Add([PSCustomObject]@{
    VMName          = $vm.VMName
    HostPoolName    = $vm.HostPoolName
    HostPoolType    = $poolType
    AppGroupType    = if ($appGroup -eq "RailApplications") { "RemoteApp" } elseif ($appGroup) { $appGroup } else { "" }
    OSDiskType      = $vm.OSDiskType
    OSDiskEphemeral = $vm.OSDiskEphemeral
    DataDiskCount   = $vm.DataDiskCount
    Findings        = ($findings -join "; ")
  })
}

$premiumOnPooled = ($storageFindingsList | Where-Object { $_.Findings -match "Premium SSD on" -and $_.HostPoolType -eq "Pooled" } | Measure-Object).Count
$nonEphemeral = ($storageFindingsList | Where-Object { $_.Findings -match "Non-ephemeral" } | Measure-Object).Count
Write-Host "  Storage findings: $premiumOnPooled premium-on-pooled, $nonEphemeral non-ephemeral" -ForegroundColor $(if ($premiumOnPooled -gt 0) { "Yellow" } else { "Green" })

# =========================================================
# Enhanced Analysis: Accelerated Networking (v3.0.0)
# =========================================================
Write-Host "Checking Accelerated Networking configuration..." -ForegroundColor Cyan

$accelNetFindings = [System.Collections.Generic.List[object]]::new()

# vCPU lookup for AccelNet eligibility (4+ vCPU generally supports AccelNet)
$vmVcpuLookup = @{
  "Standard_D2s_v4"=2; "Standard_D4s_v4"=4; "Standard_D8s_v4"=8; "Standard_D16s_v4"=16; "Standard_D32s_v4"=32
  "Standard_D2s_v5"=2; "Standard_D4s_v5"=4; "Standard_D8s_v5"=8; "Standard_D16s_v5"=16; "Standard_D32s_v5"=32; "Standard_D48s_v5"=48; "Standard_D64s_v5"=64
  "Standard_D2ads_v5"=2; "Standard_D4ads_v5"=4; "Standard_D8ads_v5"=8; "Standard_D16ads_v5"=16; "Standard_D32ads_v5"=32
  "Standard_D2s_v6"=2; "Standard_D4s_v6"=4; "Standard_D8s_v6"=8; "Standard_D16s_v6"=16; "Standard_D32s_v6"=32
  "Standard_E2s_v4"=2; "Standard_E4s_v4"=4; "Standard_E8s_v4"=8; "Standard_E16s_v4"=16; "Standard_E32s_v4"=32
  "Standard_E2s_v5"=2; "Standard_E4s_v5"=4; "Standard_E8s_v5"=8; "Standard_E16s_v5"=16; "Standard_E32s_v5"=32; "Standard_E48s_v5"=48; "Standard_E64s_v5"=64
  "Standard_E2ads_v5"=2; "Standard_E4ads_v5"=4; "Standard_E8ads_v5"=8; "Standard_E16ads_v5"=16; "Standard_E32ads_v5"=32
  "Standard_E2s_v6"=2; "Standard_E4s_v6"=4; "Standard_E8s_v6"=8; "Standard_E16s_v6"=16; "Standard_E32s_v6"=32
  "Standard_E2ads_v6"=2; "Standard_E4ads_v6"=4; "Standard_E8ads_v6"=8; "Standard_E16ads_v6"=16; "Standard_E32ads_v6"=32; "Standard_E48ads_v6"=48; "Standard_E64ads_v6"=64
  "Standard_B2s"=2; "Standard_B4ms"=4; "Standard_B8ms"=8; "Standard_B12ms"=12; "Standard_B16ms"=16
}

foreach ($vm in $vms) {
  $vcpus = $vmVcpuLookup[$vm.VMSize] ?? 0
  $eligible = ($vcpus -ge 4)
  $enabled = ($vm.AccelNetEnabled -eq $true)
  
  $finding = "N/A"
  $severity = "Normal"
  
  if ($eligible -and -not $enabled) {
    $finding = "Eligible but NOT enabled - enable for better network throughput and lower latency"
    $severity = "Medium"
  }
  elseif ($eligible -and $enabled) {
    $finding = "Enabled"
  }
  elseif (-not $eligible) {
    $finding = "Not eligible (requires 4+ vCPU SKU)"
  }
  
  $accelNetFindings.Add([PSCustomObject]@{
    VMName          = $vm.VMName
    VMSize          = $vm.VMSize
    vCPUs           = $vcpus
    AccelNetEnabled = $enabled
    Eligible        = $eligible
    Finding         = $finding
    Severity        = $severity
  })
}

$eligibleNotEnabled = ($accelNetFindings | Where-Object { $_.Eligible -and -not $_.AccelNetEnabled } | Measure-Object).Count
Write-Host "  AccelNet: $eligibleNotEnabled VMs eligible but not enabled" -ForegroundColor $(if ($eligibleNotEnabled -gt 0) { "Yellow" } else { "Green" })

# =========================================================
# Enhanced Analysis: Image & Golden Image Assessment (v4.0.0)
# =========================================================
Write-Host "Analyzing image versions, sources, and golden image maturity..." -ForegroundColor Cyan

$imageAnalysis = [System.Collections.Generic.List[object]]::new()

# --- Per-image group analysis ---
$marketplaceVms = @($vms | Where-Object { $_.PSObject.Properties['ImageSource'] -and $_.ImageSource -eq "Marketplace" })
$galleryVms = @($vms | Where-Object { $_.PSObject.Properties['ImageSource'] -and $_.ImageSource -eq "Gallery" })
$managedImageVms = @($vms | Where-Object { $_.PSObject.Properties['ImageSource'] -and $_.ImageSource -eq "ManagedImage" })
$customVms = @($vms | Where-Object { $_.PSObject.Properties['ImageSource'] -and $_.ImageSource -eq "Custom" })
$unknownSourceVms = @($vms | Where-Object { -not $_.PSObject.Properties['ImageSource'] -or $_.ImageSource -eq "Unknown" })

# Group marketplace images by publisher/offer/sku
$imageGroups = $vms | Where-Object { $_.ImagePublisher } | Group-Object ImagePublisher, ImageOffer, ImageSku

foreach ($group in $imageGroups) {
  $groupVms = @($group.Group)
  $versions = @($groupVms | Select-Object -ExpandProperty ImageVersion -Unique)
  $vmCount_img = $groupVms.Count
  $sampleVm = $groupVms[0]
  $hostPools_img = @($groupVms | Select-Object -ExpandProperty HostPoolName -Unique)
  
  # Detect OS generation from SKU name
  $skuName = $sampleVm.ImageSku ?? ""
  $osGeneration = "Unknown"
  $isAVDOptimized = $false
  $isMultiSession = $false
  $osAge = "Current"
  
  if ($skuName -match 'win11') { $osGeneration = "Windows 11" }
  elseif ($skuName -match 'win10') { $osGeneration = "Windows 10" }
  elseif ($skuName -match '2022') { $osGeneration = "Server 2022" }
  elseif ($skuName -match '2019') { $osGeneration = "Server 2019"; $osAge = "Legacy" }
  elseif ($skuName -match '2016') { $osGeneration = "Server 2016"; $osAge = "End of Life" }
  
  if ($skuName -match 'avd') { $isAVDOptimized = $true }
  if ($skuName -match 'evd|avd|multisession|multi-session') { $isMultiSession = $true }
  
  # Detect OS build freshness from SKU
  $osBuild = "Unknown"
  if ($skuName -match '24h2') { $osBuild = "24H2" }
  elseif ($skuName -match '23h2') { $osBuild = "23H2" }
  elseif ($skuName -match '22h2') { $osBuild = "22H2" }
  elseif ($skuName -match '21h2') { $osBuild = "21H2"; $osAge = "Legacy" }
  
  $findings_img = @()
  if (@($versions).Count -gt 1) { $findings_img += "Multiple image versions in use ($(@($versions).Count)) — configuration drift" }
  if ($osAge -eq "Legacy") { $findings_img += "Legacy OS build ($osBuild/$osGeneration) — consider upgrading" }
  if ($osAge -eq "End of Life") { $findings_img += "End of Life OS ($osGeneration) — upgrade immediately" }
  if (-not $isAVDOptimized -and $sampleVm.ImagePublisher -eq "MicrosoftWindowsDesktop") { $findings_img += "Not using AVD-optimized image — avd SKUs include pre-configured optimizations" }
  if ($hostPools_img.Count -gt 1 -and @($versions).Count -gt 1) { $findings_img += "Same image used across $($hostPools_img.Count) host pools with version drift" }
  if (@($findings_img).Count -eq 0) { $findings_img += "Consistent" }
  
  $imageAnalysis.Add([PSCustomObject]@{
    ImageSource     = "Marketplace"
    ImagePublisher  = $sampleVm.ImagePublisher
    ImageOffer      = $sampleVm.ImageOffer
    ImageSku        = $sampleVm.ImageSku
    VersionsInUse   = ($versions -join ", ")
    VersionCount    = @($versions).Count
    VMCount         = $vmCount_img
    HostPools       = ($hostPools_img -join ", ")
    HostPoolCount   = $hostPools_img.Count
    OSGeneration    = $osGeneration
    OSBuild         = $osBuild
    OSAge           = $osAge
    IsAVDOptimized  = $isAVDOptimized
    IsMultiSession  = $isMultiSession
    GalleryName     = $null
    GalleryImageDef = $null
    LatestVersion   = $null
    VersionAge      = $null
    Finding         = ($findings_img -join "; ")
  })
}

# Group gallery images by gallery/image definition
$galleryAnalysis = [System.Collections.Generic.List[object]]::new()
$galleryCache = @{}

foreach ($gvm in $galleryVms) {
  $gid = $gvm.ImageId
  if (-not $gid) { continue }
  
  # Parse gallery info from ID: .../galleries/{galleryName}/images/{imageDef}/versions/{version}
  if ($gid -match '/galleries/([^/]+)/images/([^/]+)(/versions/([^/]+))?') {
    $galleryName = $Matches[1]
    $imageDef = $Matches[2]
    $imgVersion = $Matches[4]
    $galleryKey = "$galleryName/$imageDef"
    
    if (-not $galleryCache.ContainsKey($galleryKey)) {
      $galleryCache[$galleryKey] = @{
        GalleryName = $galleryName
        ImageDef = $imageDef
        Versions = [System.Collections.Generic.List[string]]::new()
        VMs = [System.Collections.Generic.List[object]]::new()
        HostPools = [System.Collections.Generic.List[string]]::new()
        FullId = $gid
      }
    }
    if ($imgVersion -and $imgVersion -notin $galleryCache[$galleryKey].Versions) {
      $galleryCache[$galleryKey].Versions.Add($imgVersion)
    }
    $galleryCache[$galleryKey].VMs.Add($gvm)
    if ($gvm.HostPoolName -and $gvm.HostPoolName -notin $galleryCache[$galleryKey].HostPools) {
      $galleryCache[$galleryKey].HostPools.Add($gvm.HostPoolName)
    }
  }
}

# Check gallery image versions for age (requires API call)
foreach ($gEntry in $galleryCache.GetEnumerator()) {
  $gInfo = $gEntry.Value
  $latestVersionDate = $null
  $latestVersionName = $null
  $versionAgeDays = $null
  
  # Try to get the gallery image definition to check latest version
  try {
    # Parse RG from full ID
    $fullId = $gInfo.FullId
    if ($fullId -match '/resourceGroups/([^/]+)/') {
      $galleryRg = $Matches[1]
      $galleryVersions = Get-AzGalleryImageVersion -ResourceGroupName $galleryRg -GalleryName $gInfo.GalleryName -GalleryImageDefinitionName $gInfo.ImageDef -ErrorAction SilentlyContinue
      if ($galleryVersions) {
        $latest = $galleryVersions | Sort-Object { $_.PublishingProfile.PublishedDate } -Descending | Select-Object -First 1
        if ($latest) {
          $latestVersionName = $latest.Name
          $latestVersionDate = $latest.PublishingProfile.PublishedDate
          if ($latestVersionDate) {
            $versionAgeDays = [math]::Round(((Get-Date) - [datetime]$latestVersionDate).TotalDays, 0)
          }
          
          # Check replication status
          $replicationRegions = @()
          if ($latest.PublishingProfile.TargetRegions) {
            $replicationRegions = @($latest.PublishingProfile.TargetRegions | ForEach-Object { $_.Name })
          }
          
          # Check if image is replicated to all regions where VMs are deployed
          $vmRegions = @($gInfo.VMs | Where-Object { $_.Region } | Select-Object -ExpandProperty Region -Unique)
          $missingRegions = @($vmRegions | Where-Object { $_ -notin $replicationRegions -and ($replicationRegions | ForEach-Object { $_.Replace(" ", "").ToLower() }) -notcontains $_.Replace(" ", "").ToLower() })
        }
      }
    }
  } catch {
    # Gallery lookup failed — not critical
  }
  
  $findings_gallery = @()
  if ($gInfo.Versions.Count -gt 1) { $findings_gallery += "Multiple versions deployed ($($gInfo.Versions.Count)) — not all hosts on latest" }
  if ($versionAgeDays -and $versionAgeDays -gt 90) { $findings_gallery += "Latest image version is $versionAgeDays days old — may be missing security patches" }
  if ($versionAgeDays -and $versionAgeDays -gt 180) { $findings_gallery += "Image severely outdated ($versionAgeDays days) — update immediately" }
  if ($missingRegions -and @($missingRegions).Count -gt 0) { $findings_gallery += "Image not replicated to: $($missingRegions -join ', ')" }
  if (@($findings_gallery).Count -eq 0) { $findings_gallery += "Consistent and up to date" }
  
  $imageAnalysis.Add([PSCustomObject]@{
    ImageSource     = "Gallery"
    ImagePublisher  = $null
    ImageOffer      = $null
    ImageSku        = $null
    VersionsInUse   = ($gInfo.Versions -join ", ")
    VersionCount    = $gInfo.Versions.Count
    VMCount         = $gInfo.VMs.Count
    HostPools       = ($gInfo.HostPools -join ", ")
    HostPoolCount   = $gInfo.HostPools.Count
    OSGeneration    = $null
    OSBuild         = $null
    OSAge           = $null
    IsAVDOptimized  = $null
    IsMultiSession  = $null
    GalleryName     = $gInfo.GalleryName
    GalleryImageDef = $gInfo.ImageDef
    LatestVersion   = $latestVersionName
    VersionAge      = $versionAgeDays
    Finding         = ($findings_gallery -join "; ")
  })
  
  $galleryAnalysis.Add([PSCustomObject]@{
    GalleryName     = $gInfo.GalleryName
    ImageDefinition = $gInfo.ImageDef
    VersionsDeployed = ($gInfo.Versions -join ", ")
    VersionCount    = $gInfo.Versions.Count
    LatestVersion   = $latestVersionName
    LatestPublished = $latestVersionDate
    AgeDays         = $versionAgeDays
    VMCount         = $gInfo.VMs.Count
    HostPools       = ($gInfo.HostPools -join ", ")
    ReplicatedTo    = if ($replicationRegions) { $replicationRegions -join ", " } else { "Unknown" }
    MissingRegions  = if ($missingRegions -and @($missingRegions).Count -gt 0) { $missingRegions -join ", " } else { "" }
    Finding         = ($findings_gallery -join "; ")
  })
}

# --- Per-host-pool image consistency check ---
$hpImageConsistency = [System.Collections.Generic.List[object]]::new()
$hpImageGroups = $vms | Group-Object HostPoolName
foreach ($hpImg in $hpImageGroups) {
  $hpVms_img = @($hpImg.Group)
  $sources = @($hpVms_img | Where-Object { $_.PSObject.Properties['ImageSource'] } | Select-Object -ExpandProperty ImageSource -Unique)
  $skus = @($hpVms_img | Where-Object { $_.ImageSku } | Select-Object -ExpandProperty ImageSku -Unique)
  $versions = @($hpVms_img | Where-Object { $_.ImageVersion } | Select-Object -ExpandProperty ImageVersion -Unique)
  $publishers = @($hpVms_img | Where-Object { $_.ImagePublisher } | Select-Object -ExpandProperty ImagePublisher -Unique)
  
  $mixedSources = ($sources.Count -gt 1)
  $mixedSkus = ($skus.Count -gt 1)
  $mixedVersions = ($versions.Count -gt 1)
  
  $consistency = "Consistent"
  $hpFindings = @()
  if ($mixedSources) { $consistency = "Mixed Sources"; $hpFindings += "VMs from different image sources ($($sources -join ', ')) — standardize on a golden image" }
  elseif ($mixedSkus) { $consistency = "Mixed SKUs"; $hpFindings += "Multiple image SKUs ($($skus -join ', ')) in same pool" }
  elseif ($mixedVersions) { $consistency = "Version Drift"; $hpFindings += "Multiple image versions ($($versions -join ', ')) — reimage outdated hosts" }
  
  $hpImageConsistency.Add([PSCustomObject]@{
    HostPoolName    = $hpImg.Name
    VMCount         = $hpVms_img.Count
    ImageSources    = ($sources -join ", ")
    SourceCount     = $sources.Count
    ImageSkus       = ($skus -join ", ")
    SkuCount        = $skus.Count
    ImageVersions   = ($versions -join ", ")
    VersionCount    = $versions.Count
    Consistency     = $consistency
    Finding         = if (@($hpFindings).Count -gt 0) { $hpFindings -join "; " } else { "All VMs using same image and version" }
  })
}

# --- Golden Image Maturity Assessment ---
$totalVmCount = ($vms | Measure-Object).Count
$goldenImageScore = 0
$goldenImageFindings = @()

# Score: Using gallery images (best practice) = +30
$galleryPct = if ($totalVmCount -gt 0) { [math]::Round(($galleryVms.Count / $totalVmCount) * 100, 0) } else { 0 }
if ($galleryPct -ge 80) { $goldenImageScore += 30; $goldenImageFindings += "✅ $galleryPct% of VMs use Azure Compute Gallery images (best practice)" }
elseif ($galleryPct -gt 0) { $goldenImageScore += 15; $goldenImageFindings += "⚠️ Only $galleryPct% on gallery images — migrate remaining $($totalVmCount - $galleryVms.Count) VMs" }
else { $goldenImageFindings += "❌ No VMs using Azure Compute Gallery — consider implementing a golden image pipeline" }

# Score: Single image version per pool = +25
$consistentPools = @($hpImageConsistency | Where-Object { $_.Consistency -eq "Consistent" })
$consistentPct = if ($hpImageConsistency.Count -gt 0) { [math]::Round(($consistentPools.Count / $hpImageConsistency.Count) * 100, 0) } else { 0 }
if ($consistentPct -eq 100) { $goldenImageScore += 25; $goldenImageFindings += "✅ All host pools have consistent images" }
elseif ($consistentPct -ge 50) { $goldenImageScore += 12; $goldenImageFindings += "⚠️ $consistentPct% of host pools have consistent images" }
else { $goldenImageFindings += "❌ Most host pools have image drift — implement standardized image pipeline" }

# Score: No marketplace images in production = +20
$marketplacePct = if ($totalVmCount -gt 0) { [math]::Round(($marketplaceVms.Count / $totalVmCount) * 100, 0) } else { 0 }
if ($marketplacePct -eq 0) { $goldenImageScore += 20; $goldenImageFindings += "✅ No raw marketplace images — all using custom/gallery images" }
elseif ($marketplacePct -lt 20) { $goldenImageScore += 10; $goldenImageFindings += "⚠️ $marketplacePct% still on marketplace images — migrate to golden images" }
else { $goldenImageFindings += "❌ $marketplacePct% of VMs use marketplace images — these lack customizations and app pre-installs" }

# Score: AVD-optimized images = +15
$avdOptimizedCount = ($imageAnalysis | Where-Object { $_.IsAVDOptimized -eq $true } | ForEach-Object { $_.VMCount } | Measure-Object -Sum).Sum
if (-not $avdOptimizedCount) { $avdOptimizedCount = 0 }
$avdOptPct = if ($totalVmCount -gt 0) { [math]::Round(($avdOptimizedCount / $totalVmCount) * 100, 0) } else { 0 }
if ($avdOptPct -ge 80 -or $galleryPct -ge 80) { $goldenImageScore += 15; $goldenImageFindings += "✅ Using AVD-optimized or custom gallery images" }
elseif ($avdOptPct -gt 0) { $goldenImageScore += 7; $goldenImageFindings += "⚠️ Only $avdOptPct% on AVD-optimized images" }
else { $goldenImageFindings += "❌ No AVD-optimized images in use — avd SKUs include Windows optimizations for virtual desktops" }

# Score: Recent image versions = +10
$staleGalleryImages = @($galleryAnalysis | Where-Object { $_.AgeDays -and $_.AgeDays -gt 90 })
if ($galleryAnalysis.Count -gt 0 -and $staleGalleryImages.Count -eq 0) { $goldenImageScore += 10; $goldenImageFindings += "✅ All gallery images updated within 90 days" }
elseif ($staleGalleryImages.Count -gt 0) { $goldenImageFindings += "❌ $($staleGalleryImages.Count) gallery image(s) older than 90 days — missing security patches" }

$goldenImageGrade = if ($goldenImageScore -ge 80) { "A" } 
                    elseif ($goldenImageScore -ge 60) { "B" }
                    elseif ($goldenImageScore -ge 40) { "C" }
                    elseif ($goldenImageScore -ge 20) { "D" }
                    else { "F" }

$multiVersionImages = ($imageAnalysis | Where-Object { $_.VersionCount -gt 1 } | Measure-Object).Count
$customImages = $vms | Where-Object { -not $_.ImagePublisher -and ($_.PSObject.Properties['ImageOffer'] -and $_.ImageOffer) }
Write-Host "  Images: $($imageAnalysis.Count) image groups, $multiVersionImages with version drift" -ForegroundColor $(if ($multiVersionImages -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Sources: $($marketplaceVms.Count) Marketplace, $($galleryVms.Count) Gallery, $($managedImageVms.Count) Managed Image" -ForegroundColor Gray
Write-Host "  Golden Image Maturity: $goldenImageScore/100 (Grade: $goldenImageGrade)" -ForegroundColor $(if ($goldenImageScore -ge 60) { "Green" } elseif ($goldenImageScore -ge 40) { "Yellow" } else { "Red" })

# =========================================================
# Enhanced Analysis: Actual Cost Collection (v4.0.0)
# =========================================================
# Queries Azure Cost Management API for real billed costs per VM
$actualCostData = [System.Collections.Generic.List[object]]::new()
$vmActualMonthlyCost = @{}  # Lookup: VM resource ID → estimated monthly cost
$infraCostData = [System.Collections.Generic.List[object]]::new()  # Networking, storage, AVD service costs per RG

# Build AVD resource group set early (needed for infra cost scoping)
if (-not (Test-Path variable:avdResourceGroups)) {
  $avdResourceGroups = @{}
  foreach ($v in $vms) {
    if ($v.ResourceGroup) { $avdResourceGroups["$($v.SubscriptionId)|$($v.ResourceGroup)".ToLower()] = $true }
  }
  foreach ($hp in $hostPools) {
    $hpSubId = if ($hp.SubscriptionId) { $hp.SubscriptionId } elseif ($hp.ArmId) { (Get-SubFromArmId $hp.ArmId) } else { $null }
    if ($hpSubId -and $hp.ResourceGroup) { $avdResourceGroups["$hpSubId|$($hp.ResourceGroup)".ToLower()] = $true }
  }
}

if (-not $SkipActualCosts) {
  Write-Host "Querying Azure Cost Management for actual billed costs..." -ForegroundColor Cyan
  
  # --- Permissions pre-check ---
  # Test access with a lightweight dimensions query before running the full cost query
  $costAccessDenied = [System.Collections.Generic.List[string]]::new()
  $costAccessGranted = [System.Collections.Generic.List[string]]::new()
  
  foreach ($subId in $SubscriptionIds) {
    try {
      $testPath = "/subscriptions/$subId/providers/Microsoft.CostManagement/query?api-version=2023-11-01"
      $testBody = @{
        type = "AmortizedCost"
        dataSet = @{
          granularity = "None"
          aggregation = @{ totalCost = @{ name = "Cost"; function = "Sum" } }
        }
        timeframe = "MonthToDate"
      } | ConvertTo-Json -Depth 5
      
      $testResponse = Invoke-AzRestMethod -Path $testPath -Method POST -Payload $testBody
      
      if ($testResponse.StatusCode -eq 200) {
        $costAccessGranted.Add($subId)
      }
      elseif ($testResponse.StatusCode -eq 401 -or $testResponse.StatusCode -eq 403) {
        $costAccessDenied.Add($subId)
        $errDetail = "AccessDenied"
        try { $errDetail = ($testResponse.Content | ConvertFrom-Json).error.code } catch {}
        Write-Host "  ⚠ Subscription $subId — access denied ($errDetail)" -ForegroundColor Yellow
        Write-Host "    Required: Cost Management Reader role or higher" -ForegroundColor Yellow
      }
      elseif ($testResponse.StatusCode -eq 404 -or $testResponse.StatusCode -eq 409) {
        # 404 = Cost Management not registered; 409 = subscription not enrolled
        $costAccessDenied.Add($subId)
        Write-Host "  ⚠ Subscription $subId — Cost Management not available (HTTP $($testResponse.StatusCode))" -ForegroundColor Yellow
        Write-Host "    Ensure Microsoft.CostManagement resource provider is registered" -ForegroundColor Yellow
      }
      else {
        $costAccessDenied.Add($subId)
        Write-Host "  ⚠ Subscription $subId — unexpected response (HTTP $($testResponse.StatusCode))" -ForegroundColor Yellow
      }
    }
    catch {
      $costAccessDenied.Add($subId)
      Write-Host "  ⚠ Subscription $subId — connection error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
  }
  
  if ($costAccessGranted.Count -eq 0) {
    Write-Host ""
    Write-Host "  ✗ No subscriptions have Cost Management access." -ForegroundColor Red
    Write-Host "    To grant access, run:" -ForegroundColor Yellow
    Write-Host "      New-AzRoleAssignment -SignInName <your-email> -RoleDefinitionName 'Cost Management Reader' -Scope /subscriptions/<sub-id>" -ForegroundColor Gray
    Write-Host "    Falling back to PAYG estimates for all cost analysis." -ForegroundColor Yellow
    Write-Host ""
  } elseif ($costAccessDenied.Count -gt 0) {
    Write-Host "  ✓ $($costAccessGranted.Count) subscription(s) accessible, $($costAccessDenied.Count) denied — partial actual costs" -ForegroundColor Yellow
  } else {
    Write-Host "  ✓ Cost Management access confirmed for all $($costAccessGranted.Count) subscription(s)" -ForegroundColor Green
  }
  
  # --- Query actual costs for accessible subscriptions ---
  $costStartDate = (Get-Date).AddDays(-30).ToString("yyyy-MM-dd")
  $costEndDate = (Get-Date).ToString("yyyy-MM-dd")
  
  foreach ($subId in $costAccessGranted) {
    Write-Host "  Querying costs for subscription $subId..." -ForegroundColor Gray
    
    # Build Cost Management query — amortized costs grouped by resource ID
    # AmortizedCost spreads RI/Savings Plan purchases across covered resources,
    # giving the true effective cost per VM (ActualCost shows $0 for RI-covered VMs)
    # Scoped to AVD resource groups to avoid pulling costs for non-AVD VMs
    $subAvdRgNames = @($avdResourceGroups.Keys | Where-Object { $_.StartsWith("$subId|".ToLower()) } | ForEach-Object { ($_ -split '\|', 2)[1] })
    
    $costFilter = $null
    if ($subAvdRgNames.Count -gt 0) {
      # Scope to AVD resource groups + resource types
      $costFilter = @{
        and = @(
          @{
            dimensions = @{
              name = "ResourceType"
              operator = "In"
              values = @(
                "Microsoft.Compute/virtualMachines"
                "Microsoft.Compute/virtualMachineScaleSets"
                "Microsoft.Compute/disks"
              )
            }
          }
          @{
            dimensions = @{
              name = "ResourceGroupName"
              operator = "In"
              values = $subAvdRgNames
            }
          }
        )
      }
      Write-Host "    Scoping cost query to $($subAvdRgNames.Count) AVD resource group(s): $($subAvdRgNames -join ', ')" -ForegroundColor Gray
    } else {
      # Fallback: no RG info, query all resource types (may include non-AVD VMs)
      $costFilter = @{
        dimensions = @{
          name = "ResourceType"
          operator = "In"
          values = @(
            "Microsoft.Compute/virtualMachines"
            "Microsoft.Compute/virtualMachineScaleSets"
            "Microsoft.Compute/disks"
          )
        }
      }
      Write-Host "    ⚠ No AVD resource groups identified — querying all VMs in subscription" -ForegroundColor Yellow
    }
    
    $costQueryBody = @{
      type = "AmortizedCost"
      dataSet = @{
        granularity = "Daily"
        aggregation = @{
          totalCost = @{
            name = "Cost"
            function = "Sum"
          }
        }
        grouping = @(
          @{ type = "Dimension"; name = "ResourceId" }
          @{ type = "Dimension"; name = "ResourceType" }
          @{ type = "Dimension"; name = "MeterCategory" }
          @{ type = "Dimension"; name = "PricingModel" }
        )
        filter = $costFilter
      }
      timeframe = "Custom"
      timePeriod = @{
        from = $costStartDate
        to = $costEndDate
      }
    } | ConvertTo-Json -Depth 10
    
    try {
      $costApiPath = "/subscriptions/$subId/providers/Microsoft.CostManagement/query?api-version=2023-11-01"
      $costResponse = Invoke-AzRestMethod -Path $costApiPath -Method POST -Payload $costQueryBody
      
      if ($costResponse.StatusCode -eq 200) {
        $costResult = $costResponse.Content | ConvertFrom-Json
        $columns = $costResult.properties.columns
        $rows = $costResult.properties.rows
        
        # Find column indices
        $costIdx = 0
        $dateIdx = -1
        $resIdIdx = -1
        $resTypeIdx = -1
        $meterIdx = -1
        $pricingModelIdx = -1
        for ($i = 0; $i -lt $columns.Count; $i++) {
          switch ($columns[$i].name) {
            "Cost"         { $costIdx = $i }
            "UsageDate"    { $dateIdx = $i }
            "ResourceId"   { $resIdIdx = $i }
            "ResourceType" { $resTypeIdx = $i }
            "MeterCategory" { $meterIdx = $i }
            "PricingModel" { $pricingModelIdx = $i }
          }
        }
        
        if ($rows -and $resIdIdx -ge 0) {
          # Aggregate per resource
          $resourceCosts = @{}
          $reservationCosts = @{}  # Track RI costs separately for matching
          $totalRowCost = 0.0
          foreach ($row in $rows) {
            $resId = "$($row[$resIdIdx])".ToLower()
            $cost = [double]$row[$costIdx]
            $meterCat = if ($meterIdx -ge 0) { "$($row[$meterIdx])" } else { "Unknown" }
            $pricingModel = if ($pricingModelIdx -ge 0) { "$($row[$pricingModelIdx])" } else { "Unknown" }
            $resType = if ($resTypeIdx -ge 0) { "$($row[$resTypeIdx])" } else { "" }
            $totalRowCost += $cost
            
            # Reservation costs may appear under Microsoft.Capacity/reservationOrders
            # or under the actual VM but with PricingModel = "Reservation"
            if ($resId -match 'microsoft\.capacity/reservationorders' -or $resType -match 'Capacity') {
              # Group reservation costs — we'll distribute these later
              if (-not $reservationCosts.ContainsKey($resId)) {
                $reservationCosts[$resId] = @{ TotalCost = 0.0; MeterCategory = $meterCat }
              }
              $reservationCosts[$resId].TotalCost += $cost
              continue
            }
            
            if (-not $resourceCosts.ContainsKey($resId)) {
              $resourceCosts[$resId] = @{ TotalCost = 0.0; ComputeCost = 0.0; StorageCost = 0.0; DayCount = 0; MeterCategories = @{}; PricingModels = @{} }
            }
            $resourceCosts[$resId].TotalCost += $cost
            if ($meterCat -match "Virtual Machines|Compute") {
              $resourceCosts[$resId].ComputeCost += $cost
            } elseif ($meterCat -match "Storage|Disk") {
              $resourceCosts[$resId].StorageCost += $cost
            }
            if (-not $resourceCosts[$resId].MeterCategories.ContainsKey($meterCat)) {
              $resourceCosts[$resId].MeterCategories[$meterCat] = 0.0
            }
            $resourceCosts[$resId].MeterCategories[$meterCat] += $cost
            # Track pricing model — "Reservation", "SavingsPlan", "OnDemand", "Spot"
            if (-not $resourceCosts[$resId].PricingModels.ContainsKey($pricingModel)) {
              $resourceCosts[$resId].PricingModels[$pricingModel] = 0.0
            }
            $resourceCosts[$resId].PricingModels[$pricingModel] += $cost
          }
          
          # Diagnostic: show what Cost Management returned
          $vmResources = @($resourceCosts.Keys | Where-Object { $_ -match 'virtualmachine' })
          $diskResources = @($resourceCosts.Keys | Where-Object { $_ -match 'disk' })
          Write-Host "    Cost API returned: $($rows.Count) rows, $($resourceCosts.Count) resources ($($vmResources.Count) VMs, $($diskResources.Count) disks), `$$([math]::Round($totalRowCost, 2)) total" -ForegroundColor Gray
          if ($reservationCosts.Count -gt 0) {
            $riTotal = ($reservationCosts.Values | ForEach-Object { $_.TotalCost } | Measure-Object -Sum).Sum
            Write-Host "    Reservation line items: $($reservationCosts.Count) (`$$([math]::Round($riTotal, 2)))" -ForegroundColor Gray
          }
          # Show per-resource breakdown for troubleshooting
          foreach ($rcEntry in ($resourceCosts.GetEnumerator() | Sort-Object { $_.Value.TotalCost } -Descending | Select-Object -First 6)) {
            $rcName = ($rcEntry.Key -split '/')[-1]
            $rcType = if ($rcEntry.Key -match 'virtualmachine') { 'VM' } elseif ($rcEntry.Key -match 'disk') { 'Disk' } else { 'Other' }
            $rcModels = ($rcEntry.Value.PricingModels.Keys -join ',')
            Write-Host "      $rcType $rcName : `$$([math]::Round($rcEntry.Value.TotalCost, 2)) (Compute:`$$([math]::Round($rcEntry.Value.ComputeCost, 2)) Stor:`$$([math]::Round($rcEntry.Value.StorageCost, 2))) [$rcModels]" -ForegroundColor Gray
          }
          
          # If we have unmatched reservation costs, distribute them across AVD VMs in this sub
          # This handles the case where RI amortized cost appears under reservationOrders instead of the VM
          if ($reservationCosts.Count -gt 0) {
            $riTotal = ($reservationCosts.Values | ForEach-Object { $_.TotalCost } | Measure-Object -Sum).Sum
            if ($riTotal -gt 0) {
              # Only distribute to AVD VMs that don't already have compute cost data
              $avdVmsNeedingRi = @($vms | Where-Object { 
                ($_.SubscriptionId -eq $subId -or (TryGet-ArmId $_) -match $subId) -and $(
                  $vid = "$($_.VMId)".ToLower().Trim()
                  -not $resourceCosts.ContainsKey($vid) -or $resourceCosts[$vid].ComputeCost -eq 0
                )
              })
              if ($avdVmsNeedingRi.Count -gt 0) {
                $riPerVm = $riTotal / $avdVmsNeedingRi.Count
                Write-Host "    Distributing `$$([math]::Round($riTotal, 2)) reservation cost across $($avdVmsNeedingRi.Count) AVD VM(s) without cost data (`$$([math]::Round($riPerVm, 2))/VM)" -ForegroundColor Yellow
                foreach ($svm in $avdVmsNeedingRi) {
                  $svmId = "$($svm.VMId)".ToLower().Trim()
                  if (-not $svmId) { continue }
                  if (-not $resourceCosts.ContainsKey($svmId)) {
                    $resourceCosts[$svmId] = @{ TotalCost = 0.0; ComputeCost = 0.0; StorageCost = 0.0; DayCount = 0; MeterCategories = @{}; PricingModels = @{} }
                  }
                  $resourceCosts[$svmId].TotalCost += $riPerVm
                  $resourceCosts[$svmId].ComputeCost += $riPerVm
                  if (-not $resourceCosts[$svmId].PricingModels.ContainsKey("Reservation")) {
                    $resourceCosts[$svmId].PricingModels["Reservation"] = 0.0
                  }
                  $resourceCosts[$svmId].PricingModels["Reservation"] += $riPerVm
                }
              } else {
                Write-Host "    RI costs (`$$([math]::Round($riTotal, 2))) detected but all AVD VMs already have cost data — amortization appears correct" -ForegroundColor Gray
              }
            }
          }
          
          # Check for VMs with $0 compute cost — RI amortization may not resolve to individual VMs
          # In this case, run an RG-level query to get the true compute cost including RI benefit
          $zeroComputeVms = @($vms | Where-Object {
            $vid = "$($_.VMId)".ToLower().Trim()
            $vid -and $resourceCosts.ContainsKey($vid) -and $resourceCosts[$vid].ComputeCost -eq 0
          })
          if ($zeroComputeVms.Count -gt 0 -and $reservationCosts.Count -eq 0) {
            Write-Host "    ⚠ $($zeroComputeVms.Count) VM(s) show `$0 compute cost — querying RG-level costs to capture RI amortization" -ForegroundColor Yellow
            
            # Query at RG level grouped by MeterCategory — this DOES include RI amortized costs
            foreach ($rgName in $subAvdRgNames) {
              try {
                $rgCostBody = @{
                  type = "AmortizedCost"
                  dataSet = @{
                    granularity = "None"
                    aggregation = @{ totalCost = @{ name = "Cost"; function = "Sum" } }
                    grouping = @(
                      @{ type = "Dimension"; name = "MeterCategory" }
                      @{ type = "Dimension"; name = "ResourceType" }
                    )
                    filter = @{
                      and = @(
                        @{ dimensions = @{ name = "ResourceGroupName"; operator = "In"; values = @($rgName) } }
                        @{ dimensions = @{ name = "ResourceType"; operator = "In"; values = @("Microsoft.Compute/virtualMachines","Microsoft.Compute/virtualMachineScaleSets") } }
                      )
                    }
                  }
                  timeframe = "Custom"
                  timePeriod = @{ from = $costStartDate; to = $costEndDate }
                } | ConvertTo-Json -Depth 10
                
                $rgCostResp = Invoke-AzRestMethod -Path $costApiPath -Method POST -Payload $rgCostBody
                if ($rgCostResp.StatusCode -eq 200) {
                  $rgResult = $rgCostResp.Content | ConvertFrom-Json
                  $rgRows = $rgResult.properties.rows
                  $rgCols = $rgResult.properties.columns
                  if ($rgRows) {
                    $rgCostIdx = 0; $rgMeterIdx = -1
                    for ($ri = 0; $ri -lt $rgCols.Count; $ri++) {
                      if ($rgCols[$ri].name -eq "Cost") { $rgCostIdx = $ri }
                      if ($rgCols[$ri].name -eq "MeterCategory") { $rgMeterIdx = $ri }
                    }
                    $rgComputeTotal = 0.0
                    foreach ($rgRow in $rgRows) {
                      $rgMeter = if ($rgMeterIdx -ge 0) { "$($rgRow[$rgMeterIdx])" } else { "" }
                      if ($rgMeter -match "Virtual Machines|Compute") {
                        $rgComputeTotal += [double]$rgRow[$rgCostIdx]
                      }
                    }
                    
                    if ($rgComputeTotal -gt 0) {
                      # Distribute RG-level compute cost across zero-compute AVD VMs in this RG
                      $rgZeroVms = @($zeroComputeVms | Where-Object { $_.ResourceGroup -eq $rgName })
                      if ($rgZeroVms.Count -gt 0) {
                        $riPerVm = $rgComputeTotal / $rgZeroVms.Count
                        $daysInRange = ((Get-Date) - [datetime]$costStartDate).Days
                        $riMonthlyPerVm = [math]::Round(($riPerVm / [math]::Max(1, $daysInRange)) * 30, 2)
                        Write-Host "      RG '$rgName': `$$([math]::Round($rgComputeTotal, 2)) compute (RI amortized) -> `$$riMonthlyPerVm/mo per VM across $($rgZeroVms.Count) VM(s)" -ForegroundColor Green
                        foreach ($zvm in $rgZeroVms) {
                          $zvmId = "$($zvm.VMId)".ToLower().Trim()
                          $resourceCosts[$zvmId].TotalCost += $riPerVm
                          $resourceCosts[$zvmId].ComputeCost += $riPerVm
                          if (-not $resourceCosts[$zvmId].PricingModels.ContainsKey("Reservation (RG-level)")) {
                            $resourceCosts[$zvmId].PricingModels["Reservation (RG-level)"] = 0.0
                          }
                          $resourceCosts[$zvmId].PricingModels["Reservation (RG-level)"] += $riPerVm
                        }
                      }
                    } else {
                      Write-Host "      RG '$rgName': RG-level query also shows `$0 compute - falling back to RI pricing table" -ForegroundColor Yellow
                      # Last resort: use RI pricing table estimate
                      foreach ($zvm in @($zeroComputeVms | Where-Object { $_.ResourceGroup -eq $rgName })) {
                        $zvmId = "$($zvm.VMId)".ToLower().Trim()
                        $riRate = $null
                        if ($riPricingTable -and $riPricingTable.ContainsKey($zvm.VMSize)) {
                          $riRate = $riPricingTable[$zvm.VMSize].RI1Y
                        }
                        if (-not $riRate) {
                          $riRate = Get-EstimatedVmCostPerHour -VmSize $zvm.VMSize -Region $zvm.Region
                          if ($riRate) { $riRate = $riRate * 0.6 }
                        }
                        if ($riRate) {
                          $riMonthly = [math]::Round($riRate * 730, 2)
                          $daysInRange = ((Get-Date) - [datetime]$costStartDate).Days
                          $riPeriodCost = $riRate * 730 / 30 * $daysInRange
                          $resourceCosts[$zvmId].TotalCost += $riPeriodCost
                          $resourceCosts[$zvmId].ComputeCost += $riPeriodCost
                          if (-not $resourceCosts[$zvmId].PricingModels.ContainsKey("RI Estimate")) {
                            $resourceCosts[$zvmId].PricingModels["RI Estimate"] = 0.0
                          }
                          $resourceCosts[$zvmId].PricingModels["RI Estimate"] += $riRate * 730
                          Write-Host "      $($zvm.VMName) ($($zvm.VMSize)): estimated RI cost ~`$$riMonthly/mo" -ForegroundColor Gray
                        }
                      }
                    }
                  }
                }
              } catch {
                Write-Host "      ⚠ RG-level cost query failed for $rgName - $($_.Exception.Message)" -ForegroundColor Yellow
              }
            }
          }

          
          # Match to our VMs and project monthly
          # For VMSS: Cost Management groups costs at the scale set level, not per instance.
          # We split the scale set total evenly across instances.
          $vmssMatched = @{}  # Track VMSS costs to split across instances
          
          foreach ($vm in $vms) {
            $vmResId = "$($vm.VMId)".ToLower().Trim()
            if (-not $vmResId) { continue }
            
            $matched = $null
            
            # 1. Exact match on VM resource ID (lowercased)
            if ($resourceCosts.ContainsKey($vmResId)) {
              $matched = $resourceCosts[$vmResId]
            }
            
            # 2. Try matching by resource name from ID (handles casing/path differences)
            if (-not $matched) {
              $vmNameFromId = ($vmResId -split '/')[-1]
              if ($vmNameFromId) {
                $nameMatch = $resourceCosts.GetEnumerator() | Where-Object {
                  ($_.Key -split '/')[-1] -eq $vmNameFromId -and $_.Key -match 'virtualMachines'
                } | Select-Object -First 1
                if ($nameMatch) { $matched = $nameMatch.Value }
              }
            }
            
            # 2. For VMSS VMs, try matching the scale set resource ID
            if (-not $matched -and $vm.PSObject.Properties.Name -contains 'IsVMSS' -and $vm.IsVMSS -and $vm.VMSSName) {
              $vmssPattern = "/virtualMachineScaleSets/$($vm.VMSSName)".ToLower()
              $vmssEntry = $resourceCosts.GetEnumerator() | Where-Object { $_.Key -match [regex]::Escape($vmssPattern) } | Select-Object -First 1
              if ($vmssEntry) {
                # Count how many of our VMs belong to this VMSS
                $vmssKey = $vmssEntry.Key
                if (-not $vmssMatched.ContainsKey($vmssKey)) {
                  $instanceCount = @($vms | Where-Object { $_.PSObject.Properties.Name -contains 'VMSSName' -and $_.VMSSName -eq $vm.VMSSName }).Count
                  $vmssMatched[$vmssKey] = [math]::Max(1, $instanceCount)
                }
                # Split cost evenly across instances
                $matched = @{
                  TotalCost = $vmssEntry.Value.TotalCost / $vmssMatched[$vmssKey]
                  ComputeCost = $vmssEntry.Value.ComputeCost / $vmssMatched[$vmssKey]
                  StorageCost = $vmssEntry.Value.StorageCost / $vmssMatched[$vmssKey]
                }
              }
            }
            
            # 3. Fallback: match by VM name in resource ID
            if (-not $matched) {
              $vmNameLower = "$($vm.VMName)".ToLower()
              if ($vmNameLower) {
                $nameMatch = $resourceCosts.GetEnumerator() | Where-Object { $_.Key -match [regex]::Escape($vmNameLower) } | Select-Object -First 1
                if ($nameMatch) { $matched = $nameMatch.Value }
              }
            }
            
            if ($matched -and $matched.TotalCost -gt 0) {
              # Calculate days with data and project to 30-day month
              $daysInRange = ((Get-Date) - [datetime]$costStartDate).Days
              $dailyAvg = $matched.TotalCost / [math]::Max(1, $daysInRange)
              $monthlyEstimate = [math]::Round($dailyAvg * 30, 2)
              $computeMonthly = [math]::Round(($matched.ComputeCost / [math]::Max(1, $daysInRange)) * 30, 2)
              $storageMonthly = [math]::Round(($matched.StorageCost / [math]::Max(1, $daysInRange)) * 30, 2)
              
              # Determine pricing model (RI, Savings Plan, OnDemand, or mixed)
              $vmPricingModel = "OnDemand"
              if ($matched.PricingModels -and $matched.PricingModels.Count -gt 0) {
                $topModel = ($matched.PricingModels.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
                $vmPricingModel = if ($topModel -match 'Reservation') { "Reserved Instance" }
                                  elseif ($topModel -match 'SavingsPlan') { "Savings Plan" }
                                  elseif ($topModel -match 'Spot') { "Spot" }
                                  else { "OnDemand" }
                # Check if mixed (e.g., partial RI coverage + OnDemand overflow)
                if ($matched.PricingModels.Count -gt 1) {
                  $riPct = 0
                  $totalPricing = ($matched.PricingModels.Values | Measure-Object -Sum).Sum
                  if ($totalPricing -gt 0) {
                    $riAmount = ($matched.PricingModels.GetEnumerator() | Where-Object { $_.Key -match 'Reservation|SavingsPlan' } | ForEach-Object { $_.Value } | Measure-Object -Sum).Sum
                    $riPct = [math]::Round($riAmount / $totalPricing * 100, 0)
                  }
                  if ($riPct -gt 0 -and $riPct -lt 100) { $vmPricingModel = "$vmPricingModel ($riPct% benefit)" }
                }
              }
              
              $vmActualMonthlyCost[$vm.VMName] = $monthlyEstimate
              
              $actualCostData.Add([PSCustomObject]@{
                SubscriptionId  = $subId
                VMName          = $vm.VMName
                HostPoolName    = $vm.HostPoolName
                VMSize          = $vm.VMSize
                ResourceId      = $vm.VMId
                Period30DayCost = [math]::Round($matched.TotalCost, 2)
                DailyAvgCost    = [math]::Round($dailyAvg, 2)
                MonthlyEstimate = $monthlyEstimate
                ComputeMonthly  = $computeMonthly
                StorageMonthly  = $storageMonthly
                PricingModel    = $vmPricingModel
                CostSource      = "AzureCostManagement"
              })
            }
          }
        }
        
        # Handle pagination — Cost Management may return nextLink
        $nextLink = $costResult.properties.nextLink
        while ($nextLink) {
          try {
            $nextResponse = Invoke-AzRestMethod -Uri $nextLink -Method GET
            if ($nextResponse.StatusCode -eq 200) {
              $nextResult = $nextResponse.Content | ConvertFrom-Json
              # Process additional rows same as above
              $nextLink = $nextResult.properties.nextLink
            } else { $nextLink = $null }
          } catch { $nextLink = $null }
        }
        
        Write-Host "    Retrieved costs for $($actualCostData.Count) VMs in this subscription" -ForegroundColor Gray
        
        # Diagnostic: show first few unmatched VMs for troubleshooting
        $unmatchedInSub = @($vms | Where-Object { 
          ($_.SubscriptionId -eq $subId -or (TryGet-ArmId $_) -match $subId) -and 
          -not $vmActualMonthlyCost.ContainsKey($_.VMName)
        })
        if ($unmatchedInSub.Count -gt 0 -and $resourceCosts.Count -gt 0) {
          Write-Host "    ⚠ $($unmatchedInSub.Count) VM(s) not matched to cost data:" -ForegroundColor Yellow
          foreach ($uvm in ($unmatchedInSub | Select-Object -First 3)) {
            Write-Host "      VM: $($uvm.VMName) | ID: $($uvm.VMId)" -ForegroundColor Gray
          }
          Write-Host "    Sample cost resource IDs returned by API:" -ForegroundColor Gray
          foreach ($sampleKey in ($resourceCosts.Keys | Select-Object -First 3)) {
            Write-Host "      $sampleKey (`$$([math]::Round($resourceCosts[$sampleKey].TotalCost, 2)))" -ForegroundColor Gray
          }
        }
      }
      else {
        $errBody = "HTTP $($costResponse.StatusCode)"
        try { $errBody = ($costResponse.Content | ConvertFrom-Json).error.message } catch {}
        Write-Host "    ⚠ Cost query failed: $errBody — falling back to PAYG estimates" -ForegroundColor Yellow
      }
    }
    catch {
      Write-Host "    ⚠ Cost query error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # --- Query 2: Infrastructure costs (networking, storage accounts, AVD service, Log Analytics) ---
    # Scoped to AVD resource groups only to avoid pulling unrelated costs
    $subAvdRgs = @($avdResourceGroups.Keys | Where-Object { $_.StartsWith("$subId|".ToLower()) } | ForEach-Object { ($_ -split '\|', 2)[1] })
    if ($subAvdRgs.Count -gt 0) {
      Write-Host "    Querying infrastructure costs for $($subAvdRgs.Count) AVD resource group(s)..." -ForegroundColor Gray
      foreach ($rgName in $subAvdRgs) {
        try {
          $infraQueryBody = @{
            type = "AmortizedCost"
            dataSet = @{
              granularity = "None"
              aggregation = @{
                totalCost = @{ name = "Cost"; function = "Sum" }
              }
              grouping = @(
                @{ type = "Dimension"; name = "ResourceType" }
                @{ type = "Dimension"; name = "MeterCategory" }
              )
              filter = @{
                and = @(
                  @{
                    dimensions = @{
                      name = "ResourceGroupName"
                      operator = "In"
                      values = @($rgName)
                    }
                  }
                  @{
                    dimensions = @{
                      name = "ResourceType"
                      operator = "In"
                      values = @(
                        "Microsoft.Network/virtualNetworks"
                        "Microsoft.Network/networkSecurityGroups"
                        "Microsoft.Network/publicIPAddresses"
                        "Microsoft.Network/loadBalancers"
                        "Microsoft.Network/natGateways"
                        "Microsoft.Network/privateEndpoints"
                        "Microsoft.Network/networkInterfaces"
                        "Microsoft.Storage/storageAccounts"
                        "Microsoft.DesktopVirtualization/hostPools"
                        "Microsoft.DesktopVirtualization/workspaces"
                        "Microsoft.DesktopVirtualization/scalingPlans"
                        "Microsoft.OperationalInsights/workspaces"
                        "Microsoft.KeyVault/vaults"
                      )
                    }
                  }
                )
              }
            }
            timeframe = "Custom"
            timePeriod = @{ from = $costStartDate; to = $costEndDate }
          } | ConvertTo-Json -Depth 10
          
          $infraResponse = Invoke-AzRestMethod -Path $costApiPath -Method POST -Payload $infraQueryBody
          if ($infraResponse.StatusCode -eq 200) {
            $infraResult = $infraResponse.Content | ConvertFrom-Json
            $infraCols = $infraResult.properties.columns
            $infraRows = $infraResult.properties.rows
            if ($infraRows) {
              $iCostIdx = 0; $iTypeIdx = -1; $iMeterIdx = -1
              for ($i = 0; $i -lt $infraCols.Count; $i++) {
                switch ($infraCols[$i].name) {
                  "Cost"         { $iCostIdx = $i }
                  "ResourceType" { $iTypeIdx = $i }
                  "MeterCategory" { $iMeterIdx = $i }
                }
              }
              foreach ($iRow in $infraRows) {
                $iCost = [double]$iRow[$iCostIdx]
                $iType = if ($iTypeIdx -ge 0) { "$($iRow[$iTypeIdx])" } else { "Unknown" }
                $iMeter = if ($iMeterIdx -ge 0) { "$($iRow[$iMeterIdx])" } else { "Unknown" }
                if ($iCost -le 0) { continue }
                $daysInRange = ((Get-Date) - [datetime]$costStartDate).Days
                $iMonthly = [math]::Round(($iCost / [math]::Max(1, $daysInRange)) * 30, 2)
                $infraCostData.Add([PSCustomObject]@{
                  SubscriptionId = $subId
                  ResourceGroup  = $rgName
                  ResourceType   = $iType
                  MeterCategory  = $iMeter
                  Period30DayCost = [math]::Round($iCost, 2)
                  MonthlyEstimate = $iMonthly
                })
              }
            }
          }
        } catch {
          Write-Host "      ⚠ Infra cost query failed for RG $rgName — $($_.Exception.Message)" -ForegroundColor Yellow
        }
      }
    }
  }
  
  # Summary
  $vmsWithCosts = $vmActualMonthlyCost.Count
  $totalActualMonthly = ($actualCostData | ForEach-Object { $_.MonthlyEstimate } | Measure-Object -Sum).Sum
  $totalActualMonthly = if ($totalActualMonthly) { [math]::Round($totalActualMonthly, 0) } else { 0 }
  
  # Diagnostic: show match rate and help troubleshoot
  $totalVms = @($vms).Count
  $matchRate = if ($totalVms -gt 0) { [math]::Round(($vmsWithCosts / $totalVms) * 100, 0) } else { 0 }
  Write-Host "  Actual Costs: $vmsWithCosts of $totalVms VMs matched ($matchRate%), ~`$$totalActualMonthly/mo" -ForegroundColor $(if ($vmsWithCosts -eq $totalVms) { "Green" } elseif ($vmsWithCosts -gt 0) { "Yellow" } else { "Red" })
  
  # Show pricing model breakdown
  if ($actualCostData.Count -gt 0) {
    $riVms = @($actualCostData | Where-Object { $_.PricingModel -match 'Reserved|Savings' })
    $onDemandVms = @($actualCostData | Where-Object { $_.PricingModel -eq 'OnDemand' })
    if ($riVms.Count -gt 0) {
      $riMonthly = ($riVms | ForEach-Object { $_.MonthlyEstimate } | Measure-Object -Sum).Sum
      Write-Host "    Pricing: $($riVms.Count) VM(s) on RI/Savings Plan (~`$$([math]::Round($riMonthly, 0))/mo amortized), $($onDemandVms.Count) VM(s) on-demand" -ForegroundColor Green
      Write-Host "    Note: Using AmortizedCost — RI costs are spread across covered VMs (not $0)" -ForegroundColor Gray
    }
  }
  if ($vmsWithCosts -lt $totalVms -and $vmsWithCosts -gt 0) {
    $unmatchedVms = @($vms | Where-Object { -not $vmActualMonthlyCost.ContainsKey($_.VMName) })
    $unmatchedSubs = @($unmatchedVms | ForEach-Object { Get-SubFromArmId (TryGet-ArmId $_) } | Where-Object { $_ } | Select-Object -Unique)
    $unmatchedSubsNotQueried = @($unmatchedSubs | Where-Object { $_ -notin $costAccessGranted })
    if ($unmatchedSubsNotQueried.Count -gt 0) {
      Write-Host "    ⚠ $($unmatchedSubsNotQueried.Count) subscription(s) with unmatched VMs lack Cost Management access:" -ForegroundColor Yellow
      foreach ($us in $unmatchedSubsNotQueried) {
        $usVmCount = @($unmatchedVms | Where-Object { (Get-SubFromArmId (TryGet-ArmId $_)) -eq $us }).Count
        Write-Host "      $us ($usVmCount VMs) — needs Cost Management Reader role" -ForegroundColor Gray
      }
    }
    if (($unmatchedVms.Count - ($unmatchedSubsNotQueried | ForEach-Object { @($unmatchedVms | Where-Object { (Get-SubFromArmId (TryGet-ArmId $_)) -eq $_ }).Count } | Measure-Object -Sum).Sum) -gt 0) {
      Write-Host "    Tip: Some VMs in accessible subscriptions weren't matched. Check that VM resource IDs are consistent." -ForegroundColor Gray
      Write-Host "    Cost Management returned $($resourceCosts.Count) resource(s) across all subscriptions" -ForegroundColor Gray
    }
  }
  if ($vmsWithCosts -eq 0 -and $costAccessGranted.Count -gt 0) {
    Write-Host "    Cost Management API returned data but no VMs matched." -ForegroundColor Yellow
    Write-Host "    This usually means VMs are in a different subscription than the one(s) queried." -ForegroundColor Yellow
    Write-Host "    Subscriptions queried: $($costAccessGranted -join ', ')" -ForegroundColor Gray
    $vmSubs = @($vms | ForEach-Object { Get-SubFromArmId (TryGet-ArmId $_) } | Where-Object { $_ } | Select-Object -Unique)
    $missingSubs = @($vmSubs | Where-Object { $_ -notin $costAccessGranted })
    if ($missingSubs.Count -gt 0) {
      Write-Host "    VM subscriptions NOT queried: $($missingSubs -join ', ')" -ForegroundColor Red
      Write-Host "    → Add these to -SubscriptionIds or grant Cost Management Reader role" -ForegroundColor Yellow
    }
  }
  
  # Export
  if ($actualCostData.Count -gt 0) {
    $actualCostData | Export-Csv (Join-Path $outFolder "ENHANCED-Actual-Costs.csv") -NoTypeInformation
  }
  
  # Infrastructure cost summary
  if ($infraCostData.Count -gt 0) {
    $infraCostData | Export-Csv (Join-Path $outFolder "ENHANCED-Infrastructure-Costs.csv") -NoTypeInformation
    $infraTotal = [math]::Round(($infraCostData | ForEach-Object { $_.MonthlyEstimate } | Measure-Object -Sum).Sum, 0)
    $infraByType = $infraCostData | Group-Object MeterCategory | ForEach-Object { "$($_.Name): `$$([math]::Round(($_.Group | ForEach-Object { $_.MonthlyEstimate } | Measure-Object -Sum).Sum, 0))" }
    Write-Host "  Infrastructure Costs: ~`$$infraTotal/mo across AVD resource groups" -ForegroundColor Green
    foreach ($it in $infraByType) { Write-Host "    $it" -ForegroundColor Gray }
  }
}

# Helper: Get monthly cost for a VM — actual if available, PAYG estimate fallback
function Get-VmMonthlyCost {
  param([string]$VMName, [string]$VMSize, [string]$Region)
  
  if ($vmActualMonthlyCost.ContainsKey($VMName)) {
    $actual = $vmActualMonthlyCost[$VMName]
    return @{ Monthly = $actual; Source = "Actual"; ActualBilled = $actual; IsDeallocated = $false }
  }
  $payg = Get-EstimatedVmCostPerHour -VmSize $VMSize -Region $Region
  $paygMonthly = if ($payg) { [math]::Round($payg * 730, 2) } else { $null }
  if ($paygMonthly) {
    return @{ Monthly = $paygMonthly; Source = "Estimate"; ActualBilled = $null; IsDeallocated = $false }
  }
  return @{ Monthly = $null; Source = "Unknown"; ActualBilled = $null; IsDeallocated = $false }
}

# =========================================================
# Enhanced Analysis: SKU Diversity & Allocation Resilience (v4.0.0)
# =========================================================
Write-Host "Analyzing SKU diversity and allocation resilience..." -ForegroundColor Cyan

# Parse SKU family from VM size string (e.g., "Standard_D4ads_v5" → family "Dads", series "v5")
function Get-SkuFamilyInfo ([string]$vmSize) {
  $normalized = $vmSize -replace '^Standard_', '' -replace '^Basic_', ''
  # Match pattern: letter(s) + digits + optional suffix letters + _v(N)
  # E.g., D4ads_v5 → family=D, suffix=ads, gen=v5, cores=4
  # E.g., E8bds_v5 → family=E, suffix=bds, gen=v5, cores=8
  # E.g., A4_v2 → family=A, suffix='', gen=v2, cores=4
  if ($normalized -match '^([A-Z]+)(\d+)([a-z]*)_?(v\d+)?$') {
    $familyLetter = $Matches[1]
    $cores = $Matches[2]
    $suffix = $Matches[3]
    $generation = if ($Matches[4]) { $Matches[4] } else { "v1" }
    $familyFull = "$familyLetter$suffix"
    return @{
      FamilyLetter = $familyLetter
      FamilyFull   = $familyFull
      Cores        = [int]$cores
      Suffix       = $suffix
      Generation   = $generation
      SeriesName   = "${familyFull}_${generation}"
    }
  }
  return @{ FamilyLetter = "Unknown"; FamilyFull = "Unknown"; Cores = 0; Suffix = ""; Generation = "Unknown"; SeriesName = "Unknown" }
}

# Known SKU family compatibility groups — SKUs that can typically serve similar workloads
$skuCompatibilityGroups = @{
  "GeneralPurpose"    = @("D", "Das", "Dads", "Ds", "Dds", "Dpds", "Dplds", "Dps", "Dns", "Dnds")
  "MemoryOptimized"   = @("E", "Eas", "Eads", "Es", "Eds", "Ebds", "Epds", "Eps")
  "ComputeOptimized"  = @("F", "Fas", "Fs")
  "HighPerformance"   = @("H", "Hb")
  "StorageOptimized"  = @("L", "Las", "Ls")
  "GPUEnabled"        = @("N", "Nc", "Nv", "Nd")
}

$skuDiversityAnalysis = [System.Collections.Generic.List[object]]::new()

$hpVmGroups = $vms | Group-Object HostPoolName
foreach ($hpGroup in $hpVmGroups) {
  $hpName = $hpGroup.Name
  $hpVms = @($hpGroup.Group)
  $hpType = $hpTypeLookup[$hpName]
  $vmCount_hp = $hpVms.Count
  
  # Get unique SKUs and their families
  $skuDetails = @{}
  $familyCounts = @{}
  $seriesCounts = @{}
  $regionCounts = @{}
  $zoneCounts = @{}
  
  foreach ($v in $hpVms) {
    $info = Get-SkuFamilyInfo $v.VMSize
    $skuDetails[$v.VMSize] = $info
    $familyCounts[$info.FamilyLetter] = ($familyCounts[$info.FamilyLetter] ?? 0) + 1
    $seriesCounts[$info.SeriesName] = ($seriesCounts[$info.SeriesName] ?? 0) + 1
    $region = $v.Region ?? "Unknown"
    $regionCounts[$region] = ($regionCounts[$region] ?? 0) + 1
    $vmZone = if ($v.PSObject.Properties['Zones'] -and $v.Zones) { "$($v.Zones)" } else { "None" }
    $zoneCounts[$vmZone] = ($zoneCounts[$vmZone] ?? 0) + 1
  }
  
  $uniqueSkus = @($hpVms | Select-Object -ExpandProperty VMSize -Unique)
  $uniqueFamilies = @($familyCounts.Keys)
  $uniqueSeries = @($seriesCounts.Keys)
  $uniqueRegions = @($regionCounts.Keys)
  $dominantSku = ($hpVms | Group-Object VMSize | Sort-Object Count -Descending | Select-Object -First 1)
  $dominantPct = [math]::Round(($dominantSku.Count / $vmCount_hp) * 100, 0)
  
  # Determine risk level — Pooled pools always need diversity for scale-out; Personal pools less so
  $skuRisk = "Low"
  $regionRisk = "Low"
  $findings_sku = @()
  $recommendations_sku = @()
  $isPooledHP = ($hpType -eq "Pooled")
  
  # SKU concentration risk
  if ($uniqueFamilies.Count -eq 1) {
    if ($isPooledHP) {
      # Pooled: any single-family pool is at risk because it needs to scale
      $skuRisk = "High"
      $findings_sku += "$(if ($vmCount_hp -eq 1) { 'Single VM' } else { "All $vmCount_hp VMs" }) using one SKU family ($($uniqueFamilies[0])-series) — vulnerable to family-wide allocation failures during scale-out"
    } elseif ($vmCount_hp -gt 1) {
      # Personal with multiple VMs: still a risk
      $skuRisk = "High"
      $findings_sku += "All $vmCount_hp VMs use the same SKU family ($($uniqueFamilies[0])-series) — vulnerable to family-wide allocation failures"
    } else {
      # Personal single VM: low risk (can't diversify a single desktop)
      $skuRisk = "Low"
      $findings_sku += "Single personal desktop — SKU diversity not applicable"
    }
    
    if ($skuRisk -ne "Low") {
      # Suggest alternatives based on compatibility group
      $currentFamily = $uniqueFamilies[0]
      $currentInfo = $skuDetails[$dominantSku.Name]
      $currentGroup = $null
      foreach ($group in $skuCompatibilityGroups.GetEnumerator()) {
        if ($group.Value -contains $currentInfo.FamilyFull) {
          $currentGroup = $group.Key
          break
        }
      }
      if (-not $currentGroup) {
        foreach ($group in $skuCompatibilityGroups.GetEnumerator()) {
          if ($group.Value -contains $currentFamily) {
            $currentGroup = $group.Key
            break
          }
        }
      }
      
      if ($currentGroup) {
        $alternatives = @($skuCompatibilityGroups[$currentGroup] | Where-Object { $_ -ne $currentInfo.FamilyFull -and $_ -ne $currentFamily })
        if ($alternatives.Count -gt 0) {
          $altExamples = @()
          foreach ($alt in ($alternatives | Select-Object -First 3)) {
            $altExamples += "Standard_${alt}$($currentInfo.Cores)_$($currentInfo.Generation)"
          }
          $recommendations_sku += "Consider mixing in alternative SKUs from the same workload class ($currentGroup): $($altExamples -join ', ')"
        }
      }
      if ($isPooledHP) {
        $recommendations_sku += "Use 2+ SKU families in autoscale configuration so allocation failures on one family don't block scaling"
      } else {
        $recommendations_sku += "Split host pool across 2+ SKU families so allocation failures on one family don't prevent provisioning"
      }
    }
  }
  elseif ($uniqueSeries.Count -eq 1 -and $vmCount_hp -gt 1) {
    $skuRisk = "Medium"
    $findings_sku += "All VMs use the same series ($($uniqueSeries[0])) — single series allocation failure could impact scaling"
    $recommendations_sku += "Mix in VMs from a different generation or suffix variant for resilience"
  }
  elseif ($dominantPct -ge 80 -and $vmCount_hp -gt 3) {
    $skuRisk = "Medium"
    $findings_sku += "$dominantPct% of VMs use $($dominantSku.Name) — consider diversifying for allocation resilience"
    $recommendations_sku += "Reduce dependency on single SKU to below 70% of fleet"
  }
  else {
    $findings_sku += "Good SKU diversity across $($uniqueFamilies.Count) families and $($uniqueSeries.Count) series"
  }
  
  # Region concentration risk
  if ($uniqueRegions.Count -eq 1) {
    if ($isPooledHP) {
      $regionRisk = "High"
      $findings_sku += "All VMs deployed in single region ($($uniqueRegions[0])) — no geo-redundancy for regional capacity or outage events"
      $recommendations_sku += "Deploy a secondary host pool in a paired region for disaster recovery and allocation resilience"
    } elseif ($vmCount_hp -gt 1) {
      $regionRisk = if ($vmCount_hp -gt 5) { "High" } else { "Medium" }
      $findings_sku += "All VMs deployed in single region ($($uniqueRegions[0])) — no geo-redundancy"
      $recommendations_sku += "For mission-critical workloads, deploy a secondary host pool in a paired region"
    }
  }
  
  $overallRisk = if ($skuRisk -eq "High" -or $regionRisk -eq "High") { "High" }
                 elseif ($skuRisk -eq "Medium" -or $regionRisk -eq "Medium") { "Medium" }
                 else { "Low" }
  
  $uniqueZones = @($zoneCounts.Keys | Where-Object { $_ -ne "None" })
  $zoneList = if ($uniqueZones.Count -gt 0) {
    ($zoneCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name) ($($_.Value))" }) -join ", "
  } else { "No zones" }
  
  $skuDiversityAnalysis.Add([PSCustomObject]@{
    HostPoolName       = $hpName
    HostPoolType       = $hpType
    VMCount            = $vmCount_hp
    UniqueSkus         = $uniqueSkus.Count
    UniqueFamilies     = $uniqueFamilies.Count
    UniqueSeries       = $uniqueSeries.Count
    UniqueRegions      = $uniqueRegions.Count
    UniqueZones        = $uniqueZones.Count
    DominantSku        = $dominantSku.Name
    DominantSkuPct     = $dominantPct
    SkuList            = ($uniqueSkus -join ", ")
    FamilyList         = ($uniqueFamilies -join ", ")
    SeriesList         = ($uniqueSeries -join ", ")
    RegionList         = ($uniqueRegions -join ", ")
    ZoneList           = $zoneList
    SkuRisk            = $skuRisk
    RegionRisk         = $regionRisk
    OverallRisk        = $overallRisk
    Findings           = ($findings_sku -join "; ")
    Recommendations    = ($recommendations_sku -join "; ")
  })
}

$highRiskPools = @($skuDiversityAnalysis | Where-Object { $_.OverallRisk -eq "High" })
$medRiskPools = @($skuDiversityAnalysis | Where-Object { $_.OverallRisk -eq "Medium" })
$singleRegionPools = @($skuDiversityAnalysis | Where-Object { $_.UniqueRegions -eq 1 -and $_.VMCount -gt 1 })

Write-Host "  SKU Diversity: $($highRiskPools.Count) high-risk, $($medRiskPools.Count) medium-risk host pools" -ForegroundColor $(if ($highRiskPools.Count -gt 0) { "Red" } elseif ($medRiskPools.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Geo-Redundancy: $($singleRegionPools.Count) host pools in single region" -ForegroundColor $(if ($singleRegionPools.Count -gt 0) { "Yellow" } else { "Green" })

# =========================================================
# Enhanced Analysis: W365 Cloud PC Readiness (v4.0.0)
# =========================================================
Write-Host "Analyzing W365 Cloud PC readiness..." -ForegroundColor Cyan

# W365 Cloud PC pricing (monthly, East US baseline)
# Enterprise plans - per user/month (dedicated 1:1)
# Frontline plans - per concurrent user/month (shared, for Cloud Apps / shift workers)
$w365Pricing = @{
  # W365 Enterprise plans
  "W365_2vCPU_4GB_128GB"   = @{ vCPU = 2;  RAM = 4;   Storage = 128; Monthly = 31.00;  Name = "W365 Enterprise 2vCPU/4GB/128GB";  LicenseType = "Enterprise" }
  "W365_2vCPU_4GB_256GB"   = @{ vCPU = 2;  RAM = 4;   Storage = 256; Monthly = 39.00;  Name = "W365 Enterprise 2vCPU/4GB/256GB";  LicenseType = "Enterprise" }
  "W365_2vCPU_8GB_128GB"   = @{ vCPU = 2;  RAM = 8;   Storage = 128; Monthly = 38.00;  Name = "W365 Enterprise 2vCPU/8GB/128GB";  LicenseType = "Enterprise" }
  "W365_2vCPU_8GB_256GB"   = @{ vCPU = 2;  RAM = 8;   Storage = 256; Monthly = 46.00;  Name = "W365 Enterprise 2vCPU/8GB/256GB";  LicenseType = "Enterprise" }
  "W365_4vCPU_16GB_128GB"  = @{ vCPU = 4;  RAM = 16;  Storage = 128; Monthly = 58.00;  Name = "W365 Enterprise 4vCPU/16GB/128GB"; LicenseType = "Enterprise" }
  "W365_4vCPU_16GB_256GB"  = @{ vCPU = 4;  RAM = 16;  Storage = 256; Monthly = 66.00;  Name = "W365 Enterprise 4vCPU/16GB/256GB"; LicenseType = "Enterprise" }
  "W365_4vCPU_16GB_512GB"  = @{ vCPU = 4;  RAM = 16;  Storage = 512; Monthly = 82.00;  Name = "W365 Enterprise 4vCPU/16GB/512GB"; LicenseType = "Enterprise" }
  "W365_8vCPU_32GB_128GB"  = @{ vCPU = 8;  RAM = 32;  Storage = 128; Monthly = 99.00;  Name = "W365 Enterprise 8vCPU/32GB/128GB"; LicenseType = "Enterprise" }
  "W365_8vCPU_32GB_256GB"  = @{ vCPU = 8;  RAM = 32;  Storage = 256; Monthly = 107.00; Name = "W365 Enterprise 8vCPU/32GB/256GB"; LicenseType = "Enterprise" }
  "W365_8vCPU_32GB_512GB"  = @{ vCPU = 8;  RAM = 32;  Storage = 512; Monthly = 123.00; Name = "W365 Enterprise 8vCPU/32GB/512GB"; LicenseType = "Enterprise" }
  "W365_16vCPU_64GB_256GB" = @{ vCPU = 16; RAM = 64;  Storage = 256; Monthly = 184.00; Name = "W365 Enterprise 16vCPU/64GB/256GB"; LicenseType = "Enterprise" }
  "W365_16vCPU_64GB_512GB" = @{ vCPU = 16; RAM = 64;  Storage = 512; Monthly = 200.00; Name = "W365 Enterprise 16vCPU/64GB/512GB"; LicenseType = "Enterprise" }
  # W365 Frontline plans (Shared mode — for Cloud Apps / RemoteApp replacement)
  # Pricing is per concurrent user; Cloud PCs are shared across shifts
  "W365FL_2vCPU_4GB_128GB"  = @{ vCPU = 2;  RAM = 4;   Storage = 128; Monthly = 17.00;  Name = "W365 Frontline 2vCPU/4GB/128GB";  LicenseType = "Frontline" }
  "W365FL_2vCPU_8GB_128GB"  = @{ vCPU = 2;  RAM = 8;   Storage = 128; Monthly = 21.00;  Name = "W365 Frontline 2vCPU/8GB/128GB";  LicenseType = "Frontline" }
  "W365FL_4vCPU_16GB_128GB" = @{ vCPU = 4;  RAM = 16;  Storage = 128; Monthly = 32.00;  Name = "W365 Frontline 4vCPU/16GB/128GB"; LicenseType = "Frontline" }
  "W365FL_4vCPU_16GB_256GB" = @{ vCPU = 4;  RAM = 16;  Storage = 256; Monthly = 40.00;  Name = "W365 Frontline 4vCPU/16GB/256GB"; LicenseType = "Frontline" }
  "W365FL_8vCPU_32GB_128GB" = @{ vCPU = 8;  RAM = 32;  Storage = 128; Monthly = 55.00;  Name = "W365 Frontline 8vCPU/32GB/128GB"; LicenseType = "Frontline" }
  "W365FL_8vCPU_32GB_256GB" = @{ vCPU = 8;  RAM = 32;  Storage = 256; Monthly = 63.00;  Name = "W365 Frontline 8vCPU/32GB/256GB"; LicenseType = "Frontline" }
}

$w365Analysis = [System.Collections.Generic.List[object]]::new()

# Pre-extract connection success data for user counts (needed in W365 cost comparison)
if (-not (Test-Path variable:connSuccessData)) {
  $connSuccessData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionSuccessRate" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "HostPool" })
}

foreach ($hp in $hostPools) {
  $hpName = $hp.HostPoolName
  $hpType = $hp.HostPoolType
  if (-not $hpName) { continue }
  
  $isPersonal = ($hpType -match "Personal")
  $isPooled = ($hpType -match "Pooled")
  $maxSess = [int]$hp.MaxSessions
  $appGroupType = $hp.PreferredAppGroupType
  $isRemoteApp = ($appGroupType -eq "RailApplications")
  
  # Get VMs in this host pool
  $hpVms = @($vms | Where-Object { $_.HostPoolName -eq $hpName })
  $vmCount_w = $hpVms.Count
  if ($vmCount_w -eq 0) { continue }
  
  # Get session host data
  $hpSessionHosts = @($sessionHosts | Where-Object { $_.HostPoolName -eq $hpName })
  $assignedUsers = @($hpSessionHosts | Where-Object { $_.AssignedUser } | ForEach-Object { $_.AssignedUser } | Select-Object -Unique)
  
  # Get active sessions
  $sessSum = ($hpSessionHosts | ForEach-Object { if ($_.ActiveSessions) { [int]$_.ActiveSessions } else { 0 } } | Measure-Object -Sum).Sum
  $totalActiveSessions = if ($sessSum) { $sessSum } else { 0 }
  $avgSessionsPerHost = if ($vmCount_w -gt 0) { [math]::Round($totalActiveSessions / $vmCount_w, 1) } else { 0 }
  
  # Dominant VM size info
  $dominantSize = @($hpVms | Where-Object { $_.VMSize } | Group-Object VMSize | Sort-Object Count -Descending | Select-Object -First 1)
  $vmSize = if ($dominantSize.Count -gt 0) { $dominantSize[0].Name } else { "Unknown" }
  if (-not $vmSize -or $vmSize -eq "Unknown") { continue }
  $skuInfo = Get-SkuFamilyInfo $vmSize
  $vCPU = if ($skuInfo.Cores -gt 0) { $skuInfo.Cores } else { 4 }
  # RAM estimate: E-series = 8GB/vCPU, D-series = 4GB/vCPU, B-series = 4GB/vCPU
  $ramGB = if ($skuInfo.FamilyLetter -eq "E") { $vCPU * 8 }
           elseif ($skuInfo.FamilyLetter -eq "M") { $vCPU * 16 }
           else { $vCPU * 4 }
  
  # Current monthly cost per VM — actual if available, PAYG estimate fallback
  $costInfo = Get-VmMonthlyCost -VMName ($hpVms[0].VMName) -VMSize $vmSize -Region $hp.Location
  $monthlyPerVm = $costInfo.Monthly
  $costSource = $costInfo.Source
  
  # Initialize fit analysis lists early (used in cost sanity check below)
  $fitFactors = [System.Collections.Generic.List[string]]::new()
  $fitBlockers = [System.Collections.Generic.List[string]]::new()
  $fitAdvantages = [System.Collections.Generic.List[string]]::new()
  
  # If actual costs available, get per-pool total from actual data
  # Use Get-VmMonthlyCost which handles deallocated VMs (returns PAYG "running cost" if actual < 10% of PAYG)
  if ((-not $SkipActualCosts) -and $actualCostData.Count -gt 0) {
    $poolActualCosts = @($actualCostData | Where-Object { $_.HostPoolName -eq $hpName })
    if ($poolActualCosts.Count -gt 0) {
      $poolActualMonthly = ($poolActualCosts | ForEach-Object { $_.MonthlyEstimate } | Measure-Object -Sum).Sum
      if ($poolActualMonthly -and $poolActualMonthly -gt 0) {
        $monthlyPerVm = [math]::Round($poolActualMonthly / $vmCount_w, 2)
        $costSource = "Actual"
      }
    }
  }
  
  # Check if this pool has a scaling plan
  $hasScalingPlan = @($scalingPlanAssignments | Where-Object { $_.HostPoolName -eq $hpName -or $_.HostPoolArmId -match $hpName }).Count -gt 0
  
  # Check for GPU SKUs (W365 doesn't support these well)
  $hasGpu = @($hpVms | Where-Object { $_.VMSize -match '_NC|_ND|_NV|_NP' }).Count -gt 0
  
  # Check for large VMs (>16 vCPU) — exceeds W365 max
  $oversizedVms = @($hpVms | Where-Object { 
    $si = Get-SkuFamilyInfo $_.VMSize
    $si.Cores -gt 16
  })
  
  # Determine W365 fit score (0-100)
  $fitScore = 50  # Start neutral
  
  # --- Blockers (hard no) ---
  if ($hasGpu) {
    $fitBlockers.Add("GPU workloads — W365 GPU options are very limited")
    $fitScore = 0
  }
  if ($oversizedVms.Count -gt 0) {
    $fitBlockers.Add("$($oversizedVms.Count) VM(s) exceed W365 max of 16 vCPU/64 GB")
    $fitScore = [math]::Max(0, $fitScore - 50)
  }
  if ($isRemoteApp) {
    # W365 Cloud Apps (GA) enables published app delivery on Frontline Shared mode
    # with User Experience Sync for session persistence. However, it requires
    # Frontline licensing, doesn't support MSIX/Appx/Store apps (e.g., Teams),
    # and concurrent sessions are capped by license count per policy.
    $fitScore += 5
    $fitFactors.Add("RemoteApp pool — W365 Cloud Apps (Frontline Shared) can deliver published apps with User Experience Sync, but requires Frontline licensing model and doesn't support MSIX/Appx apps (e.g., Teams)")
    if ($maxSess -gt 2) {
      # Multi-session RemoteApp: W365 Cloud Apps is single-user per Cloud PC,
      # so you'd need 1 Frontline license per concurrent user vs shared multi-session hosts
      $fitFactors.Add("Multi-session pool ($maxSess sessions/host) — W365 Cloud Apps uses 1 Cloud PC per concurrent user; compare Frontline license cost × concurrent users vs current multi-session IaaS cost")
    }
  }
  if ($isPooled -and $maxSess -gt 2 -and -not $isRemoteApp) {
    $fitBlockers.Add("Multi-session pooled desktop (max $maxSess sessions/host) — W365 is single-user only")
    $fitScore = 0
  }
  
  # --- Strong signals toward W365 ---
  if ($isPersonal) {
    $fitScore += 25
    $fitAdvantages.Add("Personal host pool — W365 is purpose-built for 1:1 user-to-VM assignment")
  }
  if ($isPooled -and $maxSess -le 2 -and $avgSessionsPerHost -le 1.5) {
    $fitScore += 20
    $fitAdvantages.Add("Pooled pool running at effectively 1:1 density ($avgSessionsPerHost sessions/host)")
  }
  if (-not $hasScalingPlan -and $isPersonal) {
    $fitScore += 10
    $fitAdvantages.Add("No scaling plan — VMs run always-on (W365 fixed cost may be cheaper)")
  }
  if ($vmCount_w -le 10) {
    $fitScore += 5
    $fitAdvantages.Add("Small pool ($vmCount_w VMs) — AVD infrastructure overhead is proportionally high")
  }
  
  # --- Signals favoring AVD ---
  if ($isPooled -and $avgSessionsPerHost -gt 3) {
    $fitScore -= 20
    $fitFactors.Add("High session density ($avgSessionsPerHost users/host) — multi-session is AVD's advantage")
  }
  if ($hasScalingPlan) {
    $fitScore -= 10
    $fitFactors.Add("Scaling plan attached — elastic scaling reduces AVD cost during off-hours")
  }
  if ($vmCount_w -gt 50) {
    $fitScore -= 5
    $fitFactors.Add("Large pool ($vmCount_w VMs) — AVD management overhead is amortized across fleet")
  }

  # --- Session duration signal (if KQL data available) ---
  $poolSessionDur = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_SessionDuration" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "AvgDuration" -and $_.AvgDuration })
  if ($poolSessionDur.Count -gt 0) {
    $globalAvgDur = [math]::Round(($poolSessionDur | ForEach-Object { [double]$_.AvgDuration } | Measure-Object -Average).Average, 0)
    $shortSessionPct = [math]::Round((@($poolSessionDur | Where-Object { [double]$_.AvgDuration -lt 120 }).Count / $poolSessionDur.Count * 100), 0)
    if ($shortSessionPct -ge 50 -and $isRemoteApp) {
      $fitScore += 10
      $fitAdvantages.Add("$shortSessionPct% of users average <2hr sessions — ideal for Frontline shared licensing (pay per concurrent)")
    } elseif ($shortSessionPct -ge 50 -and $isPersonal) {
      $fitScore += 5
      $fitAdvantages.Add("$shortSessionPct% of users average <2hr sessions — VMs idle most of the day; W365 eliminates always-on compute waste")
    } elseif ($globalAvgDur -ge 360 -and $isPersonal) {
      $fitAdvantages.Add("Full-day usage pattern (avg ${globalAvgDur}m) — W365 Enterprise fixed cost is predictable and eliminates scaling complexity")
    }
  }

  # --- Migration complexity assessment ---
  $migrationFactors = [System.Collections.Generic.List[string]]::new()
  $migrationComplexity = "Low"
  $complexityScore = 0

  # FSLogix / profile dependency
  $hpProfileData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ProfileLoadPerformance" -and $_.QueryName -eq "AVD" })
  if ($hpProfileData.Count -gt 0 -and $isPersonal) {
    $migrationFactors.Add("FSLogix profiles in use — W365 is persistent, so profiles migrate once (not ongoing)")
    $complexityScore += 1
  } elseif ($hpProfileData.Count -gt 0 -and $isPooled) {
    $migrationFactors.Add("FSLogix profiles on pooled hosts — W365 eliminates profile containers entirely")
    $complexityScore += 0  # This is actually a benefit
  }

  # Custom images
  $hpImages = @($hpVms | Where-Object { $_.ImagePublisher })
  $hasCustomImage = @($hpImages | Where-Object { -not $_.ImagePublisher -or $_.ImagePublisher -match 'gallery|custom|image' -or [string]::IsNullOrEmpty($_.ImagePublisher) }).Count -gt 0
  if (-not $hasCustomImage) {
    # Also check for gallery image IDs (no publisher = custom/gallery image)
    $hasCustomImage = @($hpVms | Where-Object { -not $_.ImagePublisher -or $_.ImagePublisher -eq "" }).Count -gt 0
  }
  if ($hasCustomImage) {
    $migrationFactors.Add("Custom/gallery image — must rebuild for Intune deployment (Azure Image Builder → Intune Win32/LOB)")
    $complexityScore += 2
  }

  # Domain join type
  $hpSessionHosts = @($sessionHosts | Where-Object { $_.HostPoolName -eq $hpName })
  $hasAadJoin = @($hpSessionHosts | Where-Object { $_.PSObject.Properties['ResourceId'] -and $_.ResourceId -match 'Microsoft.DesktopVirtualization' }).Count -gt 0

  # Pool size affects migration effort
  if ($vmCount_w -gt 100) {
    $migrationFactors.Add("Large pool ($vmCount_w VMs) — phased migration recommended with pilot group")
    $complexityScore += 2
  } elseif ($vmCount_w -gt 30) {
    $migrationFactors.Add("Medium pool ($vmCount_w VMs) — batch migration in waves")
    $complexityScore += 1
  }

  $migrationComplexity = if ($complexityScore -ge 4) { "High" } elseif ($complexityScore -ge 2) { "Medium" } else { "Low" }
  
  # Clamp
  $fitScore = [math]::Max(0, [math]::Min(100, $fitScore))
  
  # --- Find best W365 plan match ---
  # RemoteApp pools → prefer Frontline Shared plans (Cloud Apps with UX Sync)
  # Personal/Desktop pools → prefer Enterprise plans (dedicated 1:1)
  $preferredLicenseType = if ($isRemoteApp) { "Frontline" } else { "Enterprise" }
  
  # For pooled multi-session pools: W365 replaces the per-USER experience, not the whole VM
  # So match on per-user resource needs, not total VM spec
  # For personal pools: 1:1 mapping, match on full VM spec
  $matchvCPU = $vCPU
  $matchRAM = $ramGB
  if ($isPooled -and $maxSess -gt 1) {
    # Estimate per-user resources based on workload type
    $estUsersPerVm = if ($totalActiveSessions -gt 0 -and $vmCount_w -gt 0) { 
      [math]::Max(1, [math]::Ceiling($totalActiveSessions / $vmCount_w))
    } elseif ($maxSess -gt 0) { 
      [math]::Min($maxSess, [math]::Max(1, [math]::Floor($vCPU * 4)))  # ~4 sessions/vCPU
    } else { 4 }
    
    if ($isRemoteApp) {
      # RemoteApp users consume fewer resources — typically 1-2 apps, not a full desktop
      # Microsoft guidance: RemoteApp sessions ~1 vCPU, 2-4 GB RAM per user
      $perUserCPU = [math]::Max(2, [math]::Ceiling($vCPU / $estUsersPerVm))
      $perUserRAM = [math]::Max(2, [math]::Ceiling($ramGB / $estUsersPerVm))
      # RemoteApp per-user needs are typically lower — cap at reasonable limits
      $perUserCPU = [math]::Min($perUserCPU, 4)
      $perUserRAM = [math]::Min($perUserRAM, 8)
      $fitFactors.Add("RemoteApp pool: matching W365 Frontline on per-user app needs ($perUserCPU vCPU / $($perUserRAM) GB) — lighter workload than full desktop")
    } else {
      # Pooled Desktop: full desktop experience per user, but shared VM
      # Microsoft guidance: 2-4 vCPU, 4-8 GB RAM per knowledge worker
      $perUserCPU = [math]::Max(2, [math]::Ceiling($vCPU / $estUsersPerVm))
      $perUserRAM = [math]::Max(4, [math]::Ceiling($ramGB / $estUsersPerVm))
      $fitFactors.Add("Pooled desktop: matching W365 on per-user needs ($perUserCPU vCPU / $($perUserRAM) GB per user) not total VM spec ($vCPU vCPU / $($ramGB) GB)")
    }
    
    $matchvCPU = $perUserCPU
    $matchRAM = $perUserRAM
  } elseif ($isPersonal) {
    # Personal pool: 1:1 mapping, user gets the whole VM
    # But still flag if the VM is oversized for a single user
    if ($ramGB -gt 32) {
      $fitFactors.Add("Personal pool VM has $($ramGB) GB RAM — W365 max is 64 GB. Matching on full VM spec.")
    }
  }
  
  $bestW365Plan = $null
  $bestW365Monthly = [double]::MaxValue
  foreach ($plan in $w365Pricing.Values) {
    # Filter by preferred license type first
    if ($plan.LicenseType -ne $preferredLicenseType) { continue }
    if ($plan.vCPU -ge $matchvCPU -and $plan.RAM -ge $matchRAM) {
      if ($plan.Monthly -lt $bestW365Monthly) {
        $bestW365Monthly = $plan.Monthly
        $bestW365Plan = $plan
      }
    }
  }
  # Fallback: try any license type if no preferred match
  if (-not $bestW365Plan) {
    foreach ($plan in $w365Pricing.Values) {
      if ($plan.vCPU -ge $matchvCPU -and $plan.RAM -ge $matchRAM) {
        if ($plan.Monthly -lt $bestW365Monthly) {
          $bestW365Monthly = $plan.Monthly
          $bestW365Plan = $plan
        }
      }
    }
  }
  # If no exact match (VM is larger than biggest W365), try closest smaller fit
  if (-not $bestW365Plan -and $fitScore -gt 0) {
    # Find largest W365 plan of preferred type
    $bestW365Plan = $w365Pricing.Values | Where-Object { $_.LicenseType -eq $preferredLicenseType } | Sort-Object { $_.Monthly } -Descending | Select-Object -First 1
    if (-not $bestW365Plan) {
      $bestW365Plan = $w365Pricing.Values | Sort-Object { $_.Monthly } -Descending | Select-Object -First 1
    }
    $bestW365Monthly = $bestW365Plan.Monthly
    $fitFactors.Add("Current VM spec ($vCPU vCPU / $($ramGB) GB) exceeds closest W365 plan — may need to downsize or accept reduced spec")
    $fitScore = [math]::Max(0, $fitScore - 15)
  }
  
  # --- Cost comparison ---
  # W365 is licensed per USER, not per VM. This is the key distinction:
  #   - Personal pools: 1 user per VM, so user count ≈ VM count (but check assigned users)
  #   - Pooled pools: Many users per VM, so user count >> VM count
  #   - Frontline: Licensed per concurrent user, not total user
  
  # Determine the actual user count for this pool
  $poolUniqueUsers = $assignedUsers.Count  # From session host assigned users (Personal pools)
  
  # For pooled pools, try to get unique user count from Log Analytics
  if ($isPooled -or $poolUniqueUsers -eq 0) {
    # KQL uses split(_ResourceId, '/')[-1] which gives the short resource name
    # $hpName might be full name or short — try both exact and fuzzy matching
    $poolConnData = @($connSuccessData | Where-Object { 
      $connHp = $_.HostPool
      $connHp -eq $hpName -or $connHp -eq ($hpName -split '/')[-1] -or $hpName -match [regex]::Escape($connHp)
    })
    if ($poolConnData.Count -gt 0 -and $poolConnData[0].PSObject.Properties.Name -contains 'UniqueUsers') {
      $laUniqueUsers = [int]$poolConnData[0].UniqueUsers
      if ($laUniqueUsers -gt 0) { $poolUniqueUsers = $laUniqueUsers }
    }
  }
  # Fallback: if still 0, use active sessions as lower bound, or VM count
  if ($poolUniqueUsers -eq 0) {
    $poolUniqueUsers = if ($totalActiveSessions -gt 0) { $totalActiveSessions } else { $vmCount_w }
  }
  
  # Peak concurrent users (for Frontline licensing)
  $peakConcurrentUsers = $totalActiveSessions
  if ($peakConcurrentUsers -eq 0) { $peakConcurrentUsers = $vmCount_w }
  
  $w365LicenseCount = 0
  if ($isRemoteApp -and $bestW365Plan -and $bestW365Plan.LicenseType -eq "Frontline") {
    # Frontline: 1 license per concurrent user
    $w365LicenseCount = $peakConcurrentUsers
    $w365MonthlyTotal = [math]::Round($bestW365Monthly * $w365LicenseCount, 0)
    $fitFactors.Add("Frontline cost based on $w365LicenseCount concurrent users × `$$bestW365Monthly/user/mo")
  } elseif ($isPooled) {
    # Pooled Desktop: Every unique user needs their own W365 Cloud PC
    $w365LicenseCount = $poolUniqueUsers
    $w365MonthlyTotal = if ($bestW365Plan) { [math]::Round($bestW365Monthly * $w365LicenseCount, 0) } else { $null }
    if ($poolUniqueUsers -gt $vmCount_w) {
      $fitFactors.Add("Pooled pool serves $poolUniqueUsers unique users on $vmCount_w VMs — W365 requires $poolUniqueUsers licenses ($poolUniqueUsers × `$$bestW365Monthly = `$$w365MonthlyTotal/mo)")
    }
  } else {
    # Personal: 1:1 mapping, use assigned users or VM count
    $w365LicenseCount = if ($poolUniqueUsers -gt 0) { $poolUniqueUsers } else { $vmCount_w }
    $w365MonthlyTotal = if ($bestW365Plan) { [math]::Round($bestW365Monthly * $w365LicenseCount, 0) } else { $null }
  }
  $avdMonthlyTotal = if ($monthlyPerVm) { [math]::Round($monthlyPerVm * $vmCount_w, 0) } else { $null }
  
  # If actual costs are available, they already reflect scaling plan savings (fewer running hours = lower bill)
  # Only apply the 60% discount heuristic to PAYG estimates
  $avdEffectiveMonthly = if ($costSource -eq "Actual") { $avdMonthlyTotal }
                         elseif ($hasScalingPlan -and $avdMonthlyTotal) { [math]::Round($avdMonthlyTotal * 0.6, 0) }
                         else { $avdMonthlyTotal }
  
  $costDelta = if ($w365MonthlyTotal -and $avdEffectiveMonthly) { $w365MonthlyTotal - $avdEffectiveMonthly } else { $null }
  $costVerdict = if (-not $costDelta) { "Unable to estimate" }
                 elseif ($costDelta -lt -50) { "W365 likely cheaper" }
                 elseif ($costDelta -lt 50)  { "Comparable cost" }
                 else { "AVD likely cheaper" }
  
  if ($costDelta -and $costDelta -lt -50) {
    $fitScore += 10
    $fitAdvantages.Add("W365 estimated ~`$$([math]::Abs($costDelta))/mo cheaper than current AVD IaaS spend")
  } elseif ($costDelta -and $costDelta -gt 200) {
    $fitScore -= 10
    $fitFactors.Add("AVD estimated ~`$$costDelta/mo cheaper — scaling plan and multi-session economics favor IaaS")
  }
  
  # Final clamp and recommendation
  $fitScore = [math]::Max(0, [math]::Min(100, $fitScore))
  $recommendation = if ($fitBlockers.Count -gt 0 -and $fitScore -eq 0) { "Keep AVD" }
                     elseif ($fitScore -ge 70) { "Strong W365 Candidate" }
                     elseif ($fitScore -ge 45) { "Consider W365 / Hybrid" }
                     else { "Keep AVD" }
  
  # --- Breakeven analysis (at what utilization does AVD become cheaper?) ---
  $breakevenPct = $null
  if ($bestW365Plan -and $avdMonthlyTotal -and $avdMonthlyTotal -gt 0 -and $w365MonthlyTotal -and $w365MonthlyTotal -gt 0) {
    # AVD effective cost = PAYG cost × utilization%. W365 = fixed.
    # Breakeven: PAYG × X% = W365 → X = W365 / PAYG
    $breakevenPct = [math]::Round($w365MonthlyTotal / $avdMonthlyTotal * 100, 0)
    if ($breakevenPct -gt 100) { $breakevenPct = $null }  # W365 always more expensive
  }

  # === ENHANCEMENT 1: Usage-based W365 SKU right-sizing ===
  # Instead of matching W365 plan to current VM SKU, match to actual user workload
  $usageBasedPlan = $null
  $usageBasedMonthly = $null
  $usageSavingsVsCurrent = $null
  $poolRsData = @($vmRightSizing | Where-Object { $_.HostPoolName -eq $hpName -and $_.AvgCPU -gt 0 })
  if ($poolRsData.Count -gt 0 -and ($isPersonal -or ($isPooled -and $maxSess -le 2))) {
    # For 1:1 pools: use actual per-VM metrics to find the smallest W365 plan that fits
    $poolAvgCPU = [math]::Round(($poolRsData | ForEach-Object { $_.AvgCPU } | Measure-Object -Average).Average, 1)
    $poolPeakCPU = [math]::Round(($poolRsData | ForEach-Object { $_.PeakCPU } | Measure-Object -Maximum).Maximum, 1)
    $poolAvgMem = [math]::Round(($poolRsData | ForEach-Object { $_.AvgMemoryUsedGB } | Measure-Object -Average).Average, 1)
    $poolPeakMem = [math]::Round(($poolRsData | ForEach-Object { $_.PeakMemoryUsedGB } | Measure-Object -Maximum).Maximum, 1)
    
    # Determine minimum W365 spec needed based on usage (with 30% headroom on peak)
    $neededCPU = [math]::Max(2, [math]::Ceiling(($poolPeakCPU / 100 * $vCPU) * 1.3))
    $neededRAM = [math]::Max(4, [math]::Ceiling($poolPeakMem * 1.3))
    
    # Find cheapest W365 plan that meets the usage requirements
    $preferredLicense = if ($isRemoteApp) { "Frontline" } else { "Enterprise" }
    $usageCandidates = @($w365Pricing.GetEnumerator() | Where-Object {
      $_.Value.vCPU -ge $neededCPU -and $_.Value.RAM -ge $neededRAM -and
      $_.Value.LicenseType -eq $preferredLicense
    } | Sort-Object { $_.Value.Monthly })
    if ($usageCandidates.Count -eq 0) {
      $usageCandidates = @($w365Pricing.GetEnumerator() | Where-Object {
        $_.Value.vCPU -ge $neededCPU -and $_.Value.RAM -ge $neededRAM
      } | Sort-Object { $_.Value.Monthly })
    }
    if ($usageCandidates.Count -gt 0) {
      $usageBasedPlan = $usageCandidates[0].Value
      $usageBasedMonthly = $usageBasedPlan.Monthly
      if ($bestW365Plan -and $usageBasedMonthly -lt $bestW365Monthly) {
        $usageSavingsVsCurrent = [math]::Round(($bestW365Monthly - $usageBasedMonthly) * $vmCount_w, 0)
        $fitFactors.Add("Usage-based SKU: actual workload ($poolAvgCPU% avg CPU, $($poolPeakMem) GB peak RAM) fits $($usageBasedPlan.vCPU)vCPU/$($usageBasedPlan.RAM)GB — saves `$$usageSavingsVsCurrent/mo vs spec-matched plan")
      }
    }
  }

  # === ENHANCEMENT 2: TCO components (beyond compute) ===
  # Estimate AVD hidden costs that W365 eliminates
  $tcoStorageCost = 0
  $tcoProfileCost = 0
  $tcoNetworkCost = 0
  
  # OS disk costs (W365 includes storage)
  $poolDisks = @($storageFindingsList | Where-Object { $_.HostPoolName -eq $hpName })
  if ($poolDisks.Count -gt 0) {
    foreach ($disk in $poolDisks) {
      $diskMonthlyCost = switch -Wildcard ($disk.OSDiskType) {
        "*Premium*"  { 19.71 }  # P10 128 GB
        "*StandardSSD*" { 7.68 }   # E10 128 GB
        "*Standard_LRS*" { 5.89 }  # S10 128 GB
        default { 7.68 }
      }
      if (-not $disk.OSDiskEphemeral) { $tcoStorageCost += $diskMonthlyCost }
    }
  } else {
    # Estimate from VM count if no storage data
    $tcoStorageCost = $vmCount_w * 7.68  # Assume Standard SSD
  }
  $tcoStorageCost = [math]::Round($tcoStorageCost, 0)
  
  # Profile storage estimate (FSLogix Azure Files)
  # Rough estimate: ~$0.06/GB/month for Azure Files Premium, 30 GB per user avg
  if ($hpProfileData.Count -gt 0) {
    $estProfileUsers = if ($assignedUsers.Count -gt 0) { $assignedUsers.Count } else { $vmCount_w }
    $tcoProfileCost = [math]::Round($estProfileUsers * 30 * 0.06, 0)  # 30 GB × $0.06/GB
  }
  
  # VNet / NSG overhead (small but real)
  $tcoNetworkCost = [math]::Round($vmCount_w * 1.50, 0)  # ~$1.50/VM for VNet, NSG, PIP overhead
  
  $tcoTotal = $tcoStorageCost + $tcoProfileCost + $tcoNetworkCost
  $avdFullTCO = if ($avdEffectiveMonthly) { $avdEffectiveMonthly + $tcoTotal } else { $null }
  $tcoCostDelta = if ($avdFullTCO -and $w365MonthlyTotal) { $w365MonthlyTotal - $avdFullTCO } else { $null }

  # === ENHANCEMENT 3: Entra ID / Intune readiness ===
  $joinType = "Unknown"
  $intuneReady = $false
  # Check VMs for AADLoginForWindows extension (definitive Entra ID join indicator)
  $aadExtCount = @($hpVms | Where-Object { $_.HasAadExtension -eq $true }).Count
  $identityCount = @($hpVms | Where-Object { $_.IdentityType }).Count
  if ($aadExtCount -gt 0) {
    $joinType = "Entra ID Joined"
    $intuneReady = $true
    if ($aadExtCount -lt $vmCount_w) {
      $joinType = "Mixed ($aadExtCount/$vmCount_w Entra ID)"
    }
  } elseif ($identityCount -gt 0) {
    # Has managed identity but no AAD extension — could be hybrid or AD-only with managed identity
    $joinType = "AD Joined (managed identity present)"
    $intuneReady = $false
  } else {
    $joinType = "AD Joined (no Entra extension)"
    $intuneReady = $false
  }

  # === ENHANCEMENT 4: Pilot suitability score ===
  $pilotScore = 0
  $pilotReasons = [System.Collections.Generic.List[string]]::new()
  if ($fitScore -ge 60) { $pilotScore += 30; $pilotReasons.Add("High W365 fit score ($fitScore/100)") }
  elseif ($fitScore -ge 45) { $pilotScore += 15; $pilotReasons.Add("Moderate W365 fit score ($fitScore/100)") }
  if ($isPersonal) { $pilotScore += 20; $pilotReasons.Add("Personal pool — simplest migration path") }
  if ($vmCount_w -le 15) { $pilotScore += 20; $pilotReasons.Add("Small pool ($vmCount_w VMs) — manageable pilot size") }
  elseif ($vmCount_w -le 30) { $pilotScore += 10; $pilotReasons.Add("Medium pool ($vmCount_w VMs)") }
  if ($migrationComplexity -eq "Low") { $pilotScore += 15; $pilotReasons.Add("Low migration complexity") }
  elseif ($migrationComplexity -eq "Medium") { $pilotScore += 5 }
  if ($intuneReady) { $pilotScore += 10; $pilotReasons.Add("Likely Entra ID joined — Intune-ready") }
  if (-not $hasGpu -and $fitBlockers.Count -eq 0) { $pilotScore += 5; $pilotReasons.Add("No blockers") }
  $pilotScore = [math]::Min(100, $pilotScore)

  $w365Analysis.Add([PSCustomObject]@{
    HostPoolName          = $hpName
    HostPoolType          = $hpType
    AppGroupType          = $appGroupType
    VMCount               = $vmCount_w
    UniqueUsers           = $poolUniqueUsers
    W365LicensesNeeded    = $w365LicenseCount
    DominantSku           = $vmSize
    vCPU                  = $vCPU
    RamGB                 = $ramGB
    ActiveSessions        = $totalActiveSessions
    AvgSessionsPerHost    = $avgSessionsPerHost
    HasScalingPlan        = $hasScalingPlan
    HasGPU                = $hasGpu
    FitScore              = $fitScore
    Recommendation        = $recommendation
    BestW365Plan          = if ($bestW365Plan) { $bestW365Plan.Name } else { "No suitable plan" }
    W365MonthlyPerUser    = if ($bestW365Plan) { $bestW365Monthly } else { $null }
    W365MonthlyTotal      = $w365MonthlyTotal
    AVDMonthlyTotal       = $avdMonthlyTotal
    AVDEffectiveMonthly   = $avdEffectiveMonthly
    CostDelta             = $costDelta
    CostVerdict           = $costVerdict
    CostSource            = $costSource
    Advantages            = ($fitAdvantages -join "; ")
    Considerations        = ($fitFactors -join "; ")
    Blockers              = ($fitBlockers -join "; ")
    MigrationComplexity   = $migrationComplexity
    MigrationFactors      = ($migrationFactors -join "; ")
    BreakevenUtilization  = $breakevenPct
    # New: Usage-based SKU
    UsageBasedPlan        = if ($usageBasedPlan) { "$($usageBasedPlan.vCPU)vCPU/$($usageBasedPlan.RAM)GB" } else { $null }
    UsageBasedMonthly     = $usageBasedMonthly
    UsageSavings          = $usageSavingsVsCurrent
    PoolAvgCPU            = if ($poolRsData.Count -gt 0) { $poolAvgCPU } else { $null }
    PoolPeakMem           = if ($poolRsData.Count -gt 0) { $poolPeakMem } else { $null }
    # New: TCO
    TCOStorageCost        = $tcoStorageCost
    TCOProfileCost        = $tcoProfileCost
    TCONetworkCost        = $tcoNetworkCost
    TCOTotal              = $tcoTotal
    AVDFullTCO            = $avdFullTCO
    TCOCostDelta          = $tcoCostDelta
    # New: Migration readiness
    JoinType              = $joinType
    IntuneReady           = $intuneReady
    # New: Pilot
    PilotScore            = $pilotScore
    PilotReasons          = ($pilotReasons -join "; ")
  })
}

# Summary
$w365Candidates = @($w365Analysis | Where-Object { $_.Recommendation -match "W365" })
$w365KeepAvd = @($w365Analysis | Where-Object { $_.Recommendation -eq "Keep AVD" })

Write-Host "  W365 Readiness: $($w365Candidates.Count) candidate(s), $($w365KeepAvd.Count) keep AVD" -ForegroundColor $(if ($w365Candidates.Count -gt 0) { "Yellow" } else { "Green" })

# Export
if ($w365Analysis.Count -gt 0) {
  $w365Analysis | Export-Csv (Join-Path $outFolder "ENHANCED-W365-Readiness.csv") -NoTypeInformation
}

# =========================================================
# Enhanced Analysis: Network Readiness (v4.0.0)
# =========================================================
Write-Host "Analyzing network readiness..." -ForegroundColor Cyan

$networkFindings = [System.Collections.Generic.List[object]]::new()

# --- 1. RDP Shortpath Analysis (from KQL data already collected) ---
$shortpathData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_WVDShortpathUsage" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "TransportType" })
$shortpathSummary = @{ TotalConnections = 0; UdpConnections = 0; TcpConnections = 0; ShortpathPct = 0 }
foreach ($sp in $shortpathData) {
  $count = [int]$sp.Connections
  $shortpathSummary.TotalConnections += $count
  $transport = "$($sp.TransportType)".Trim()
  if ($transport -match 'UDP|True|1|Shortpath') {
    $shortpathSummary.UdpConnections += $count
  } else {
    $shortpathSummary.TcpConnections += $count
  }
}
if ($shortpathSummary.TotalConnections -gt 0) {
  $shortpathSummary.ShortpathPct = [math]::Round(($shortpathSummary.UdpConnections / $shortpathSummary.TotalConnections) * 100, 0)
}
$shortpathEnabled = $shortpathSummary.ShortpathPct -gt 0

$networkFindings.Add([PSCustomObject]@{
  Check       = "RDP Shortpath"
  Status      = if ($shortpathSummary.ShortpathPct -ge 80) { "Good" } elseif ($shortpathEnabled) { "Partial" } else { "Not Enabled" }
  Detail      = "$($shortpathSummary.UdpConnections) of $($shortpathSummary.TotalConnections) connections using UDP ($($shortpathSummary.ShortpathPct)%)"
  Impact      = if (-not $shortpathEnabled) { "High" } elseif ($shortpathSummary.ShortpathPct -lt 50) { "Medium" } else { "Low" }
  Recommendation = if (-not $shortpathEnabled) { "Enable RDP Shortpath for managed/public networks — reduces latency by using direct UDP between client and host, bypassing the gateway relay" }
                   elseif ($shortpathSummary.ShortpathPct -lt 80) { "Some clients still using TCP relay — check client versions and network policies (UDP 3390 must be allowed)" }
                   else { "RDP Shortpath working well across most connections" }
})

Write-Host "  RDP Shortpath: $($shortpathSummary.ShortpathPct)% UDP" -ForegroundColor $(if ($shortpathSummary.ShortpathPct -ge 80) { "Green" } elseif ($shortpathEnabled) { "Yellow" } else { "Red" })

# --- 2. Subnet Analysis ---
# Collect unique subnets used by session hosts
$subnetAnalysis = [System.Collections.Generic.List[object]]::new()
$subnetVmCounts = @{}
$subnetHostPools = @{}
$uniqueSubnetIds = @($vms | Where-Object { $_.PSObject.Properties['SubnetId'] -and $_.SubnetId } | Select-Object -ExpandProperty SubnetId -Unique)

foreach ($vm_net in ($vms | Where-Object { $_.PSObject.Properties['SubnetId'] -and $_.SubnetId })) {
  $sid = $vm_net.SubnetId
  $subnetVmCounts[$sid] = ($subnetVmCounts[$sid] ?? 0) + 1
  if (-not $subnetHostPools[$sid]) { $subnetHostPools[$sid] = [System.Collections.Generic.List[string]]::new() }
  if ($vm_net.HostPoolName -notin $subnetHostPools[$sid]) { $subnetHostPools[$sid].Add($vm_net.HostPoolName) }
}

# Fetch subnet details (batched by VNet — typically 1-2 VNets)
$vnetCache = @{}
foreach ($subnetId in $uniqueSubnetIds) {
  try {
    # Parse VNet info from subnet ID: .../virtualNetworks/{vnetName}/subnets/{subnetName}
    $parts = $subnetId -split '/'
    $vnetRg = $parts[4]
    $vnetName = $parts[8]
    $subnetName = $parts[10]
    $vnetKey = "$vnetRg/$vnetName"
    
    if (-not $vnetCache.ContainsKey($vnetKey)) {
      $vnet = Get-AzVirtualNetwork -ResourceGroupName $vnetRg -Name $vnetName -ErrorAction SilentlyContinue
      $vnetCache[$vnetKey] = $vnet
    }
    $vnet = $vnetCache[$vnetKey]
    if (-not $vnet) { continue }
    
    $subnet = $vnet.Subnets | Where-Object { $_.Name -eq $subnetName } | Select-Object -First 1
    if (-not $subnet) { continue }
    
    # Calculate subnet capacity
    $prefixParts = $subnet.AddressPrefix -split '/'
    $cidr = [int]$prefixParts[1]
    $totalIPs = [math]::Pow(2, 32 - $cidr)
    $usableIPs = $totalIPs - 5  # Azure reserves 5 IPs per subnet
    $usedIPs = ($subnet.IpConfigurations | Measure-Object).Count
    $availableIPs = [math]::Max(0, $usableIPs - $usedIPs)
    $usagePct = if ($usableIPs -gt 0) { [math]::Round(($usedIPs / $usableIPs) * 100, 0) } else { 100 }
    
    # Check for subnet NSG
    $subnetNsg = if ($subnet.PSObject.Properties['NetworkSecurityGroup'] -and $subnet.NetworkSecurityGroup) { $subnet.NetworkSecurityGroup.Id } else { $null }
    
    # Check for subnet route table (UDR)
    $subnetRouteTable = if ($subnet.PSObject.Properties['RouteTable'] -and $subnet.RouteTable) { $subnet.RouteTable.Id } else { $null }
    
    $subnetAnalysis.Add([PSCustomObject]@{
      SubnetId      = $subnetId
      SubnetName    = $subnetName
      VNetName      = $vnetName
      ResourceGroup = $vnetRg
      AddressPrefix = $subnet.AddressPrefix
      CIDR          = $cidr
      TotalIPs      = [int]$totalIPs
      UsableIPs     = [int]$usableIPs
      UsedIPs       = $usedIPs
      AvailableIPs  = [int]$availableIPs
      UsagePct      = $usagePct
      SessionHostVMs = $subnetVmCounts[$subnetId] ?? 0
      HostPools     = ($subnetHostPools[$subnetId] -join ", ")
      HasNSG        = [bool]$subnetNsg
      NsgId         = $subnetNsg
      HasRouteTable = [bool]$subnetRouteTable
      RouteTableId  = $subnetRouteTable
    })
    
    # Findings
    if ($usagePct -ge 90) {
      $networkFindings.Add([PSCustomObject]@{
        Check          = "Subnet Capacity"
        Status         = "Critical"
        Detail         = "$subnetName ($($subnet.AddressPrefix)) is $usagePct% full — only $availableIPs IPs remaining for $($subnetVmCounts[$subnetId]) session hosts"
        Impact         = "High"
        Recommendation = "Expand subnet or migrate session hosts to a larger subnet before scaling. A /$([math]::Max($cidr - 2, 16)) would provide $([int]([math]::Pow(2, 32 - [math]::Max($cidr - 2, 16))) - 5) usable IPs."
      })
    }
    elseif ($usagePct -ge 70) {
      $networkFindings.Add([PSCustomObject]@{
        Check          = "Subnet Capacity"
        Status         = "Warning"
        Detail         = "$subnetName ($($subnet.AddressPrefix)) is $usagePct% used — $availableIPs IPs remaining"
        Impact         = "Medium"
        Recommendation = "Plan subnet expansion before next scale-out. Current capacity may not support doubling the fleet."
      })
    }
    
    if (-not $subnetNsg) {
      $networkFindings.Add([PSCustomObject]@{
        Check          = "Subnet NSG"
        Status         = "Missing"
        Detail         = "Subnet $subnetName ($vnetName) has no Network Security Group attached"
        Impact         = "Medium"
        Recommendation = "Attach an NSG to enforce network segmentation. AVD requires outbound HTTPS (443) to *.wvd.microsoft.com and related services."
      })
    }
  } catch {
    # Subnet lookup failed — not critical
  }
}

Write-Host "  Subnets: $($subnetAnalysis.Count) analyzed" -ForegroundColor Gray

# --- 3. VNet DNS and Peering Analysis ---
$vnetAnalysis = [System.Collections.Generic.List[object]]::new()
foreach ($vnetEntry in $vnetCache.GetEnumerator()) {
  $vnet = $vnetEntry.Value
  if (-not $vnet) { continue }
  
  $dnsServers = if ($vnet.PSObject.Properties['DhcpOptions'] -and $vnet.DhcpOptions.PSObject.Properties['DnsServers']) { $vnet.DhcpOptions.DnsServers } else { $null }
  $hasCustomDns = ($dnsServers -and @($dnsServers).Count -gt 0)
  $hasAzureDns = (-not $hasCustomDns)
  
  # Check peerings
  $peerings = if ($vnet.PSObject.Properties['VirtualNetworkPeerings']) { $vnet.VirtualNetworkPeerings } else { @() }
  $peeringCount = @($peerings).Count
  $disconnectedPeerings = @($peerings | Where-Object { $_.PeeringState -ne "Connected" })
  
  $vnetAnalysis.Add([PSCustomObject]@{
    VNetName          = $vnet.Name
    ResourceGroup     = $vnet.ResourceGroupName
    Location          = $vnet.Location
    AddressSpace      = ($vnet.AddressSpace.AddressPrefixes -join ", ")
    DnsType           = if ($hasCustomDns) { "Custom" } else { "Azure Default" }
    DnsServers        = if ($hasCustomDns) { ($dnsServers -join ", ") } else { "Azure-provided" }
    PeeringCount      = $peeringCount
    DisconnectedPeers = $disconnectedPeerings.Count
    PeerDetails       = ($peerings | ForEach-Object { "$($_.Name): $($_.PeeringState)" }) -join "; "
  })
  
  # DNS findings
  if ($hasCustomDns) {
    $networkFindings.Add([PSCustomObject]@{
      Check          = "DNS Configuration"
      Status         = "Custom DNS"
      Detail         = "$($vnet.Name) uses custom DNS servers: $($dnsServers -join ', '). Ensure these can resolve Azure Private DNS zones and *.file.core.windows.net for FSLogix."
      Impact         = "Low"
      Recommendation = "Verify DNS can resolve storage account private endpoints and AD domain controller(s). Consider adding Azure DNS (168.63.129.16) as a fallback."
    })
  } else {
    $networkFindings.Add([PSCustomObject]@{
      Check          = "DNS Configuration"
      Status         = "Azure Default"
      Detail         = "$($vnet.Name) uses Azure-provided DNS. Custom DNS is typically required for domain join and FSLogix profile storage."
      Impact         = "Medium"
      Recommendation = "If using Azure AD DS or on-prem AD, configure custom DNS pointing to domain controllers. Azure DNS alone may not resolve domain-specific resources."
    })
  }
  
  # Peering findings
  if ($disconnectedPeerings.Count -gt 0) {
    $disconnectedNames = ($disconnectedPeerings | ForEach-Object { "$($_.Name) ($($_.PeeringState))" }) -join ", "
    $networkFindings.Add([PSCustomObject]@{
      Check          = "VNet Peering"
      Status         = "Disconnected"
      Detail         = "$($disconnectedPeerings.Count) peering(s) not in Connected state on $($vnet.Name): $disconnectedNames"
      Impact         = "High"
      Recommendation = "Reconnect peerings — disconnected peerings can break FSLogix profile access, domain authentication, and application connectivity."
    })
  }
}

Write-Host "  VNets: $($vnetAnalysis.Count) analyzed, DNS: $(($vnetAnalysis | Where-Object { $_.DnsType -eq 'Custom' } | Measure-Object).Count) custom / $(($vnetAnalysis | Where-Object { $_.DnsType -eq 'Azure Default' } | Measure-Object).Count) default" -ForegroundColor Gray

# --- 4. Private Endpoint Check for Host Pools ---
$privateEndpointFindings = @()
foreach ($hpItem in $hostPools) {
  $hpId = $hpItem.Id
  if (-not $hpId) { continue }
  try {
    $hpRg = $hpItem.ResourceGroup
    $peConnections = Get-AzPrivateEndpointConnection -PrivateLinkResourceId $hpId -ErrorAction SilentlyContinue
    $hasPE = ($peConnections | Measure-Object).Count -gt 0
    $privateEndpointFindings += [PSCustomObject]@{
      HostPoolName    = $hpItem.HostPoolName
      HasPrivateEndpoint = $hasPE
      EndpointCount   = ($peConnections | Measure-Object).Count
      Status          = if ($hasPE) { ($peConnections | ForEach-Object { $_.PrivateLinkServiceConnectionState.Status }) -join ", " } else { "None" }
    }
  } catch {
    # Private endpoint check not available — might not have permissions
    $privateEndpointFindings += [PSCustomObject]@{
      HostPoolName    = $hpItem.HostPoolName
      HasPrivateEndpoint = $null
      EndpointCount   = 0
      Status          = "Unknown (check permissions)"
    }
  }
}

$hpWithoutPE = @($privateEndpointFindings | Where-Object { $_.HasPrivateEndpoint -eq $false })
if ($hpWithoutPE.Count -gt 0) {
  $networkFindings.Add([PSCustomObject]@{
    Check          = "Private Endpoints"
    Status         = "Not Configured"
    Detail         = "$($hpWithoutPE.Count) host pool(s) without private endpoints: $(($hpWithoutPE.HostPoolName) -join ', ')"
    Impact         = "Medium"
    Recommendation = "Configure private endpoints for host pools to keep control plane traffic on the Microsoft backbone. This improves security and can reduce connection establishment latency."
  })
}

# --- 5. NIC-level NSG Check ---
$vmsWithoutNsg = @($vms | Where-Object { $_.PSObject.Properties['NsgId'] -and (-not $_.NsgId) -and $_.PSObject.Properties['SubnetId'] -and $_.SubnetId })
$subnetNsgCoverage = @($subnetAnalysis | Where-Object { $_.HasNSG })
$vmsWithNoNsgAtAll = @()
foreach ($vm_nsg in $vmsWithoutNsg) {
  $subMatch = $subnetAnalysis | Where-Object { $_.SubnetId -eq $vm_nsg.SubnetId -and $_.HasNSG } | Select-Object -First 1
  if (-not $subMatch) {
    $vmsWithNoNsgAtAll += $vm_nsg
  }
}
if ($vmsWithNoNsgAtAll.Count -gt 0) {
  $networkFindings.Add([PSCustomObject]@{
    Check          = "NSG Coverage"
    Status         = "Unprotected"
    Detail         = "$($vmsWithNoNsgAtAll.Count) session host(s) have no NSG at NIC or subnet level"
    Impact         = "High"
    Recommendation = "Attach an NSG at the subnet level (preferred) or NIC level. AVD session hosts should have NSGs restricting inbound traffic."
  })
}

# Summary
$criticalNetFindings = @($networkFindings | Where-Object { $_.Impact -eq "High" })
$warningNetFindings = @($networkFindings | Where-Object { $_.Impact -eq "Medium" })
Write-Host "  Network Readiness: $($criticalNetFindings.Count) critical, $($warningNetFindings.Count) warnings, $($networkFindings.Count) total checks" -ForegroundColor $(if ($criticalNetFindings.Count -gt 0) { "Red" } elseif ($warningNetFindings.Count -gt 0) { "Yellow" } else { "Green" })

# =========================================================
# Enhanced Analysis: Security Posture (v4.0.0)
# =========================================================
Write-Host "Analyzing security posture..." -ForegroundColor Cyan

$securityPosture = [System.Collections.Generic.List[object]]::new()

foreach ($hp in $hostPools) {
  $hpName = $hp.HostPoolName
  if (-not $hpName) { continue }
  $hpVms = @($vms | Where-Object { $_.HostPoolName -eq $hpName })
  $vmCt = $hpVms.Count
  if ($vmCt -eq 0) { continue }
  
  $trustedLaunch = @($hpVms | Where-Object { $_.SecurityType -match 'TrustedLaunch' }).Count
  $secureBoot = @($hpVms | Where-Object { $_.SecureBoot -eq $true }).Count
  $vtpm = @($hpVms | Where-Object { $_.VTpm -eq $true }).Count
  $hostEnc = @($hpVms | Where-Object { $_.HostEncryption -eq $true }).Count
  $ephemeralDisk = @($hpVms | Where-Object { $_.OSDiskEphemeral -eq $true }).Count
  $accelNet = @($hpVms | Where-Object { $_.AccelNetEnabled -eq $true }).Count
  
  # Check for NSG coverage from network findings
  $hpNsgMissing = @($hpVms | Where-Object { -not $_.NsgId -and -not $_.SubnetId }).Count
  
  # Check private endpoint (from network findings if available)
  $hasPrivateEndpoint = @($networkFindings | Where-Object { $_.Check -eq "Private Endpoints" -and $_.Status -eq "Good" }).Count -gt 0
  
  # Score (0-100)
  $secScore = 0
  $secFindings = [System.Collections.Generic.List[string]]::new()
  
  # Trusted Launch (25 points)
  $tlPct = [math]::Round(($trustedLaunch / $vmCt) * 100, 0)
  if ($tlPct -eq 100) { $secScore += 25 }
  elseif ($tlPct -ge 50) { $secScore += 15; $secFindings.Add("$($vmCt - $trustedLaunch) VM(s) without Trusted Launch") }
  else { $secFindings.Add("$($vmCt - $trustedLaunch)/$vmCt VM(s) without Trusted Launch — strongly recommended for new deployments") }
  
  # Secure Boot (20 points)
  $sbPct = [math]::Round(($secureBoot / $vmCt) * 100, 0)
  if ($sbPct -eq 100) { $secScore += 20 }
  elseif ($sbPct -ge 50) { $secScore += 10; $secFindings.Add("$($vmCt - $secureBoot) VM(s) without Secure Boot") }
  else { $secFindings.Add("$($vmCt - $secureBoot)/$vmCt VM(s) without Secure Boot") }
  
  # vTPM (15 points)
  $vtpmPct = [math]::Round(($vtpm / $vmCt) * 100, 0)
  if ($vtpmPct -eq 100) { $secScore += 15 }
  elseif ($vtpmPct -ge 50) { $secScore += 8; $secFindings.Add("$($vmCt - $vtpm) VM(s) without vTPM") }
  else { $secFindings.Add("$($vmCt - $vtpm)/$vmCt VM(s) without vTPM") }
  
  # Host Encryption (15 points)
  $hePct = [math]::Round(($hostEnc / $vmCt) * 100, 0)
  if ($hePct -eq 100) { $secScore += 15 }
  elseif ($hePct -ge 50) { $secScore += 8; $secFindings.Add("$($vmCt - $hostEnc) VM(s) without host-based encryption") }
  else { $secFindings.Add("$($vmCt - $hostEnc)/$vmCt VM(s) without host-based encryption") }
  
  # Accelerated Networking (10 points)
  $anPct = [math]::Round(($accelNet / $vmCt) * 100, 0)
  if ($anPct -eq 100) { $secScore += 10 }
  elseif ($anPct -ge 50) { $secScore += 5 }
  else { $secFindings.Add("$($vmCt - $accelNet)/$vmCt VM(s) without Accelerated Networking") }
  
  # Ephemeral OS Disk (10 points - reduces attack surface, faster reimage)
  $edPct = [math]::Round(($ephemeralDisk / $vmCt) * 100, 0)
  if ($edPct -eq 100) { $secScore += 10 }
  elseif ($edPct -ge 50) { $secScore += 5 }
  
  # Private Endpoints (5 points)
  if ($hasPrivateEndpoint) { $secScore += 5 }
  else { $secFindings.Add("No private endpoints for control plane traffic") }
  
  $secGrade = if ($secScore -ge 90) { "A" } elseif ($secScore -ge 75) { "B" } elseif ($secScore -ge 60) { "C" } elseif ($secScore -ge 40) { "D" } else { "F" }
  
  $securityPosture.Add([PSCustomObject]@{
    HostPoolName        = $hpName
    VMCount             = $vmCt
    TrustedLaunchPct    = $tlPct
    SecureBootPct       = $sbPct
    VTpmPct             = $vtpmPct
    HostEncryptionPct   = $hePct
    AccelNetPct         = $anPct
    EphemeralDiskPct    = $edPct
    SecurityScore       = $secScore
    SecurityGrade       = $secGrade
    Findings            = ($secFindings -join "; ")
  })
}

$lowSecPools = @($securityPosture | Where-Object { $_.SecurityScore -lt 60 })
Write-Host "  Security Posture: $($lowSecPools.Count) pool(s) below grade C, $(($securityPosture | Measure-Object).Count) analyzed" -ForegroundColor $(if ($lowSecPools.Count -gt 0) { "Yellow" } else { "Green" })

if ($securityPosture.Count -gt 0) {
  $securityPosture | Export-Csv (Join-Path $outFolder "ENHANCED-Security-Posture.csv") -NoTypeInformation
}

# =========================================================
# Enhanced Analysis: Orphaned Resources (v4.0.0)
# =========================================================
Write-Host "Scanning for orphaned resources..." -ForegroundColor Cyan

$orphanedResources = [System.Collections.Generic.List[object]]::new()

# Build set of resource groups that contain AVD resources (scope orphan scan to AVD RGs only)
$avdResourceGroups = @{}
foreach ($v in $vms) {
  if ($v.ResourceGroup) { $avdResourceGroups["$($v.SubscriptionId)|$($v.ResourceGroup)".ToLower()] = $true }
}
foreach ($hp in $hostPools) {
  if ($hp.ResourceGroup) {
    $hpSubId = if ($hp.PSObject.Properties.Name -contains 'SubscriptionId') { $hp.SubscriptionId } else { $SubscriptionIds[0] }
    $avdResourceGroups["$hpSubId|$($hp.ResourceGroup)".ToLower()] = $true
  }
}

foreach ($subId in $SubscriptionIds) {
  # Save and restore context to avoid side effects
  $previousContext = Get-AzContext
  try { Set-AzContext -SubscriptionId $subId -ErrorAction Stop | Out-Null } catch { continue }
  
  # Get AVD RGs for this subscription
  $subAvdRgs = @($avdResourceGroups.Keys | Where-Object { $_.StartsWith("$subId|".ToLower()) } | ForEach-Object { ($_ -split '\|', 2)[1] })
  if ($subAvdRgs.Count -eq 0) { continue }
  
  foreach ($rgName in $subAvdRgs) {
    # Orphaned disks — unattached managed disks in AVD resource groups
    try {
      $rgDisks = Get-AzDisk -ResourceGroupName $rgName -ErrorAction SilentlyContinue
      foreach ($disk in SafeArray $rgDisks) {
        if ($disk.DiskState -eq "Unattached" -and $null -eq $disk.ManagedBy) {
          $diskMonthlyCost = switch ($disk.Sku.Name) {
            "Premium_LRS"    { if ($disk.DiskSizeGB -le 32) { 5.28 } elseif ($disk.DiskSizeGB -le 64) { 10.21 } elseif ($disk.DiskSizeGB -le 128) { 19.71 } elseif ($disk.DiskSizeGB -le 256) { 38.02 } elseif ($disk.DiskSizeGB -le 512) { 73.22 } else { 135.17 } }
            "StandardSSD_LRS" { if ($disk.DiskSizeGB -le 32) { 1.54 } elseif ($disk.DiskSizeGB -le 64) { 3.07 } elseif ($disk.DiskSizeGB -le 128) { 6.14 } elseif ($disk.DiskSizeGB -le 256) { 12.29 } elseif ($disk.DiskSizeGB -le 512) { 24.58 } else { 49.15 } }
            "Standard_LRS"    { if ($disk.DiskSizeGB -le 32) { 1.54 } elseif ($disk.DiskSizeGB -le 64) { 2.56 } elseif ($disk.DiskSizeGB -le 128) { 5.12 } else { 10.24 } }
            default { 5.00 }
          }
          $orphanedResources.Add([PSCustomObject]@{
            SubscriptionId = $subId
            ResourceType   = "Disk"
            ResourceName   = $disk.Name
            ResourceGroup  = $disk.ResourceGroupName
            Details        = "$($disk.Sku.Name), $($disk.DiskSizeGB) GB, State: $($disk.DiskState)"
            EstMonthlyCost = [math]::Round($diskMonthlyCost, 2)
            CreatedDate    = $disk.TimeCreated
          })
        }
      }
    } catch {}
    
    # Orphaned NICs — not attached to any VM, in AVD resource groups
    try {
      $rgNics = Get-AzNetworkInterface -ResourceGroupName $rgName -ErrorAction SilentlyContinue
      foreach ($nic in SafeArray $rgNics) {
        if (-not $nic.VirtualMachine -and -not $nic.PrivateEndpoint) {
          $orphanedResources.Add([PSCustomObject]@{
            SubscriptionId = $subId
            ResourceType   = "NIC"
            ResourceName   = $nic.Name
            ResourceGroup  = $nic.ResourceGroupName
            Details        = "IP: $(if ($nic.IpConfigurations) { $nic.IpConfigurations[0].PrivateIpAddress } else { 'None' }), NSG: $(if ($nic.NetworkSecurityGroup) { 'Yes' } else { 'None' })"
            EstMonthlyCost = 0
            CreatedDate    = $null
          })
        }
      }
    } catch {}
    
    # Orphaned Public IPs — not associated, in AVD resource groups
    try {
      $rgPips = Get-AzPublicIpAddress -ResourceGroupName $rgName -ErrorAction SilentlyContinue
      foreach ($pip in SafeArray $rgPips) {
        if (-not $pip.IpConfiguration) {
          $pipCost = if ($pip.Sku.Name -eq "Standard") { 3.65 } else { 2.63 }
          $orphanedResources.Add([PSCustomObject]@{
            SubscriptionId = $subId
            ResourceType   = "PublicIP"
            ResourceName   = $pip.Name
            ResourceGroup  = $pip.ResourceGroupName
            Details        = "SKU: $($pip.Sku.Name), IP: $($pip.IpAddress), Allocation: $($pip.PublicIpAllocationMethod)"
            EstMonthlyCost = $pipCost
            CreatedDate    = $null
          })
        }
      }
    } catch {}
  }
  
  # Restore previous context
  if ($previousContext) {
    try { Set-AzContext -Context $previousContext -ErrorAction SilentlyContinue | Out-Null } catch {}
  }
}

$orphanedWaste = ($orphanedResources | ForEach-Object { $_.EstMonthlyCost } | Measure-Object -Sum).Sum
$orphanedWaste = if ($orphanedWaste) { [math]::Round($orphanedWaste, 0) } else { 0 }
Write-Host "  Orphaned Resources: $($orphanedResources.Count) found, ~`$$orphanedWaste/mo wasted" -ForegroundColor $(if ($orphanedResources.Count -gt 0) { "Yellow" } else { "Green" })

if ($orphanedResources.Count -gt 0) {
  $orphanedResources | Export-Csv (Join-Path $outFolder "ENHANCED-Orphaned-Resources.csv") -NoTypeInformation
}

# =========================================================
# Enhanced Analysis: Profile Health (v4.0.0)
# =========================================================
Write-Host "Analyzing profile load performance..." -ForegroundColor Cyan

$profileHealth = [System.Collections.Generic.List[object]]::new()
$profileLoadData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ProfileLoadPerformance" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "SessionHostName" })

if ($profileLoadData.Count -gt 0) {
  foreach ($pl in $profileLoadData) {
    $p95 = if ($pl.PSObject.Properties.Name -contains "P95ProfileLoadSec") { [double]$pl.P95ProfileLoadSec } else { 0 }
    $avg = if ($pl.PSObject.Properties.Name -contains "AvgProfileLoadSec") { [double]$pl.AvgProfileLoadSec } else { 0 }
    $max = if ($pl.PSObject.Properties.Name -contains "MaxProfileLoadSec") { [double]$pl.MaxProfileLoadSec } else { 0 }
    $count = if ($pl.PSObject.Properties.Name -contains "ConnectionCount") { [int]$pl.ConnectionCount } else { 0 }
    $hostName = "$($pl.SessionHostName)"
    
    $severity = if ($p95 -ge 60) { "Critical" } elseif ($p95 -ge 30) { "Warning" } else { "Good" }
    
    $profileHealth.Add([PSCustomObject]@{
      SessionHostName = $hostName
      HostPoolName    = if ($pl.PSObject.Properties.Name -contains "HostPool") { $pl.HostPool }
                        else {
                          # Fallback: match against session hosts, then VMs
                          $shortName = ($hostName -split '\.')[0]
                          $shMatch = $sessionHosts | Where-Object { $_.SessionHostName -like "*$shortName*" } | Select-Object -First 1
                          if ($shMatch) { $shMatch.HostPoolName }
                          else {
                            $vmMatch = $vms | Where-Object { $_.VMName -eq $shortName -or $_.SessionHostName -like "*$shortName*" } | Select-Object -First 1
                            if ($vmMatch) { $vmMatch.HostPoolName } else { "Unknown" }
                          }
                        }
      ConnectionCount = $count
      AvgLoadSec      = $avg
      P95LoadSec      = $p95
      MaxLoadSec      = $max
      Severity        = $severity
    })
  }
}

$slowProfiles = @($profileHealth | Where-Object { $_.Severity -ne "Good" })
Write-Host "  Profile Health: $($slowProfiles.Count) host(s) with slow profile loads (P95 >= 30s)" -ForegroundColor $(if (@($profileHealth | Where-Object { $_.Severity -eq "Critical" }).Count -gt 0) { "Red" } elseif ($slowProfiles.Count -gt 0) { "Yellow" } else { "Green" })

if ($profileHealth.Count -gt 0) {
  $profileHealth | Export-Csv (Join-Path $outFolder "ENHANCED-Profile-Health.csv") -NoTypeInformation
}

# =========================================================
# Enhanced Analysis: User Experience Score (v4.0.0)
# =========================================================
Write-Host "Calculating user experience scores..." -ForegroundColor Cyan

$uxScores = [System.Collections.Generic.List[object]]::new()

# Gather per-host-pool signals from existing data
foreach ($hp in $hostPools) {
  $hpName = $hp.HostPoolName
  if (-not $hpName) { continue }
  
  # 1. Profile load score (0-25): P95 across pool
  $poolProfiles = @($profileHealth | Where-Object { $_.HostPoolName -eq $hpName })
  $poolP95 = if ($poolProfiles.Count -gt 0) { ($poolProfiles | ForEach-Object { $_.P95LoadSec } | Measure-Object -Maximum).Maximum } else { $null }
  $profileScore = if (-not $poolP95) { $null }
                  elseif ($poolP95 -lt 5) { 25 }
                  elseif ($poolP95 -lt 10) { 22 }
                  elseif ($poolP95 -lt 20) { 18 }
                  elseif ($poolP95 -lt 30) { 12 }
                  elseif ($poolP95 -lt 60) { 6 }
                  else { 0 }
  
  # 2. Connection quality score (0-25): RTT from KQL
  $connQuality = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionQuality" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "HostPool" -and $_.HostPool -eq $hpName })
  $avgRtt = if ($connQuality.Count -gt 0 -and $connQuality[0].PSObject.Properties.Name -contains "AvgRTTms") { [double]$connQuality[0].AvgRTTms } else { $null }
  $rttScore = if (-not $avgRtt) { $null }
              elseif ($avgRtt -lt 50) { 25 }
              elseif ($avgRtt -lt 100) { 22 }
              elseif ($avgRtt -lt 150) { 18 }
              elseif ($avgRtt -lt 200) { 12 }
              elseif ($avgRtt -lt 300) { 6 }
              else { 0 }
  
  # 3. Disconnect rate score (0-25): abnormal disconnects
  $poolDisconnects = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_DisconnectReasons" -and $_.QueryName -eq "AVD" })
  $totalSessions = ($poolDisconnects | ForEach-Object { if ($_.PSObject.Properties.Name -contains "SessionCount") { [int]$_.SessionCount } else { 0 } } | Measure-Object -Sum).Sum
  # Use w365Analysis or disconnect data to get abnormal rate
  $poolDcByHost = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_DisconnectByHost" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "HostPool" -and $_.HostPool -eq $hpName })
  $abnormalPct = if ($poolDcByHost.Count -gt 0 -and $poolDcByHost[0].PSObject.Properties.Name -contains "AbnormalPct") { [double]$poolDcByHost[0].AbnormalPct } else { $null }
  $dcScore = if (-not $abnormalPct) { $null }
             elseif ($abnormalPct -lt 5) { 25 }
             elseif ($abnormalPct -lt 10) { 20 }
             elseif ($abnormalPct -lt 20) { 15 }
             elseif ($abnormalPct -lt 30) { 8 }
             else { 0 }
  
  # 4. Connection error rate score (0-25)
  $connErrors = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionErrors" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "HostPool" -and $_.HostPool -eq $hpName })
  $errorCount = if ($connErrors.Count -gt 0 -and $connErrors[0].PSObject.Properties.Name -contains "ErrorCount") { [int]$connErrors[0].ErrorCount } else { $null }
  $errorScore = if (-not $errorCount) { $null }
                elseif ($errorCount -eq 0) { 25 }
                elseif ($errorCount -lt 5) { 20 }
                elseif ($errorCount -lt 20) { 15 }
                elseif ($errorCount -lt 50) { 8 }
                else { 0 }
  
  # Composite score
  $componentScores = @(@($profileScore, $rttScore, $dcScore, $errorScore) | Where-Object { $null -ne $_ })
  if ($componentScores.Count -eq 0) { continue }
  
  # Scale to 100 based on available components
  $maxPossible = $componentScores.Count * 25
  $rawTotal = ($componentScores | Measure-Object -Sum).Sum
  $uxScore = [math]::Round(($rawTotal / $maxPossible) * 100, 0)
  $uxGrade = if ($uxScore -ge 90) { "A" } elseif ($uxScore -ge 75) { "B" } elseif ($uxScore -ge 60) { "C" } elseif ($uxScore -ge 40) { "D" } else { "F" }
  
  $uxScores.Add([PSCustomObject]@{
    HostPoolName       = $hpName
    UXScore            = $uxScore
    UXGrade            = $uxGrade
    ProfileLoadScore   = $profileScore
    ProfileP95Sec      = $poolP95
    ConnectionScore    = $rttScore
    AvgRTTms           = $avgRtt
    DisconnectScore    = $dcScore
    AbnormalDcPct      = $abnormalPct
    ErrorScore         = $errorScore
    ErrorCount         = $errorCount
    ComponentsAvailable = $componentScores.Count
  })
}

$poorUx = @($uxScores | Where-Object { $_.UXScore -lt 60 })
Write-Host "  UX Scores: $($poorUx.Count) pool(s) below grade C, $(($uxScores | Measure-Object).Count) scored" -ForegroundColor $(if ($poorUx.Count -gt 0) { "Red" } elseif (@($uxScores | Where-Object { $_.UXScore -lt 75 }).Count -gt 0) { "Yellow" } else { "Green" })

if ($uxScores.Count -gt 0) {
  $uxScores | Export-Csv (Join-Path $outFolder "ENHANCED-UX-Scores.csv") -NoTypeInformation
}

# =========================================================
# Enhanced Analysis: Connection Success Rate per Pool (v4.1)
# =========================================================
$connSuccessData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionSuccessRate" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "HostPool" })
if ($connSuccessData.Count -gt 0) {
  Write-Host "  Connection Success: $($connSuccessData.Count) pool(s) analyzed" -ForegroundColor $(if (@($connSuccessData | Where-Object { [double]$_.FailureRate -gt 10 }).Count -gt 0) { "Red" } else { "Green" })
}

# =========================================================
# Enhanced Analysis: Login Time per Pool (v4.1)
# =========================================================
$loginTimeData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_LoginTime" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "HostPool" })
if ($loginTimeData.Count -gt 0) {
  $slowLogins = @($loginTimeData | Where-Object { [double]$_.P95LoginSec -gt 60 })
  Write-Host "  Login Times: $($loginTimeData.Count) pool(s), $($slowLogins.Count) with P95 > 60s" -ForegroundColor $(if ($slowLogins.Count -gt 0) { "Yellow" } else { "Green" })
}

# =========================================================
# Enhanced Analysis: Drain Mode Detection (v4.1)
# =========================================================
$drainModeHosts = @($sessionHosts | Where-Object { $_.AllowNewSession -eq $false })
$drainByPool = @{}
foreach ($dh in $drainModeHosts) {
  $pool = $dh.HostPoolName
  if (-not $drainByPool.ContainsKey($pool)) { $drainByPool[$pool] = 0 }
  $drainByPool[$pool]++
}
if ($drainModeHosts.Count -gt 0) {
  Write-Host "  Drain Mode: $($drainModeHosts.Count) host(s) in drain across $($drainByPool.Count) pool(s)" -ForegroundColor Yellow
}

# =========================================================
# Enhanced Analysis: Per-User Cost (v4.1)
# =========================================================
$perUserCost = [System.Collections.Generic.List[object]]::new()
# Calculated after cost analysis, so placeholder populated later

# =========================================================
# Enhanced Analysis: W365 Feature Gap (v4.1)
# =========================================================
# Built from research: features available in AVD that W365 lacks or handles differently
# This is checked per pool based on detected workload characteristics
$w365FeatureGaps = @{
  "MultiSession" = @{
    Feature = "Multi-session hosting"
    AVD = "Supports Windows 10/11 Enterprise multi-session — multiple users share one VM"
    W365 = "Single-user only (1 Cloud PC per user). No multi-session support."
    Impact = "High"
    Detection = "Pooled pool with MaxSessions > 2"
  }
  "RemoteApp" = @{
    Feature = "Published RemoteApp"
    AVD = "Native RemoteApp publishing — deliver individual apps without full desktop"
    W365 = "W365 Cloud Apps (Frontline Shared) can deliver apps but requires Frontline licensing. No native RemoteApp equivalent in Enterprise plans."
    Impact = "Medium"
    Detection = "RemoteApp pool type"
  }
  "GPU" = @{
    Feature = "GPU-accelerated workloads"
    AVD = "Full GPU support (NC, NV, ND series) for CAD, 3D, video rendering, ML/AI"
    W365 = "GPU Cloud PCs available but limited SKUs (max 16 vCPU). No NV/NC/ND-equivalent customization."
    Impact = "High"
    Detection = "GPU VM SKUs detected"
  }
  "AppAttach" = @{
    Feature = "App Attach / MSIX app delivery"
    AVD = "App Attach dynamically mounts apps to sessions — simplifies image management"
    W365 = "Not available. Apps must be installed in the image or deployed via Intune Win32 apps."
    Impact = "Medium"
    Detection = "Always flag for awareness"
  }
  "Autoscale" = @{
    Feature = "Autoscale / elastic capacity"
    AVD = "Native autoscale with ramp-up, peak, ramp-down, off-peak schedules. Pay only for running VMs."
    W365 = "Fixed cost per user per month. No elastic scaling — you pay whether the Cloud PC is used or not."
    Impact = "High"
    Detection = "Has scaling plan attached"
  }
  "PrivateLink" = @{
    Feature = "Azure Private Link"
    AVD = "Supports Private Link for session host traffic and feed discovery"
    W365 = "Enterprise supports Azure Network Connection for hybrid access, but no Private Link equivalent for Cloud PC traffic."
    Impact = "Medium"
    Detection = "VNet with private endpoints detected"
  }
  "CustomNetworking" = @{
    Feature = "Custom VNet / NSG / UDR control"
    AVD = "Full Azure networking — custom VNets, NSGs, route tables, firewall integration"
    W365 = "Enterprise supports Azure Network Connection, but networking is simplified. Business edition has no network control."
    Impact = "Medium"
    Detection = "Custom NSGs or UDRs detected"
  }
  "TeamsOptimization" = @{
    Feature = "Microsoft Teams media optimization"
    AVD = "Supported — WebRTC redirector offloads audio/video to local client"
    W365 = "Supported — same WebRTC redirector mechanism. Requires manual update of WebRTC Redirector Service."
    Impact = "Low"
    Detection = "Informational — both platforms support this"
  }
  "MultiMonitor" = @{
    Feature = "Multi-monitor support"
    AVD = "Up to 16 monitors at 8K resolution (RDP 10)"
    W365 = "Supported via Windows App. Up to 4 monitors in most configurations."
    Impact = "Low"
    Detection = "Informational"
  }
  "MMR" = @{
    Feature = "Multimedia redirection (MMR)"
    AVD = "Supported — redirects video playback from browser to client for improved performance"
    W365 = "Supported — same multimedia redirection capabilities via Windows App."
    Impact = "Low"
    Detection = "Informational — both platforms support this"
  }
  "RDPShortpath" = @{
    Feature = "RDP Shortpath (UDP transport)"
    AVD = "Supported — STUN direct + TURN relay for public and managed private networks"
    W365 = "Supported — enabled by default. STUN + TURN relay (GA). Private network Shortpath requires Azure Network Connection."
    Impact = "Low"
    Detection = "Informational — both platforms support this"
  }
}

# =========================================================
# Enhanced Analysis: Cost Analysis
# =========================================================
Write-ProgressSection -Section "Step 5: Cost Analysis" -Status Start -EstimatedMinutes 1 -Message "Calculating potential savings"

$totalCurrentCost = 0
$totalRecommendedCost = 0
$costSourceLabel = "PAYG Estimate"

# Build region lookup by VM name to avoid Where-Object per recommendation
$vmRegionLookup = @{}
foreach ($v in $vms) { if ($v.VMName) { $vmRegionLookup[$v.VMName] = $v.Region } }

# Build host pool lookup for VMs
$vmHostPoolLookup = @{}
foreach ($v in $vms) { if ($v.VMName -and $v.HostPoolName) { $vmHostPoolLookup[$v.VMName] = $v.HostPoolName } }

foreach ($rec in $vmRightSizing) {
  $region = $vmRegionLookup[$rec.VMName]
  
  # Use actual cost if available, otherwise PAYG estimate
  $currentCostInfo = Get-VmMonthlyCost -VMName $rec.VMName -VMSize $rec.CurrentSize -Region $region
  if ($currentCostInfo.Monthly) {
    $totalCurrentCost += $currentCostInfo.Monthly
    if ($currentCostInfo.Source -eq "Actual") { $costSourceLabel = "Actual Billing" }
  }
  
  if ($rec.RecommendedSize -ne "Keep Current" -and $rec.RecommendedSize -ne "Unknown") {
    # For recommended size, always use PAYG estimate (we don't have actual costs for a VM that doesn't exist yet)
    $recCost = Get-EstimatedVmCostPerHour -VmSize $rec.RecommendedSize -Region $region
    if ($recCost) {
      $totalRecommendedCost += $recCost * 730
    } elseif ($currentCostInfo.Monthly) {
      # If we can't price the recommended SKU, assume same cost
      $totalRecommendedCost += $currentCostInfo.Monthly
    }
  } else {
    if ($currentCostInfo.Monthly) {
      $totalRecommendedCost += $currentCostInfo.Monthly
    }
  }
}

# Fallback: if right-sizing loop produced zero (e.g., no metrics → no recommendations),
# compute fleet cost directly from VM list
if ($totalCurrentCost -eq 0 -and $vms.Count -gt 0) {
  foreach ($v in $vms) {
    $vcInfo = Get-VmMonthlyCost -VMName $v.VMName -VMSize $v.VMSize -Region $v.Region
    if ($vcInfo.Monthly) {
      $totalCurrentCost += $vcInfo.Monthly
      $totalRecommendedCost += $vcInfo.Monthly  # No recommendation → same cost
      if ($vcInfo.Source -eq "Actual") { $costSourceLabel = "Actual Billing" }
    }
  }
}

# If actual costs available, also compute per-host-pool cost breakdown
$hostPoolCosts = @{}
if ($actualCostData.Count -gt 0) {
  foreach ($ac in $actualCostData) {
    $hpn = $ac.HostPoolName
    if (-not $hpn) { continue }
    if (-not $hostPoolCosts.ContainsKey($hpn)) {
      $hostPoolCosts[$hpn] = @{ TotalMonthly = 0.0; ComputeMonthly = 0.0; StorageMonthly = 0.0; VMCount = 0; InfraMonthly = 0.0 }
    }
    $hostPoolCosts[$hpn].TotalMonthly += $ac.MonthlyEstimate
    $hostPoolCosts[$hpn].ComputeMonthly += $ac.ComputeMonthly
    $hostPoolCosts[$hpn].StorageMonthly += $ac.StorageMonthly
    $hostPoolCosts[$hpn].VMCount++
  }
}

# --- Aggregate infrastructure costs by category (v4.1) ---
$infraCostByCategory = @{}
$infraCostTotal = 0
if ($infraCostData.Count -gt 0) {
  foreach ($ic in $infraCostData) {
    $cat = $ic.MeterCategory
    if (-not $infraCostByCategory.ContainsKey($cat)) { $infraCostByCategory[$cat] = 0.0 }
    $infraCostByCategory[$cat] += $ic.MonthlyEstimate
    $infraCostTotal += $ic.MonthlyEstimate
  }
  # Distribute infra costs evenly across host pools (shared infrastructure)
  $hpCount = [math]::Max(1, $hostPoolCosts.Count)
  if ($hpCount -eq 0) { $hpCount = [math]::Max(1, @($hostPools).Count) }
  $infraPerPool = $infraCostTotal / $hpCount
  foreach ($hpn in @($hostPoolCosts.Keys)) {
    $hostPoolCosts[$hpn].InfraMonthly = [math]::Round($infraPerPool, 2)
    $hostPoolCosts[$hpn].TotalMonthly += $infraPerPool
  }
}

$potentialMonthlySavings = $totalCurrentCost - $totalRecommendedCost

$costAnalysis.Add([PSCustomObject]@{
  TotalVMs = ($vms | Measure-Object).Count
  EstimatedCurrentMonthlyCost = [math]::Round($totalCurrentCost, 2)
  EstimatedRecommendedMonthlyCost = [math]::Round($totalRecommendedCost, 2)
  PotentialMonthlySavings = [math]::Round($potentialMonthlySavings, 2)
  PotentialAnnualSavings = [math]::Round($potentialMonthlySavings * 12, 2)
  CostSource = $costSourceLabel
  SavingsPercentage = if ($totalCurrentCost -gt 0) { 
    [math]::Round(($potentialMonthlySavings / $totalCurrentCost) * 100, 1) 
  } else { 0 }
})

# Export per-host-pool costs if actual data available
if ($hostPoolCosts.Count -gt 0) {
  $hpCostExport = foreach ($hpn in $hostPoolCosts.Keys) {
    $hpc = $hostPoolCosts[$hpn]
    [PSCustomObject]@{
      HostPoolName    = $hpn
      VMCount         = $hpc.VMCount
      TotalMonthly    = [math]::Round($hpc.TotalMonthly, 2)
      ComputeMonthly  = [math]::Round($hpc.ComputeMonthly, 2)
      StorageMonthly  = [math]::Round($hpc.StorageMonthly, 2)
      AvgPerVM        = [math]::Round($hpc.TotalMonthly / [math]::Max(1, $hpc.VMCount), 2)
    }
  }
  $hpCostExport | Export-Csv (Join-Path $outFolder "ENHANCED-HostPool-Costs.csv") -NoTypeInformation
}

Write-ProgressSection -Section "Step 5: Cost Analysis" -Status Complete -Message "Potential monthly savings: ~`$$([math]::Round($potentialMonthlySavings, 0)) ($costSourceLabel)"

# --- Per-User Cost Calculation (v4.1) ---
foreach ($hp in $hostPools) {
  $hpName = $hp.HostPoolName
  if (-not $hpName) { continue }
  $poolMonthly = 0
  if ($hostPoolCosts.ContainsKey($hpName)) {
    $poolMonthly = $hostPoolCosts[$hpName].TotalMonthly
  } else {
    # Fallback: sum PAYG estimates from right-sizing data
    $hpRecs = @($vmRightSizing | Where-Object { $_.HostPoolName -eq $hpName -and $_.CurrentMonthlyCost })
    $poolMonthly = ($hpRecs | ForEach-Object { [double]$_.CurrentMonthlyCost } | Measure-Object -Sum).Sum
  }
  if ($poolMonthly -le 0) { continue }
  $hpSH = @($sessionHosts | Where-Object { $_.HostPoolName -eq $hpName })
  
  # User count: try multiple sources in priority order
  # 1. Assigned users (reliable for personal pools)
  $userCount = @($hpSH | Where-Object { $_.AssignedUser } | ForEach-Object { $_.AssignedUser } | Select-Object -Unique).Count
  
  # 2. Log Analytics unique users from connection success rate KQL (best for pooled)
  if ($userCount -eq 0 -and $connSuccessData.Count -gt 0) {
    $poolConnDataPU = @($connSuccessData | Where-Object { 
      $connHpName = $_.HostPool
      $connHpName -eq $hpName -or $connHpName -eq ($hpName -split '/')[-1] -or $hpName -match [regex]::Escape($connHpName)
    })
    if ($poolConnDataPU.Count -gt 0 -and $poolConnDataPU[0].PSObject.Properties.Name -contains 'UniqueUsers') {
      $laUsers = [int]$poolConnDataPU[0].UniqueUsers
      if ($laUsers -gt 0) { $userCount = $laUsers }
    }
  }
  
  # 3. Active sessions as lower bound (point-in-time)
  if ($userCount -eq 0) {
    $sessTotal = ($hpSH | ForEach-Object { if ($_.ActiveSessions) { [int]$_.ActiveSessions } else { 0 } } | Measure-Object -Sum).Sum
    if ($sessTotal -gt 0) { $userCount = $sessTotal }
  }
  
  # 4. Last resort: VM count (1 user per VM minimum)
  if ($userCount -eq 0) { $userCount = [math]::Max(1, $hpSH.Count) }
  
  $costPerUser = [math]::Round($poolMonthly / $userCount, 2)
  $drainCount = @($hpSH | Where-Object { $_.AllowNewSession -eq $false }).Count
  $effectiveCapacity = $hpSH.Count - $drainCount
  
  $userCountSource = if (@($hpSH | Where-Object { $_.AssignedUser }).Count -gt 0) { "Assigned" }
                     elseif ($connSuccessData.Count -gt 0 -and $userCount -gt $hpSH.Count) { "Log Analytics (30d)" }
                     else { "Active Sessions" }
  
  $perUserCost.Add([PSCustomObject]@{
    HostPoolName      = $hpName
    TotalMonthly      = [math]::Round($poolMonthly, 0)
    UserCount         = $userCount
    UserCountSource   = $userCountSource
    CostPerUser       = $costPerUser
    VMCount           = @($vms | Where-Object { $_.HostPoolName -eq $hpName }).Count
    DrainModeHosts    = $drainCount
    EffectiveCapacity = $effectiveCapacity
  })
}
if ($perUserCost.Count -gt 0) {
  $avgCostPerUser = [math]::Round(($perUserCost | ForEach-Object { $_.CostPerUser } | Measure-Object -Average).Average, 2)
  Write-Host "  Per-User Cost: avg `$$avgCostPerUser/user/month across $($perUserCost.Count) pool(s)" -ForegroundColor Gray
}

# =========================================================
# Reservation Analysis (v3.0)
# =========================================================
if ($IncludeReservationAnalysis) {
  Write-ProgressSection -Section "Step 5b: Reservation Analysis" -Status Start -EstimatedMinutes 2 -Message "Analyzing RI coverage and savings opportunities"
  
  # --- Collect existing reservations ---
  if ($hasAzReservations) {
    try {
      Import-Module Az.Reservations -ErrorAction Stop
      Write-Host "  Fetching existing reservations..." -ForegroundColor Gray
      
      # Enumerate via reservation orders — this is the reliable path
      $allOrders = @(Get-AzReservationOrder -ErrorAction Stop)
      Write-Host "    Found $($allOrders.Count) reservation order(s)" -ForegroundColor Gray
      
      foreach ($order in $allOrders) {
        $orderId = ($order.Id -split '/')[-1] ?? $order.Name
        if (-not $orderId) { continue }
        
        try {
          $orderReservations = @(Get-AzReservation -ReservationOrderId $orderId -ErrorAction Stop)
        } catch {
          Write-Host "    ⚠ Could not read order $orderId : $($_.Exception.Message)" -ForegroundColor Yellow
          continue
        }
        
        foreach ($res in $orderReservations) {
          # Defensive property extraction — Az.Reservations objects vary by module version
          # Some versions use $res.Sku (string), others $res.Sku.Name (object), others $res.SkuName
          $skuName = $null
          if ($res.PSObject.Properties['Sku']) {
            $skuName = if ($res.Sku -is [string]) { $res.Sku } elseif ($res.Sku.PSObject.Properties['Name']) { $res.Sku.Name } else { "$($res.Sku)" }
          }
          $skuName = $skuName ?? (SafeProp $res 'SkuName') ?? (SafeProp $res 'ReservedResourceType') ?? "Unknown"
          
          $location = (SafeProp $res 'Location') ?? ""
          $quantity = (SafeProp $res 'Quantity') ?? 0
          $provState = (SafeProp $res 'ProvisioningState') ?? (SafeProp $res 'State') ?? "Unknown"
          $displayName = (SafeProp $res 'DisplayName') ?? (SafeProp $res 'Name') ?? ""
          $term = (SafeProp $res 'Term') ?? ""
          $appliedScope = (SafeProp $res 'AppliedScopeType') ?? (SafeProp $res 'UserFriendlyAppliedScopeType') ?? ""
          
          # Expiry — try multiple property names
          $expiry = (SafeProp $res 'ExpiryDate') ?? (SafeProp $res 'ExpiryDateTime') ?? $null
          if ($expiry -and $expiry -is [string]) {
            try { $expiry = [datetime]::Parse($expiry) } catch { $expiry = $null }
          }
          
          $effectiveDate = (SafeProp $res 'EffectiveDateTime') ?? (SafeProp $res 'BenefitStartTime') ?? $null
          if ($effectiveDate -and $effectiveDate -is [string]) {
            try { $effectiveDate = [datetime]::Parse($effectiveDate) } catch { $effectiveDate = $null }
          }
          
          $existingReservations.Add([PSCustomObject]@{
            ReservationId       = $res.Id ?? ""
            ReservationName     = $displayName
            SKU                 = $skuName
            Location            = $location
            Quantity            = [int]$quantity
            ProvisioningState   = $provState
            ExpiryDate          = $expiry
            EffectiveDate       = $effectiveDate
            Term                = $term
            AppliedScopeType    = $appliedScope
            Status              = if ($provState -eq "Succeeded") { "Active" } else { $provState }
            DaysUntilExpiry     = if ($expiry) { [math]::Max(0, [math]::Round(($expiry - (Get-Date)).TotalDays, 0)) } else { "Unknown" }
          })
        }
      }
      Write-Host "  ✓ Found $($existingReservations.Count) reservation(s) across $($allOrders.Count) order(s)" -ForegroundColor Green
    }
    catch {
      Write-Host "  ⚠ Could not read reservations: $($_.Exception.Message)" -ForegroundColor Yellow
      Write-Host "    This usually means the account lacks Reservations Reader role at the tenant level" -ForegroundColor Gray
    }
  }
  
  # --- Build reservation coverage lookup ---
  # Map existing RIs by normalized SKU name + region for matching
  $riCoverageLookup = @{}
  foreach ($ri in ($existingReservations | Where-Object { $_.Status -eq "Active" })) {
    $key = "$($ri.SKU)|$($ri.Location)".ToLower()
    if (-not $riCoverageLookup.ContainsKey($key)) {
      $riCoverageLookup[$key] = @{ ReservedQty = 0; Reservations = @() }
    }
    $riCoverageLookup[$key].ReservedQty += $ri.Quantity
    $riCoverageLookup[$key].Reservations += $ri
  }
  
  # --- Analyze each VM for RI opportunity ---
  # Group VMs by SKU + Region to find bulk RI opportunities
  $vmSkuGroups = @{}
  foreach ($vm in $vms) {
    $key = "$($vm.VMSize)|$($vm.Region)".ToLower()
    if (-not $vmSkuGroups.ContainsKey($key)) {
      $vmSkuGroups[$key] = @{ Size = $vm.VMSize; Region = $vm.Region; VMs = @(); HostPools = @{} }
    }
    $vmSkuGroups[$key].VMs += $vm
    $vmSkuGroups[$key].HostPools[$vm.HostPoolName] = $true
  }
  
  # Determine host pool type/appgroup/load balancer lookups
  $hpTypeLookup = @{}
  $hpAppGroupLookup = @{}
  $hpLoadBalancerLookup = @{}
  foreach ($hp in $hostPools) { 
    if ($hp.HostPoolName) {
      $hpTypeLookup[$hp.HostPoolName] = $hp.HostPoolType
      $hpAppGroupLookup[$hp.HostPoolName] = $hp.PreferredAppGroupType
      $hpLoadBalancerLookup[$hp.HostPoolName] = $hp.LoadBalancer
    }
  }
  
  foreach ($groupKey in $vmSkuGroups.Keys) {
    $group = $vmSkuGroups[$groupKey]
    $vmSize = $group.Size
    $region = $group.Region
    $vmCount = @($group.VMs).Count
    $hostPoolNames = ($group.HostPools.Keys | Sort-Object) -join "; "
    
    # Determine if these are personal (always-on) or pooled
    $isAlwaysOn = $false
    foreach ($hpName in $group.HostPools.Keys) {
      if ($hpTypeLookup[$hpName] -eq "Personal") { $isAlwaysOn = $true; break }
    }
    
    # Check RI coverage
    $coverageKey = "$vmSize|$region".ToLower()
    $reservedQty = 0
    $existingRIs = @()
    if ($riCoverageLookup.ContainsKey($coverageKey)) {
      $reservedQty = $riCoverageLookup[$coverageKey].ReservedQty
      $existingRIs = $riCoverageLookup[$coverageKey].Reservations
    }
    
    $uncoveredCount = [math]::Max(0, $vmCount - $reservedQty)
    $overProvisionedCount = [math]::Max(0, $reservedQty - $vmCount)
    
    # Calculate savings potential
    $riPricing = $riPricingTable[$vmSize]
    $paygHourly = if ($riPricing) { $riPricing.PAYG } else { $null }
    $ri1yHourly = if ($riPricing) { $riPricing.RI1Y } else { $null }
    $ri3yHourly = if ($riPricing) { $riPricing.RI3Y } else { $null }
    
    $paygMonthly = if ($paygHourly) { [math]::Round($paygHourly * 730, 2) } else { "Unknown" }
    $ri1yMonthly = if ($ri1yHourly) { [math]::Round($ri1yHourly * 730, 2) } else { "Unknown" }
    $ri3yMonthly = if ($ri3yHourly) { [math]::Round($ri3yHourly * 730, 2) } else { "Unknown" }
    
    $savings1yPerVm = if ($paygHourly -and $ri1yHourly) { [math]::Round(($paygHourly - $ri1yHourly) * 730, 2) } else { 0 }
    $savings3yPerVm = if ($paygHourly -and $ri3yHourly) { [math]::Round(($paygHourly - $ri3yHourly) * 730, 2) } else { 0 }
    $savings1yPct = if ($paygHourly -and $ri1yHourly) { [math]::Round((1 - $ri1yHourly / $paygHourly) * 100, 0) } else { 0 }
    $savings3yPct = if ($paygHourly -and $ri3yHourly) { [math]::Round((1 - $ri3yHourly / $paygHourly) * 100, 0) } else { 0 }
    
    $totalSavings1y = [math]::Round($savings1yPerVm * $uncoveredCount, 2)
    $totalSavings3y = [math]::Round($savings3yPerVm * $uncoveredCount, 2)
    
    # Determine status and recommendation
    $status = "Uncovered"
    $recommendation = ""
    $priority = "Medium"
    
    if ($uncoveredCount -eq 0 -and $overProvisionedCount -eq 0) {
      $status = "Fully Covered"
      $recommendation = "No action needed — all $vmCount VM(s) covered by existing reservations"
      $priority = "None"
    }
    elseif ($uncoveredCount -eq 0 -and $overProvisionedCount -gt 0) {
      $status = "Over-Provisioned"
      $recommendation = "$overProvisionedCount excess RI(s) — consider exchanging or letting expire. Wasted cost: ~`$$([math]::Round($paygMonthly * $overProvisionedCount * 0.37, 0))/mo"
      $priority = "Medium"
    }
    elseif ($uncoveredCount -gt 0 -and $isAlwaysOn) {
      $status = "Uncovered (Always-On)"
      $recommendation = "Strong RI candidate — $uncoveredCount always-on personal desktop(s) paying PAYG rates. 3-year RI saves ~`$$totalSavings3y/mo ($savings3yPct%)"
      $priority = "High"
    }
    elseif ($uncoveredCount -gt 0 -and $vmCount -ge 3) {
      $status = "Uncovered"
      $recommendation = "$uncoveredCount VM(s) at PAYG rates. Consistent fleet of $vmCount suggests RI value. 1-year saves ~`$$totalSavings1y/mo ($savings1yPct%), 3-year saves ~`$$totalSavings3y/mo ($savings3yPct%)"
      $priority = "High"
    }
    elseif ($uncoveredCount -gt 0) {
      $status = "Uncovered (Small Fleet)"
      $recommendation = "$uncoveredCount VM(s) at PAYG. Small fleet — evaluate if workload is stable enough for commitment. 1-year saves ~`$$totalSavings1y/mo"
      $priority = "Low"
    }
    
    # Check for expiring reservations
    $expiringRIs = @($existingRIs | Where-Object { $_.DaysUntilExpiry -ne "Unknown" -and [int]$_.DaysUntilExpiry -le 90 })
    $expiryWarning = ""
    if ($expiringRIs.Count -gt 0) {
      $soonestExpiry = ($expiringRIs | Sort-Object DaysUntilExpiry | Select-Object -First 1).DaysUntilExpiry
      $expiryWarning = "⚠️ $($expiringRIs.Count) RI(s) expiring within 90 days (soonest: $soonestExpiry days)"
      if ($priority -eq "None") { $priority = "Medium" }
    }
    
    $reservationAnalysis.Add([PSCustomObject]@{
      VMSize              = $vmSize
      Region              = $region
      DeployedVMs         = $vmCount
      HostPools           = $hostPoolNames
      IsAlwaysOn          = $isAlwaysOn
      ReservedQty         = $reservedQty
      UncoveredVMs        = $uncoveredCount
      OverProvisionedRIs  = $overProvisionedCount
      Status              = $status
      PAYGMonthlyPerVM    = $paygMonthly
      RI1YMonthlyPerVM    = $ri1yMonthly
      RI3YMonthlyPerVM    = $ri3yMonthly
      Savings1YPerVM      = $savings1yPerVm
      Savings3YPerVM      = $savings3yPerVm
      Savings1YPct        = "$savings1yPct%"
      Savings3YPct        = "$savings3yPct%"
      TotalMonthlySavings1Y = $totalSavings1y
      TotalMonthlySavings3Y = $totalSavings3y
      ExpiryWarning       = $expiryWarning
      Priority            = $priority
      Recommendation      = $recommendation
    })
  }
  
  # Summary stats
  $totalUncoveredVMs = ($reservationAnalysis | Measure-Object -Property UncoveredVMs -Sum).Sum
  $totalRI1ySavings = ($reservationAnalysis | Measure-Object -Property TotalMonthlySavings1Y -Sum).Sum
  $totalRI3ySavings = ($reservationAnalysis | Measure-Object -Property TotalMonthlySavings3Y -Sum).Sum
  $totalOverProvisioned = ($reservationAnalysis | Measure-Object -Property OverProvisionedRIs -Sum).Sum
  
  Write-ProgressSection -Section "Step 5b: Reservation Analysis" -Status Complete -Message "$totalUncoveredVMs uncovered VMs | ~`$$([math]::Round($totalRI3ySavings, 0))/mo potential savings (3yr RI)"
}

# =========================================================
# Incident Window Comparative Analysis
# =========================================================
if ($IncludeIncidentWindowQueries -and (SafeCount $vmMetricsIncident) -gt 0) {
  Write-Host "Performing incident vs baseline comparative analysis..." -ForegroundColor Cyan
  
  # Pre-index incident metrics by VM ID (same pattern as baseline)
  $incidentMetricsByVm = @{}
  foreach ($metric in $vmMetricsIncident) {
    $vmId = $metric.VmId
    if (-not $vmId) { continue }
    if (-not $incidentMetricsByVm.ContainsKey($vmId)) {
      $incidentMetricsByVm[$vmId] = @{ CpuAvg = @(); CpuMax = @(); Mem = @() }
    }
    if ($metric.Metric -eq "Percentage CPU" -and $metric.Aggregation -eq "Average") {
      $incidentMetricsByVm[$vmId].CpuAvg += $metric
    } elseif ($metric.Metric -eq "Percentage CPU" -and $metric.Aggregation -eq "Maximum") {
      $incidentMetricsByVm[$vmId].CpuMax += $metric
    } elseif ($metric.Metric -eq "Available Memory Bytes") {
      $incidentMetricsByVm[$vmId].Mem += $metric
    }
  }
  
  foreach ($vm in $vms) {
    $vmId = $vm.VMId
    if (-not $vmId) { continue }
    
    # Baseline metrics from pre-built index
    $baselineData = $metricsByVm[$vmId]
    
    if ($baselineData) {
      $baselineCpuAvgMeasure = $baselineData.CpuAvg | Measure-Object Value -Average
      $baselineCpuMaxMeasure = $baselineData.CpuMax | Measure-Object Value -Maximum
      $baselineMemAvgMeasure = $baselineData.Mem | Where-Object { $_.Aggregation -eq "Average" } | Measure-Object Value -Average
      $baselineMemMinMeasure = $baselineData.Mem | Measure-Object Value -Minimum
    } else {
      $baselineCpuAvgMeasure = $null
      $baselineCpuMaxMeasure = $null
      $baselineMemAvgMeasure = $null
      $baselineMemMinMeasure = $null
    }
    
    $baselineAvgCpu = if ((SafeMeasure $baselineCpuAvgMeasure 'Average')) { (SafeMeasure $baselineCpuAvgMeasure 'Average') } else { 0 }
    $baselinePeakCpu = if ((SafeMeasure $baselineCpuMaxMeasure 'Maximum')) { (SafeMeasure $baselineCpuMaxMeasure 'Maximum') } else { 0 }
    $baselineAvgFreeMem = if ((SafeMeasure $baselineMemAvgMeasure 'Average')) { (SafeMeasure $baselineMemAvgMeasure 'Average') } else { 0 }
    $baselineMinFreeMem = if ((SafeMeasure $baselineMemMinMeasure 'Minimum')) { (SafeMeasure $baselineMemMinMeasure 'Minimum') } else { 0 }
    
    # Incident metrics from pre-built index
    $incidentData = $incidentMetricsByVm[$vmId]
    
    if ($incidentData) {
      $incidentCpuAvgMeasure = $incidentData.CpuAvg | Measure-Object Value -Average
      $incidentCpuMaxMeasure = $incidentData.CpuMax | Measure-Object Value -Maximum
      $incidentMemAvgMeasure = $incidentData.Mem | Where-Object { $_.Aggregation -eq "Average" } | Measure-Object Value -Average
      $incidentMemMinMeasure = $incidentData.Mem | Measure-Object Value -Minimum
    } else {
      $incidentCpuAvgMeasure = $null
      $incidentCpuMaxMeasure = $null
      $incidentMemAvgMeasure = $null
      $incidentMemMinMeasure = $null
    }
    
    $incidentAvgCpu = if ((SafeMeasure $incidentCpuAvgMeasure 'Average')) { (SafeMeasure $incidentCpuAvgMeasure 'Average') } else { 0 }
    $incidentPeakCpu = if ((SafeMeasure $incidentCpuMaxMeasure 'Maximum')) { (SafeMeasure $incidentCpuMaxMeasure 'Maximum') } else { 0 }
    $incidentAvgFreeMem = if ((SafeMeasure $incidentMemAvgMeasure 'Average')) { (SafeMeasure $incidentMemAvgMeasure 'Average') } else { 0 }
    $incidentMinFreeMem = if ((SafeMeasure $incidentMemMinMeasure 'Minimum')) { (SafeMeasure $incidentMemMinMeasure 'Minimum') } else { 0 }
    
    # Calculate differences
    $cpuAvgDelta = $incidentAvgCpu - $baselineAvgCpu
    $cpuPeakDelta = $incidentPeakCpu - $baselinePeakCpu
    $memAvgDelta = ($incidentAvgFreeMem - $baselineAvgFreeMem) / 1GB
    $memMinDelta = ($incidentMinFreeMem - $baselineMinFreeMem) / 1GB
    
    # Determine findings
    $findings = @()
    $severity = "Normal"
    
    if ($cpuAvgDelta -gt 20) {
      $findings += "Average CPU increased by $([math]::Round($cpuAvgDelta, 1))% during incident"
      $severity = "High"
    } elseif ($cpuAvgDelta -gt 10) {
      $findings += "Moderate CPU increase of $([math]::Round($cpuAvgDelta, 1))%"
      $severity = "Medium"
    } elseif ($cpuAvgDelta -lt -20) {
      $findings += "Average CPU decreased by $([math]::Round([math]::Abs($cpuAvgDelta), 1))% during incident"
      $severity = "Medium"
    }
    
    if ($cpuPeakDelta -gt 15) {
      $findings += "Peak CPU spiked $([math]::Round($cpuPeakDelta, 1))% higher"
      if ($severity -ne "High") { $severity = "Medium" }
    }
    
    if ($memMinDelta -lt -2) {
      $findings += "Memory pressure increased ($(([math]::Round([math]::Abs($memMinDelta), 1)))GB less free memory)"
      $severity = "High"
    } elseif ($memMinDelta -lt -1) {
      $findings += "Moderate memory pressure increase"
      if ($severity -eq "Normal") { $severity = "Medium" }
    }
    
    if ($incidentPeakCpu -gt 90) {
      $findings += "VM was critically overloaded (peak: $([math]::Round($incidentPeakCpu, 1))%)"
      $severity = "Critical"
    } elseif ($incidentPeakCpu -gt 80) {
      $findings += "VM was under high load (peak: $([math]::Round($incidentPeakCpu, 1))%)"
      if ($severity -ne "Critical") { $severity = "High" }
    }
    
    if ($incidentMinFreeMem -lt 512MB) {
      $findings += "Critically low memory during incident ($([math]::Round($incidentMinFreeMem / 1MB, 0))MB free)"
      $severity = "Critical"
    }
    
    if (@($findings).Count -eq 0) {
      $findings += "No significant performance changes detected"
    }
    
    $incidentAnalysis.Add([PSCustomObject]@{
      SubscriptionId = $vm.SubscriptionId
      ResourceGroup = $vm.ResourceGroup
      HostPoolName = $vm.HostPoolName
      VMName = $vm.VMName
      VMSize = $vm.VMSize
      
      # Baseline metrics
      BaselineAvgCPU = [math]::Round($baselineAvgCpu, 1)
      BaselinePeakCPU = [math]::Round($baselinePeakCpu, 1)
      BaselineAvgFreeMemoryGB = [math]::Round($baselineAvgFreeMem / 1GB, 2)
      BaselineMinFreeMemoryGB = [math]::Round($baselineMinFreeMem / 1GB, 2)
      
      # Incident metrics
      IncidentAvgCPU = [math]::Round($incidentAvgCpu, 1)
      IncidentPeakCPU = [math]::Round($incidentPeakCpu, 1)
      IncidentAvgFreeMemoryGB = [math]::Round($incidentAvgFreeMem / 1GB, 2)
      IncidentMinFreeMemoryGB = [math]::Round($incidentMinFreeMem / 1GB, 2)
      
      # Deltas
      CPUAvgDelta = [math]::Round($cpuAvgDelta, 1)
      CPUPeakDelta = [math]::Round($cpuPeakDelta, 1)
      MemoryAvgDeltaGB = [math]::Round($memAvgDelta, 2)
      MemoryMinDeltaGB = [math]::Round($memMinDelta, 2)
      
      # Analysis
      Severity = $severity
      Findings = ($findings -join "; ")
    })
  }
}

# =========================================================
# Export
# =========================================================
Write-ProgressSection -Section "Step 7: Exporting Results" -Status Start -EstimatedMinutes 1 -Message "Creating CSV files and reports"

# Core data
$hostPools              | Export-Csv (Join-Path $outFolder "AVD-HostPools.csv") -NoTypeInformation
$sessionHosts           | Export-Csv (Join-Path $outFolder "AVD-SessionHosts.csv") -NoTypeInformation
$vms                    | Export-Csv (Join-Path $outFolder "AVD-VMs.csv") -NoTypeInformation
$vmss                   | Export-Csv (Join-Path $outFolder "AVD-ScaleSets.csv") -NoTypeInformation
$vmssInstances          | Export-Csv (Join-Path $outFolder "AVD-ScaleSet-Instances.csv") -NoTypeInformation
$scalingPlans           | Export-Csv (Join-Path $outFolder "AVD-ScalingPlans.csv") -NoTypeInformation
$scalingPlanAssignments | Export-Csv (Join-Path $outFolder "AVD-ScalingPlanAssignments.csv") -NoTypeInformation
$scalingPlanSchedules   | Export-Csv (Join-Path $outFolder "AVD-ScalingPlanSchedules.csv") -NoTypeInformation
$vmMetrics              | Export-Csv (Join-Path $outFolder "AVD-VM-Metrics-Baseline.csv") -NoTypeInformation
$vmMetricsIncident      | Export-Csv (Join-Path $outFolder "AVD-VM-Metrics-Incident.csv") -NoTypeInformation
$laResults              | Export-Csv (Join-Path $outFolder "AVD-LogAnalytics-Results.csv") -NoTypeInformation

# Enhanced analysis
$vmRightSizing          | Export-Csv (Join-Path $outFolder "ENHANCED-VM-RightSizing-Recommendations.csv") -NoTypeInformation
$zoneResiliency         | Export-Csv (Join-Path $outFolder "ENHANCED-Zone-Resiliency-Analysis.csv") -NoTypeInformation
$costAnalysis           | Export-Csv (Join-Path $outFolder "ENHANCED-Cost-Analysis.csv") -NoTypeInformation
$incidentAnalysis       | Export-Csv (Join-Path $outFolder "ENHANCED-Incident-Comparative-Analysis.csv") -NoTypeInformation

# New analysis (v3.0.0)
$sessionHostHealth      | Export-Csv (Join-Path $outFolder "ENHANCED-SessionHost-Health.csv") -NoTypeInformation
$storageFindingsList    | Export-Csv (Join-Path $outFolder "ENHANCED-Storage-Optimization.csv") -NoTypeInformation
$accelNetFindings       | Export-Csv (Join-Path $outFolder "ENHANCED-AccelNet-Analysis.csv") -NoTypeInformation
$imageAnalysis          | Export-Csv (Join-Path $outFolder "ENHANCED-Image-Analysis.csv") -NoTypeInformation
if ($galleryAnalysis.Count -gt 0) {
  $galleryAnalysis      | Export-Csv (Join-Path $outFolder "ENHANCED-Gallery-Image-Versions.csv") -NoTypeInformation
}
$hpImageConsistency     | Export-Csv (Join-Path $outFolder "ENHANCED-HostPool-Image-Consistency.csv") -NoTypeInformation
$skuDiversityAnalysis   | Export-Csv (Join-Path $outFolder "ENHANCED-SKU-Diversity-Analysis.csv") -NoTypeInformation
$networkFindings        | Export-Csv (Join-Path $outFolder "ENHANCED-Network-Readiness.csv") -NoTypeInformation
if ($subnetAnalysis.Count -gt 0) {
  $subnetAnalysis       | Export-Csv (Join-Path $outFolder "ENHANCED-Subnet-Analysis.csv") -NoTypeInformation
}
if ($vnetAnalysis.Count -gt 0) {
  $vnetAnalysis         | Export-Csv (Join-Path $outFolder "ENHANCED-VNet-Analysis.csv") -NoTypeInformation
}

# Cross-region connection analysis (v4.0.0) — exported even before HTML enrichment
# The actual analysis is computed during HTML generation; export raw KQL data here
$crossRegionRawExport = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_CrossRegionConnections" -and $_.QueryName -eq "AVD" })
if ($crossRegionRawExport.Count -gt 0) {
  $crossRegionRawExport | Export-Csv (Join-Path $outFolder "ENHANCED-CrossRegion-Connections-Raw.csv") -NoTypeInformation
}

# Reservation analysis (v3.0.0)
if ($IncludeReservationAnalysis) {
  $reservationAnalysis    | Export-Csv (Join-Path $outFolder "ENHANCED-Reservation-Analysis.csv") -NoTypeInformation
  if ($existingReservations.Count -gt 0) {
    $existingReservations | Export-Csv (Join-Path $outFolder "ENHANCED-Existing-Reservations.csv") -NoTypeInformation
  }
}

# =========================================================
# PII Scrubbing — Post-process CSV files (v4.1)
# =========================================================
if ($ScrubPII) {
  Write-Host "`n🔒 PII Scrubbing enabled — anonymizing CSV data..." -ForegroundColor Cyan
  
  # Define PII column patterns and their scrub functions
  $piiColumns = @{
    'VMName'          = { param($v) Scrub-VMName $v }
    'SessionHostName' = { param($v) Scrub-VMName $v }
    'HostPoolName'    = { param($v) Scrub-HostPoolName $v }
    'AssignedUser'    = { param($v) Scrub-Username $v }
    'UserName'        = { param($v) Scrub-Username $v }
    'SubscriptionId'  = { param($v) Scrub-SubscriptionId $v }
    'ResourceGroup'   = { param($v) Scrub-ResourceGroup $v }
    'PrivateIp'       = { param($v) Scrub-IP $v }
    'SessionHostArmName' = { param($v) Scrub-VMName $v }
  }
  
  $csvFiles = Get-ChildItem (Join-Path $outFolder "*.csv") -ErrorAction SilentlyContinue
  $scrubCount = 0
  foreach ($csvFile in $csvFiles) {
    try {
      $rows = @(Import-Csv $csvFile.FullName)
      if ($rows.Count -eq 0) { continue }
      $headers = @($rows[0].PSObject.Properties.Name)
      $columnsToScrub = @($headers | Where-Object { $piiColumns.ContainsKey($_) })
      if ($columnsToScrub.Count -eq 0) { continue }
      
      foreach ($row in $rows) {
        foreach ($col in $columnsToScrub) {
          $val = $row.$col
          if (-not [string]::IsNullOrEmpty($val)) {
            $row.$col = & $piiColumns[$col] $val
          }
        }
      }
      $rows | Export-Csv $csvFile.FullName -NoTypeInformation
      $scrubCount++
    } catch {
      Write-Host "  ⚠ Could not scrub $($csvFile.Name): $_" -ForegroundColor Yellow
    }
  }
  Write-Host "  ✓ Scrubbed PII from $scrubCount CSV file(s)" -ForegroundColor Green
  Write-Host "  ✓ HTML report anonymized inline during generation" -ForegroundColor Green
  Write-Host "  Note: Same entity always maps to same anonymous ID (cross-referenceable)" -ForegroundColor Gray
}

# =========================================================
# Enhanced Analysis: Azure Advisor Recommendations
# =========================================================
$advisorRecommendations.Clear()

if ($IncludeAzureAdvisor) {
  # Estimate scales with subscription count
  $estimatedAdvisorMinutes = [math]::Max(1, (SafeCount $SubscriptionIds))  # ~1 min per subscription
  
  Write-ProgressSection -Section "Step 6: Azure Advisor Integration" -Status Start -EstimatedMinutes $estimatedAdvisorMinutes -Message "Collecting Microsoft recommendations"
  
  $subsProcessed = 0
  foreach ($subId in $SubscriptionIds) {
    $subsProcessed++
    Write-ProgressSection -Section "Step 6: Azure Advisor Integration" -Status Progress -Current $subsProcessed -Total (SafeCount $SubscriptionIds) -Message "Subscription: $subId"
    
    try {
      Set-AzContext -SubscriptionId $subId | Out-Null
      
      # Get all Advisor recommendations for this subscription
      $recommendations = Invoke-WithRetry -ScriptBlock {
        Get-AzAdvisorRecommendation -ErrorAction Stop
      } -OperationName "Get Azure Advisor Recommendations"
      
      foreach ($rec in $recommendations) {
        # Filter for recommendations related to AVD VMs
        $isRelevant = $false
        $vmName = $null
        
        # Check if this is a VM-related recommendation
        if ($rec.ImpactedValue -match "/virtualMachines/([^/]+)") {
          $vmName = $matches[1]
          # Check if this VM is in our AVD list
          $isRelevant = $vms | Where-Object { $_.VMName -eq $vmName }
        }
        
        if ($isRelevant -or $rec.Category -eq "Cost") {
          # Safely access properties that might not exist
          $shortDesc = if ($rec.PSObject.Properties.Name -contains 'ShortDescription' -and $rec.ShortDescription) {
            if ($rec.ShortDescription.PSObject.Properties.Name -contains 'Problem') {
              $rec.ShortDescription.Problem
            } else {
              $rec.Name
            }
          } else {
            $rec.Name
          }
          
          $solution = if ($rec.PSObject.Properties.Name -contains 'ShortDescription' -and $rec.ShortDescription) {
            if ($rec.ShortDescription.PSObject.Properties.Name -contains 'Solution') {
              $rec.ShortDescription.Solution
            } else {
              "See Azure Advisor portal for details"
            }
          } else {
            "See Azure Advisor portal for details"
          }
          
          $resourceId = if ($rec.PSObject.Properties.Name -contains 'ResourceMetadata' -and $rec.ResourceMetadata) {
            if ($rec.ResourceMetadata.PSObject.Properties.Name -contains 'ResourceId') {
              $rec.ResourceMetadata.ResourceId
            } else {
              $rec.ImpactedValue
            }
          } else {
            $rec.ImpactedValue
          }
          
          $advisorRecommendations.Add([PSCustomObject]@{
            SubscriptionId = $subId
            VMName = $vmName
            Category = $rec.Category
            Impact = $rec.Impact
            Risk = $rec.Risk
            RecommendationName = $rec.Name
            ShortDescription = $shortDesc
            Solution = $solution
            ImpactedField = $rec.ImpactedField
            ImpactedValue = $rec.ImpactedValue
            LastUpdated = $rec.LastUpdated
            ResourceId = $resourceId
          })
        }
      }
      
      Write-Host "  Found $(SafeCount $recommendations) Advisor recommendations in subscription" -ForegroundColor Gray
    }
    catch {
      Write-Host "  ⚠ Could not retrieve Advisor recommendations: $($_.Exception.Message)" -ForegroundColor Yellow
    }
  }
  
  Write-ProgressSection -Section "Step 6: Azure Advisor Integration" -Status Complete -Message "Collected $(SafeCount $advisorRecommendations) relevant recommendations"
  
  # Export Advisor recommendations
  if ((SafeCount $advisorRecommendations) -gt 0) {
    $advisorRecommendations | Export-Csv (Join-Path $outFolder "ENHANCED-Azure-Advisor-Recommendations.csv") -NoTypeInformation
  }
}
else {
  Write-ProgressSection -Section "Step 6: Azure Advisor Integration" -Status Skip -Message "-IncludeAzureAdvisor not specified (use -IncludeAzureAdvisor to enable)"
}

# Get current UTC time for timestamps
$nowUtc = (Get-Date).ToUniversalTime()

# Create README with important disclaimers
$readmeContent = @"
AVD EVIDENCE PACK - IMPORTANT INFORMATION
==========================================

COST ESTIMATES DISCLAIMER
-------------------------
All cost figures in this analysis are ESTIMATES ONLY based on standard Azure retail pricing.

Your actual costs will vary based on:
- Enterprise Agreement (EA) pricing
- Cloud Solution Provider (CSP) pricing
- Reserved Instance commitments
- Spot instance usage
- Negotiated contract rates
- Regional pricing variations
- Billing currency and exchange rates

RECOMMENDATIONS:
1. Use PERCENTAGES (e.g., "36% reduction") rather than absolute dollar amounts
2. Validate all cost figures against your actual Azure billing data
3. Recalculate savings using your specific contract rates
4. The EstimatedMonthlySavings column provides directional guidance only

FILES WITH COST ESTIMATES:
- ENHANCED-VM-RightSizing-Recommendations.csv (EstimatedMonthlySavings column)
- ENHANCED-Cost-Analysis.csv (PotentialMonthlySavings, PotentialAnnualSavings)
- ENHANCED-Executive-Summary.txt (Cost Optimization Opportunity section)

The optimization opportunities identified are real - just verify the dollar amounts
against your actual pricing before presenting to stakeholders.

ANALYSIS DATE: $($nowUtc.ToString("yyyy-MM-dd HH:mm:ss")) UTC
"@

$readmeContent | Out-File (Join-Path $outFolder "README-COST-DISCLAIMER.txt") -Encoding utf8

# =========================================================
# Enhanced Summary Report
# =========================================================

# Utilization summary (uses pre-built metricsByVm index)
$vmUtilizationSummary = @()
foreach ($vm in $vms) {
  $vmId = $vm.VMId
  if (-not $vmId) { continue }

  $vmMetricsData = $metricsByVm[$vmId]

  if ($vmMetricsData) {
    $cpuMaxMeasure = $vmMetricsData.CpuMax | Measure-Object Value -Maximum
    $cpuAvgMeasure = $vmMetricsData.CpuAvg | Measure-Object Value -Average
    $memMinMeasure = $vmMetricsData.Mem | Measure-Object Value -Minimum
  } else {
    $cpuMaxMeasure = $null
    $cpuAvgMeasure = $null
    $memMinMeasure = $null
  }

  $vmUtilizationSummary += [PSCustomObject]@{
    VMName        = $vm.VMName
    VMSize        = $vm.VMSize
    PeakCPU       = if ((SafeMeasure $cpuMaxMeasure 'Maximum')) { (SafeMeasure $cpuMaxMeasure 'Maximum') } else { 0 }
    AvgCPU        = if ((SafeMeasure $cpuAvgMeasure 'Average')) { (SafeMeasure $cpuAvgMeasure 'Average') } else { 0 }
    MinFreeMemory = if ((SafeMeasure $memMinMeasure 'Minimum')) { (SafeMeasure $memMinMeasure 'Minimum') } else { 0 }
  }
}

# Session density (uses pre-built sessionHostsByVm index)
$sessionDensitySummary = @()

# Build host pool MaxSessions lookup
$hpMaxSessionsLookup = @{}
foreach ($hp in $hostPools) { if ($hp.HostPoolName) { $hpMaxSessionsLookup[$hp.HostPoolName] = $hp.MaxSessions } }

foreach ($vm in $vms) {
  $vmSessions = $sessionHostsByVm[$vm.VMName]
  $sessionsSumMeasure = if ($vmSessions) { $vmSessions | Measure-Object ActiveSessions -Sum } else { $null }
  $sessions = if ((SafeMeasure $sessionsSumMeasure 'Sum')) { (SafeMeasure $sessionsSumMeasure 'Sum') } else { 0 }

  $maxSessions = $hpMaxSessionsLookup[$vm.HostPoolName]
  $density = if ($maxSessions -gt 0) { [math]::Round(($sessions / $maxSessions) * 100, 1) } else { 0 }

  $sessionDensitySummary += [PSCustomObject]@{
    VMName          = $vm.VMName
    VMSize          = $vm.VMSize
    ActiveSessions  = $sessions
    MaxSessions     = $maxSessions
    DensityPercent  = $density
  }
}

# Key findings (based on actual recommendations - mutually exclusive categories)
$downsizeCandidates = ($vmRightSizing | Where-Object { 
  $_.RecommendedSize -ne "Keep Current" -and 
  $_.RecommendedSize -ne "Unknown" -and
  $_.EstimatedMonthlySavings -ne "N/A" -and
  [double]$_.EstimatedMonthlySavings -gt 0
} | Measure-Object).Count

$upsizeCandidates = ($vmRightSizing | Where-Object { 
  $_.RecommendedSize -ne "Keep Current" -and 
  $_.RecommendedSize -ne "Unknown" -and
  ($_.EstimatedMonthlySavings -eq "N/A" -or [double]$_.EstimatedMonthlySavings -le 0)
} | Measure-Object).Count

$appropriatelySized = ($vmRightSizing | Where-Object { 
  $_.RecommendedSize -eq "Keep Current" -or $_.RecommendedSize -eq "Unknown"
} | Measure-Object).Count

$avgResiliencyMeasure = $zoneResiliency | Measure-Object ResiliencyScore -Average
$avgResiliencyScore = if ((SafeMeasure $avgResiliencyMeasure 'Average')) { (SafeMeasure $avgResiliencyMeasure 'Average') } else { 0 }

# Calculate capacity analysis metrics before summary
$peakCpuMeasure = $vmUtilizationSummary | Measure-Object PeakCPU -Maximum
$avgCpuMeasure = $vmUtilizationSummary | Measure-Object AvgCPU -Average
$peakSessionsMeasure = $sessionDensitySummary | Measure-Object ActiveSessions -Maximum
$avgDensityMeasure = $sessionDensitySummary | Measure-Object DensityPercent -Average

# Initialize variables that are built later (HTML generation) but referenced in summary
if (-not (Test-Path variable:crossRegionAnalysis)) { $crossRegionAnalysis = @() }
if (-not (Test-Path variable:priorityItems)) { $priorityItems = @() }
if (-not (Test-Path variable:w365Analysis)) { $w365Analysis = [System.Collections.Generic.List[object]]::new() }
if (-not (Test-Path variable:w365Candidates)) { $w365Candidates = @() }
if (-not (Test-Path variable:costAccessGranted)) { $costAccessGranted = [System.Collections.Generic.List[string]]::new() }
if (-not (Test-Path variable:costAccessDenied)) { $costAccessDenied = [System.Collections.Generic.List[string]]::new() }
if (-not (Test-Path variable:costSourceLabel)) { $costSourceLabel = "PAYG Estimate" }
if (-not (Test-Path variable:hostPoolCosts)) { $hostPoolCosts = @{} }
if (-not (Test-Path variable:securityPosture)) { $securityPosture = [System.Collections.Generic.List[object]]::new() }
if (-not (Test-Path variable:orphanedResources)) { $orphanedResources = [System.Collections.Generic.List[object]]::new() }
if (-not (Test-Path variable:orphanedWaste)) { $orphanedWaste = 0 }
if (-not (Test-Path variable:profileHealth)) { $profileHealth = [System.Collections.Generic.List[object]]::new() }
if (-not (Test-Path variable:uxScores)) { $uxScores = [System.Collections.Generic.List[object]]::new() }

$summary = [PSCustomObject]@{
  CollectedAtUtc = $nowUtc.ToString("o")
  TenantId = $TenantId
  Subscriptions = (SafeCount $SubscriptionIds)
  
  # Inventory
  HostPools = (SafeCount $hostPools)
  SessionHosts = (SafeCount $sessionHosts)
  VMs = (SafeCount $vms)
  ScaleSets = (SafeCount $vmss)
  ScaleSetInstances = (SafeCount $vmssInstances)
  AVDScaleSets = ($vmss | Where-Object { $_.IsAVD } | Measure-Object).Count
  ScalingPlans = (SafeCount $scalingPlans)
  
  # Right-Sizing Insights
  RightSizingAnalysis = @{
    TotalVMsAnalyzed = ($vmRightSizing | Measure-Object).Count
    DownsizeCandidates = $downsizeCandidates
    UpsizeCandidates = $upsizeCandidates
    AppropriatelySized = $appropriatelySized
    PotentialMonthlySavings = $costAnalysis[0].PotentialMonthlySavings
    PotentialAnnualSavings = $costAnalysis[0].PotentialAnnualSavings
    SavingsPercentage = "$($costAnalysis[0].SavingsPercentage)%"
  }
  
  # Zone Resiliency Insights
  ZoneResiliencyAnalysis = @{
    HostPoolsAnalyzed = (SafeCount $zoneResiliency)
    AverageResiliencyScore = [math]::Round($avgResiliencyScore, 0)
    HighResiliency = ($zoneResiliency | Where-Object { $_.ResiliencyScore -ge 75 } | Measure-Object).Count
    MediumResiliency = ($zoneResiliency | Where-Object { $_.ResiliencyScore -ge 40 -and $_.ResiliencyScore -lt 75 } | Measure-Object).Count
    LowResiliency = ($zoneResiliency | Where-Object { $_.ResiliencyScore -lt 40 } | Measure-Object).Count
    ZoneEnabledVMs = ($vms | Where-Object { $_.Zones } | Measure-Object).Count
    NonZoneVMs = ($vms | Where-Object { -not $_.Zones } | Measure-Object).Count
  }
  
  # Capacity Insights
  CapacityAnalysis = @{
    PeakCPU = if ((SafeMeasure $peakCpuMeasure 'Maximum')) { (SafeMeasure $peakCpuMeasure 'Maximum') } else { 0 }
    AvgCPU = if ((SafeMeasure $avgCpuMeasure 'Average')) { (SafeMeasure $avgCpuMeasure 'Average') } else { 0 }
    PeakSessions = if ((SafeMeasure $peakSessionsMeasure 'Maximum')) { (SafeMeasure $peakSessionsMeasure 'Maximum') } else { 0 }
    AvgSessionDensity = if ((SafeMeasure $avgDensityMeasure 'Average')) { (SafeMeasure $avgDensityMeasure 'Average') } else { 0 }
  }
  
  # Key Recommendations
  TopRecommendations = @(
    if ($downsizeCandidates -gt 0) { 
      "$downsizeCandidates VMs are underutilized and could be downsized for cost savings" 
    }
    if ($upsizeCandidates -gt 0) { 
      "$upsizeCandidates VMs show high CPU utilization and should be upsized" 
    }
    if ($avgResiliencyScore -lt 75) { 
      "Average zone resiliency score is $([math]::Round($avgResiliencyScore, 0))% - consider distributing VMs across availability zones" 
    }
    if (($vms | Where-Object { -not $_.Zones } | Measure-Object).Count -gt 0) { 
      "$(($vms | Where-Object { -not $_.Zones } | Measure-Object).Count) VMs are not zone-enabled" 
    }
    if ($costAnalysis[0].PotentialMonthlySavings -gt 100) { 
      "Potential monthly savings of `$$($costAnalysis[0].PotentialMonthlySavings) identified through right-sizing" 
    }
    if ($IncludeIncidentWindowQueries -and ($incidentAnalysis | Where-Object { $_.Severity -eq "Critical" } | Measure-Object).Count -gt 0) {
      "$(($incidentAnalysis | Where-Object { $_.Severity -eq 'Critical' } | Measure-Object).Count) VMs were critically overloaded during incident window"
    }
    if ($IncludeIncidentWindowQueries -and ($incidentAnalysis | Where-Object { $_.Severity -eq "High" } | Measure-Object).Count -gt 0) {
      "$(($incidentAnalysis | Where-Object { $_.Severity -eq 'High' } | Measure-Object).Count) VMs experienced high load during incident window"
    }
    # New recommendations (v3.0.0)
    if ($stuckHosts -gt 0) {
      "$stuckHosts session hosts are stuck in drain mode with stale heartbeats - investigate and recover"
    }
    if ($unavailableHosts -gt 0) {
      "$unavailableHosts session hosts have health issues requiring attention"
    }
    if ($premiumOnPooled -gt 0) {
      "$premiumOnPooled pooled VMs use Premium SSD - consider Standard SSD or Ephemeral OS disks"
    }
    if ($nonEphemeral -gt 0) {
      "$nonEphemeral pooled VMs use non-ephemeral OS disks - ephemeral disks improve performance and reduce cost"
    }
    if ($eligibleNotEnabled -gt 0) {
      "$eligibleNotEnabled VMs are eligible for Accelerated Networking but don't have it enabled"
    }
    if ($multiVersionImages -gt 0) {
      "$multiVersionImages image groups have version drift across session hosts - consider standardizing"
    }
    if ($goldenImageScore -lt 40) {
      "Golden Image maturity is low ($goldenImageScore/100, Grade: $goldenImageGrade) - implement Azure Compute Gallery with a golden image pipeline"
    }
    if ($marketplaceVms.Count -gt 0 -and $totalVmCount -gt 0 -and $marketplacePct -ge 50) {
      "$($marketplaceVms.Count) VMs ($marketplacePct%) using raw marketplace images - no golden image pipeline in place"
    }
    if (@($galleryAnalysis | Where-Object { $_.AgeDays -and $_.AgeDays -gt 90 }).Count -gt 0) {
      "$(@($galleryAnalysis | Where-Object { $_.AgeDays -and $_.AgeDays -gt 90 }).Count) gallery image(s) older than 90 days - missing security patches"
    }
    if ((SafeCount $networkFindings) -gt 0) {
      $critNet = @($networkFindings | Where-Object { $_.Impact -eq "High" }).Count
      if ($critNet -gt 0) { "$critNet critical network readiness issue(s) found - review Network tab for details" }
    }
    if ($shortpathSummary -and $shortpathSummary.ShortpathPct -lt 50) {
      "RDP Shortpath adoption is low ($($shortpathSummary.ShortpathPct)%) - enable UDP for better connection quality"
    }
    if (@($subnetAnalysis | Where-Object { $_.UsagePct -and [double]$_.UsagePct -gt 90 }).Count -gt 0) {
      "$(@($subnetAnalysis | Where-Object { $_.UsagePct -and [double]$_.UsagePct -gt 90 }).Count) subnet(s) at >90% IP capacity - cannot scale without expanding"
    }
    if (@($skuDiversityAnalysis | Where-Object { $_.OverallRisk -eq "High" }).Count -gt 0) {
      "$(@($skuDiversityAnalysis | Where-Object { $_.OverallRisk -eq "High" }).Count) host pool(s) have high allocation risk - single SKU family or single region"
    }
    if ($IncludeReservationAnalysis -and $totalUncoveredVMs -gt 0) {
      "$totalUncoveredVMs VMs without Reserved Instance coverage - potential savings of ~`$$([math]::Round($totalRI3ySavings, 0))/mo with 3-year RIs"
    }
    if ($IncludeReservationAnalysis -and $totalOverProvisioned -gt 0) {
      "$totalOverProvisioned over-provisioned RI(s) - consider exchanging or re-scoping"
    }
    if ($w365Candidates.Count -gt 0) {
      "$($w365Candidates.Count) host pool(s) are candidates for W365 Cloud PC migration - evaluate for cost savings and simplified management"
    }
  )
  
  # Session Host Health (v3.0.0)
  SessionHostHealth = @{
    TotalHosts = (SafeCount $sessionHostHealth)
    Healthy = ($sessionHostHealth | Where-Object { $_.Severity -eq "Normal" } | Measure-Object).Count
    DrainMode = $drainedHosts
    StuckInDrain = $stuckHosts
    Unavailable = ($sessionHostHealth | Where-Object { $_.Finding -match "Unavailable" } | Measure-Object).Count
    StaleHeartbeat = ($sessionHostHealth | Where-Object { $_.Finding -match "Stale" } | Measure-Object).Count
  }
  
  # Storage Optimization (v3.0.0)
  StorageOptimization = @{
    PremiumOnPooled = $premiumOnPooled
    NonEphemeralPooled = $nonEphemeral
    StandardHDD = ($storageFindingsList | Where-Object { $_.Findings -match "Standard HDD" } | Measure-Object).Count
  }
  
  # Network (v3.0.0)
  AcceleratedNetworking = @{
    EligibleNotEnabled = $eligibleNotEnabled
    Enabled = ($accelNetFindings | Where-Object { $_.AccelNetEnabled } | Measure-Object).Count
    NotEligible = ($accelNetFindings | Where-Object { -not $_.Eligible } | Measure-Object).Count
  }
  
  # Image Analysis (v3.0.0, enhanced v4.0.0)
  ImageAnalysis = @{
    UniqueImages = (SafeCount $imageAnalysis)
    MultiVersionGroups = $multiVersionImages
    GoldenImageGrade = $goldenImageGrade
    GoldenImageScore = $goldenImageScore
    MarketplaceVMs = $marketplaceVms.Count
    GalleryVMs = $galleryVms.Count
    ManagedImageVMs = $managedImageVms.Count
    StaleGalleryImages = @($galleryAnalysis | Where-Object { $_.AgeDays -and $_.AgeDays -gt 90 }).Count
    ConsistentPools = @($hpImageConsistency | Where-Object { $_.Consistency -eq "Consistent" }).Count
    InconsistentPools = @($hpImageConsistency | Where-Object { $_.Consistency -ne "Consistent" }).Count
  }
  
  # Network Readiness (v4.0.0)
  NetworkReadiness = @{
    TotalFindings = (SafeCount $networkFindings)
    CriticalFindings = @($networkFindings | Where-Object { $_.Impact -eq "High" }).Count
    WarningFindings = @($networkFindings | Where-Object { $_.Impact -eq "Medium" }).Count
    SubnetsAnalyzed = (SafeCount $subnetAnalysis)
    SubnetsCritical = @($subnetAnalysis | Where-Object { $_.UsagePct -and [double]$_.UsagePct -gt 90 }).Count
    SubnetsWarning = @($subnetAnalysis | Where-Object { $_.UsagePct -and [double]$_.UsagePct -gt 70 -and [double]$_.UsagePct -le 90 }).Count
    RDPShortpathPct = if ($shortpathSummary) { $shortpathSummary.ShortpathPct } else { $null }
    PrivateEndpoints = @($networkFindings | Where-Object { $_.Check -match "Private Endpoint" -and $_.Status -eq "OK" }).Count
    DisconnectedPeerings = @($networkFindings | Where-Object { $_.Check -match "Peering" -and $_.Status -ne "OK" }).Count
    NoNSGHosts = @($networkFindings | Where-Object { $_.Check -match "NSG" -and $_.Status -ne "OK" }).Count
  }
  
  # SKU Diversity & Allocation Resilience (v4.0.0)
  SKUDiversity = @{
    HostPoolsAnalyzed = (SafeCount $skuDiversityAnalysis)
    HighRisk = @($skuDiversityAnalysis | Where-Object { $_.OverallRisk -eq "High" }).Count
    MediumRisk = @($skuDiversityAnalysis | Where-Object { $_.OverallRisk -eq "Medium" }).Count
    LowRisk = @($skuDiversityAnalysis | Where-Object { $_.OverallRisk -eq "Low" }).Count
  }
  
  # Cross-Region Connections (v4.0.0)
  CrossRegion = if ((SafeCount $crossRegionAnalysis) -gt 0) {
    @{
      Enabled = $true
      TotalPaths = (SafeCount $crossRegionAnalysis)
      CrossContinentPaths = @($crossRegionAnalysis | Where-Object { $_.CrossContinent -eq $true }).Count
      HighLatencyPaths = @($crossRegionAnalysis | Where-Object { $_.ExcessRTTms -and [double]$_.ExcessRTTms -gt 50 }).Count
    }
  } else {
    @{ Enabled = $false }
  }
  
  # W365 Cloud PC Readiness (v4.0.0)
  W365Readiness = @{
    HostPoolsAnalyzed = (SafeCount $w365Analysis)
    StrongCandidates = @($w365Analysis | Where-Object { $_.Recommendation -eq "Strong W365 Candidate" }).Count
    ConsiderHybrid = @($w365Analysis | Where-Object { $_.Recommendation -match "Consider" }).Count
    KeepAVD = @($w365Analysis | Where-Object { $_.Recommendation -eq "Keep AVD" }).Count
  }
  
  # Actual Cost Intelligence (v4.0.0)
  # Security Posture (v4.0.0)
  SecurityPosture = @{
    HostPoolsAnalyzed = (SafeCount $securityPosture)
    PoolsBelowGradeC = @($securityPosture | Where-Object { $_.SecurityScore -lt 60 }).Count
    AvgSecurityScore = if ($securityPosture.Count -gt 0) { [math]::Round(($securityPosture | ForEach-Object { $_.SecurityScore } | Measure-Object -Average).Average, 0) } else { 0 }
  }
  
  # Orphaned Resources (v4.0.0)
  OrphanedResources = @{
    TotalFound = (SafeCount $orphanedResources)
    OrphanedDisks = @($orphanedResources | Where-Object { $_.ResourceType -eq "Disk" }).Count
    OrphanedNICs = @($orphanedResources | Where-Object { $_.ResourceType -eq "NIC" }).Count
    OrphanedPublicIPs = @($orphanedResources | Where-Object { $_.ResourceType -eq "PublicIP" }).Count
    EstMonthlyWaste = $orphanedWaste
  }
  
  # Profile Health (v4.0.0)
  ProfileHealth = @{
    HostsAnalyzed = (SafeCount $profileHealth)
    SlowProfiles = @($profileHealth | Where-Object { $_.Severity -ne "Good" }).Count
    CriticalProfiles = @($profileHealth | Where-Object { $_.Severity -eq "Critical" }).Count
  }
  
  # User Experience Scores (v4.0.0)
  UserExperience = @{
    HostPoolsScored = (SafeCount $uxScores)
    PoolsBelowGradeC = @($uxScores | Where-Object { $_.UXScore -lt 60 }).Count
    AvgUXScore = if ($uxScores.Count -gt 0) { [math]::Round(($uxScores | ForEach-Object { $_.UXScore } | Measure-Object -Average).Average, 0) } else { 0 }
  }
  
  ActualCosts = @{
    Enabled = (-not $SkipActualCosts)
    VMsWithBillingData = $vmActualMonthlyCost.Count
    SubscriptionsAccessible = if (-not $SkipActualCosts) { $costAccessGranted.Count } else { 0 }
    SubscriptionsDenied = if (-not $SkipActualCosts) { $costAccessDenied.Count } else { 0 }
    TotalMonthlyEstimate = if ($actualCostData.Count -gt 0) { [math]::Round(($actualCostData | ForEach-Object { $_.MonthlyEstimate } | Measure-Object -Sum).Sum, 0) } else { 0 }
  }
  
  # Reservation Analysis (v3.0.0)
  ReservationAnalysis = if ($IncludeReservationAnalysis -and (SafeCount $reservationAnalysis) -gt 0) {
    @{
      Enabled = $true
      SKUGroupsAnalyzed = (SafeCount $reservationAnalysis)
      TotalUncoveredVMs = $totalUncoveredVMs
      TotalOverProvisionedRIs = $totalOverProvisioned
      ExistingReservations = (SafeCount $existingReservations)
      PotentialMonthlySavings1Y = [math]::Round($totalRI1ySavings, 0)
      PotentialMonthlySavings3Y = [math]::Round($totalRI3ySavings, 0)
    }
  } else {
    @{
      Enabled = $false
      Note = "Reservation analysis not requested. Use -IncludeReservationAnalysis to enable."
    }
  }
  
  # Incident Window Analysis Summary
  IncidentWindowAnalysis = if ($IncludeIncidentWindowQueries -and (SafeCount $incidentAnalysis) -gt 0) {
    @{
      Enabled = $true
      VMsAnalyzed = (SafeCount $incidentAnalysis)
      CriticalIssues = ($incidentAnalysis | Where-Object { $_.Severity -eq "Critical" } | Measure-Object).Count
      HighIssues = ($incidentAnalysis | Where-Object { $_.Severity -eq "High" } | Measure-Object).Count
      MediumIssues = ($incidentAnalysis | Where-Object { $_.Severity -eq "Medium" } | Measure-Object).Count
      NoIssues = ($incidentAnalysis | Where-Object { $_.Severity -eq "Normal" } | Measure-Object).Count
      AvgCPUIncrease = $(  $m = SafeMeasure ($incidentAnalysis | Measure-Object CPUAvgDelta -Average) 'Average'; if ($m) { [math]::Round($m, 1) } else { 0 }  )
      MaxCPUIncrease = $(  $m = SafeMeasure ($incidentAnalysis | Measure-Object CPUAvgDelta -Maximum) 'Maximum'; if ($m) { [math]::Round($m, 1) } else { 0 }  )
      TimeWindow = "$($IncidentWindowStart.ToString('yyyy-MM-dd HH:mm')) to $($IncidentWindowEnd.ToString('yyyy-MM-dd HH:mm'))"
    }
  } else {
    @{
      Enabled = $false
      Note = "Incident window analysis not requested or no data available"
    }
  }
  
  # Data Quality
  MetricsCollected = -not $SkipAzureMonitorMetrics
  IncidentMetricsCollected = $IncludeIncidentWindowQueries -and (SafeCount $vmMetricsIncident) -gt 0
  LogAnalyticsQueriesExecuted = -not $SkipLogAnalyticsQueries
  MetricDataPoints = (SafeCount $vmMetrics)
  LogAnalyticsRows = (SafeCount $laResults)
}

$summary | ConvertTo-Json -Depth 10 | Out-File (Join-Path $outFolder "ENHANCED-Summary.json") -Encoding utf8

# Create executive summary text file
$executiveSummary = @"
=======================================================================
ENHANCED AVD ENVIRONMENT ANALYSIS - EXECUTIVE SUMMARY  (v4.0.0)
=======================================================================
Generated: $($nowUtc.ToString("yyyy-MM-dd HH:mm:ss UTC"))

ENVIRONMENT OVERVIEW
---------------------------------------------------------------------
Host Pools: $($summary.HostPools)
Session Hosts: $($summary.SessionHosts)
Virtual Machines: $($summary.VMs)
Virtual Machine Scale Sets: $($summary.ScaleSets) ($($summary.AVDScaleSets) used for AVD)
Scale Set Instances: $($summary.ScaleSetInstances)
Scaling Plans: $($summary.ScalingPlans)

RIGHT-SIZING ANALYSIS
---------------------------------------------------------------------
VMs Analyzed: $($summary.RightSizingAnalysis.TotalVMsAnalyzed)
Downsize Candidates: $($summary.RightSizingAnalysis.DownsizeCandidates) VMs
Upsize Candidates: $($summary.RightSizingAnalysis.UpsizeCandidates) VMs
Appropriately Sized: $($summary.RightSizingAnalysis.AppropriatelySized) VMs

COST OPTIMIZATION OPPORTUNITY
---------------------------------------------------------------------
Potential Monthly Savings: `$$($summary.RightSizingAnalysis.PotentialMonthlySavings) (ESTIMATED)
Potential Annual Savings: `$$($summary.RightSizingAnalysis.PotentialAnnualSavings) (ESTIMATED)
Savings Percentage: $($summary.RightSizingAnalysis.SavingsPercentage)

IMPORTANT: Cost estimates are based on standard Azure retail pricing.
Actual costs and savings will vary based on your specific contract terms,
enterprise agreements, CSP pricing, and negotiated rates. Use percentages
for more reliable comparison. Validate all figures against actual billing data.

ZONE RESILIENCY ANALYSIS
---------------------------------------------------------------------
Average Resiliency Score: $($summary.ZoneResiliencyAnalysis.AverageResiliencyScore)/100
High Resiliency (75-100): $($summary.ZoneResiliencyAnalysis.HighResiliency) host pools
Medium Resiliency (40-74): $($summary.ZoneResiliencyAnalysis.MediumResiliency) host pools
Low Resiliency (0-39): $($summary.ZoneResiliencyAnalysis.LowResiliency) host pools

Zone Distribution:
  - Zone-Enabled VMs: $($summary.ZoneResiliencyAnalysis.ZoneEnabledVMs)
  - Non-Zone VMs: $($summary.ZoneResiliencyAnalysis.NonZoneVMs)

CAPACITY & UTILIZATION
---------------------------------------------------------------------
Peak CPU: $([math]::Round($summary.CapacityAnalysis.PeakCPU, 1))%
Average CPU: $([math]::Round($summary.CapacityAnalysis.AvgCPU, 1))%
Peak Concurrent Sessions: $($summary.CapacityAnalysis.PeakSessions)
Average Session Density: $([math]::Round($summary.CapacityAnalysis.AvgSessionDensity, 1))%

SESSION HOST HEALTH
---------------------------------------------------------------------
Total Session Hosts: $($summary.SessionHostHealth.TotalHosts)
Healthy: $($summary.SessionHostHealth.Healthy)
Drain Mode Active: $($summary.SessionHostHealth.DrainMode)
Stuck in Drain: $($summary.SessionHostHealth.StuckInDrain)
Unavailable: $($summary.SessionHostHealth.Unavailable)
Stale Heartbeat: $($summary.SessionHostHealth.StaleHeartbeat)

NETWORK READINESS
---------------------------------------------------------------------
Total Findings: $($summary.NetworkReadiness.TotalFindings) ($($summary.NetworkReadiness.CriticalFindings) critical, $($summary.NetworkReadiness.WarningFindings) warning)
Subnets Analyzed: $($summary.NetworkReadiness.SubnetsAnalyzed) ($($summary.NetworkReadiness.SubnetsCritical) critical capacity, $($summary.NetworkReadiness.SubnetsWarning) warning)
RDP Shortpath: $(if ($null -ne $summary.NetworkReadiness.RDPShortpathPct) { "$($summary.NetworkReadiness.RDPShortpathPct)% of connections using UDP" } else { "No KQL data available" })
Disconnected VNet Peerings: $($summary.NetworkReadiness.DisconnectedPeerings)
Hosts Without NSG: $($summary.NetworkReadiness.NoNSGHosts)

STORAGE & DISK OPTIMIZATION
---------------------------------------------------------------------
Premium SSD on Pooled VMs: $($summary.StorageOptimization.PremiumOnPooled) (consider Standard SSD or Ephemeral)
Non-Ephemeral Pooled VMs: $($summary.StorageOptimization.NonEphemeralPooled) (ephemeral reduces cost + improves perf)
Standard HDD: $($summary.StorageOptimization.StandardHDD) (too slow for AVD, upgrade to SSD)

ACCELERATED NETWORKING
---------------------------------------------------------------------
Eligible but Not Enabled: $($summary.AcceleratedNetworking.EligibleNotEnabled) VMs
Enabled: $($summary.AcceleratedNetworking.Enabled) VMs
Not Eligible (< 4 vCPU): $($summary.AcceleratedNetworking.NotEligible) VMs

GOLDEN IMAGE ASSESSMENT
---------------------------------------------------------------------
Maturity Score: $($summary.ImageAnalysis.GoldenImageScore)/100 (Grade: $($summary.ImageAnalysis.GoldenImageGrade))
Image Sources: $($summary.ImageAnalysis.GalleryVMs) Gallery, $($summary.ImageAnalysis.MarketplaceVMs) Marketplace, $($summary.ImageAnalysis.ManagedImageVMs) Managed Image
Unique Image Groups: $($summary.ImageAnalysis.UniqueImages)
Groups with Version Drift: $($summary.ImageAnalysis.MultiVersionGroups)
Stale Gallery Images (>90 days): $($summary.ImageAnalysis.StaleGalleryImages)
Consistent Host Pools: $($summary.ImageAnalysis.ConsistentPools) / $(($summary.ImageAnalysis.ConsistentPools + $summary.ImageAnalysis.InconsistentPools))

SKU DIVERSITY & ALLOCATION RESILIENCE
---------------------------------------------------------------------
Host Pools Analyzed: $($summary.SKUDiversity.HostPoolsAnalyzed)
High Allocation Risk: $($summary.SKUDiversity.HighRisk) host pool(s)
Medium Allocation Risk: $($summary.SKUDiversity.MediumRisk) host pool(s)
Low Allocation Risk: $($summary.SKUDiversity.LowRisk) host pool(s)

W365 CLOUD PC READINESS
---------------------------------------------------------------------
Host Pools Analyzed: $($summary.W365Readiness.HostPoolsAnalyzed)
Strong W365 Candidates: $($summary.W365Readiness.StrongCandidates) host pool(s)
Consider W365 / Hybrid: $($summary.W365Readiness.ConsiderHybrid) host pool(s)
Keep AVD: $($summary.W365Readiness.KeepAVD) host pool(s)

SECURITY POSTURE
---------------------------------------------------------------------
Host Pools Analyzed: $($summary.SecurityPosture.HostPoolsAnalyzed)
Average Security Score: $($summary.SecurityPosture.AvgSecurityScore)/100
Pools Below Grade C: $($summary.SecurityPosture.PoolsBelowGradeC)

ORPHANED RESOURCES
---------------------------------------------------------------------
Total Found: $($summary.OrphanedResources.TotalFound)
Disks: $($summary.OrphanedResources.OrphanedDisks) | NICs: $($summary.OrphanedResources.OrphanedNICs) | Public IPs: $($summary.OrphanedResources.OrphanedPublicIPs)
Estimated Monthly Waste: `$$($summary.OrphanedResources.EstMonthlyWaste)

USER EXPERIENCE
---------------------------------------------------------------------
Host Pools Scored: $($summary.UserExperience.HostPoolsScored)
Average UX Score: $($summary.UserExperience.AvgUXScore)/100
Pools Below Grade C: $($summary.UserExperience.PoolsBelowGradeC)
$(if ($summary.ActualCosts.Enabled) { @"

ACTUAL COST INTELLIGENCE
---------------------------------------------------------------------
VMs with Billing Data: $($summary.ActualCosts.VMsWithBillingData)
Total Monthly Estimate: `$$($summary.ActualCosts.TotalMonthlyEstimate)
Source: Azure Cost Management API (last 30 days)
"@ })

$(if ($summary.CrossRegion.Enabled) {@"
CROSS-REGION CONNECTION ANALYSIS
---------------------------------------------------------------------
Connection Paths Analyzed: $($summary.CrossRegion.TotalPaths)
Cross-Continent Paths: $($summary.CrossRegion.CrossContinentPaths)
High Latency Paths (>50ms excess): $($summary.CrossRegion.HighLatencyPaths)

"@ })$(if ($summary.IncidentWindowAnalysis.Enabled) {@"
INCIDENT WINDOW COMPARATIVE ANALYSIS
---------------------------------------------------------------------
Analysis Period: $($summary.IncidentWindowAnalysis.TimeWindow)
VMs Analyzed: $($summary.IncidentWindowAnalysis.VMsAnalyzed)

Performance Impact During Incident:
  - Critical Issues: $($summary.IncidentWindowAnalysis.CriticalIssues) VMs
  - High Impact: $($summary.IncidentWindowAnalysis.HighIssues) VMs
  - Medium Impact: $($summary.IncidentWindowAnalysis.MediumIssues) VMs
  - No Significant Change: $($summary.IncidentWindowAnalysis.NoIssues) VMs

Average CPU Increase: $($summary.IncidentWindowAnalysis.AvgCPUIncrease)%
Maximum CPU Increase: $($summary.IncidentWindowAnalysis.MaxCPUIncrease)%

"@ })$(if ($summary.ReservationAnalysis.Enabled) {@"
RESERVATION ANALYSIS
---------------------------------------------------------------------
SKU Groups Analyzed: $($summary.ReservationAnalysis.SKUGroupsAnalyzed)
Uncovered VMs: $($summary.ReservationAnalysis.TotalUncoveredVMs)
Over-Provisioned RIs: $($summary.ReservationAnalysis.TotalOverProvisionedRIs)
Potential Monthly Savings (1-yr RI): `$$($summary.ReservationAnalysis.PotentialMonthlySavings1Y)
Potential Monthly Savings (3-yr RI): `$$($summary.ReservationAnalysis.PotentialMonthlySavings3Y)

"@ })TOP RECOMMENDATIONS
---------------------------------------------------------------------
$($summary.TopRecommendations -join "`n")

PRIORITY MATRIX SUMMARY
---------------------------------------------------------------------
$(if ((SafeCount $priorityItems) -gt 0) {
  $qw = @($priorityItems | Where-Object { $_.QuickWin }).Count
  $plan = @($priorityItems | Where-Object { -not $_.QuickWin -and $_.Impact -eq "High" }).Count
  $consider = @($priorityItems | Where-Object { -not $_.QuickWin -and $_.Impact -ne "High" }).Count
  "Quick Wins (high impact, low effort): $qw items`nPlan (high impact, high effort): $plan items`nConsider (lower impact): $consider items"
} else { "No priority items generated" })

DETAILED REPORTS
---------------------------------------------------------------------
Please review the following files for detailed analysis:
- ENHANCED-Analysis-Report.html (interactive HTML dashboard)
- ENHANCED-VM-RightSizing-Recommendations.csv
- ENHANCED-Zone-Resiliency-Analysis.csv
- ENHANCED-Cost-Analysis.csv
- ENHANCED-SessionHost-Health.csv
- ENHANCED-Storage-Optimization.csv
- ENHANCED-AccelNet-Analysis.csv
- ENHANCED-Image-Analysis.csv
- ENHANCED-Gallery-Image-Versions.csv
- ENHANCED-HostPool-Image-Consistency.csv
- ENHANCED-Network-Readiness.csv
- ENHANCED-Subnet-Analysis.csv
- ENHANCED-VNet-Analysis.csv
- ENHANCED-SKU-Diversity-Analysis.csv
- ENHANCED-CrossRegion-Analysis.csv
$(if ($summary.IncidentWindowAnalysis.Enabled) {"- ENHANCED-Incident-Comparative-Analysis.csv"})
$(if ($summary.ReservationAnalysis.Enabled) {"- ENHANCED-Reservation-Analysis.csv"})

=======================================================================
"@

$executiveSummary | Out-File (Join-Path $outFolder "ENHANCED-Executive-Summary.txt") -Encoding utf8

# =========================================================
# Optional ZIP
# =========================================================
if ($CreateZip) {
  $zipPath = "$outFolder.zip"
  if (Test-Path $zipPath) { Remove-Item $zipPath -Force -ErrorAction SilentlyContinue }
  Compress-Archive -Path (Join-Path $outFolder "*") -DestinationPath $zipPath -Force
  Write-Host ""
  Write-Host "ZIP archive created: $zipPath" -ForegroundColor Green
}

# =========================================================
# Optional HTML Report
# =========================================================
if ($GenerateHtmlReport) {
  Write-ProgressSection -Section "Step 8: Generating HTML Report" -Status Start -Message "Creating visual report"
  
  # -------------------------------------------------------
  # Branding
  # -------------------------------------------------------
  $logoBase64 = ""
  if ($LogoPath -and (Test-Path $LogoPath)) {
    try {
      $logoBytes = [System.IO.File]::ReadAllBytes($LogoPath)
      $logoBase64 = [System.Convert]::ToBase64String($logoBytes)
      $logoExt = [System.IO.Path]::GetExtension($LogoPath).TrimStart('.').ToLower()
      if ($logoExt -eq 'jpg') { $logoExt = 'jpeg' }
      $logoDataUri = "data:image/$logoExt;base64,$logoBase64"
      Write-Host "  ✓ Logo embedded: $LogoPath" -ForegroundColor Green
    } catch {
      Write-Host "  ⚠ Could not read logo file: $LogoPath" -ForegroundColor Yellow
      $logoDataUri = ""
    }
  } else {
    $logoDataUri = ""
  }
  $brandName = if ($CompanyName) { $CompanyName } else { "" }
  $brandAnalyst = if ($AnalystName) { $AnalystName } else { "" }
  
  # -------------------------------------------------------
  # Pre-process data for HTML dashboard
  # -------------------------------------------------------
  
  # CPU distribution buckets for chart
  $cpuBuckets = @{ "0-20" = 0; "20-40" = 0; "40-60" = 0; "60-80" = 0; "80-100" = 0 }
  foreach ($vm in $vmRightSizing) {
    $cpu = $vm.AvgCPU
    if     ($cpu -lt 20) { $cpuBuckets["0-20"]++ }
    elseif ($cpu -lt 40) { $cpuBuckets["20-40"]++ }
    elseif ($cpu -lt 60) { $cpuBuckets["40-60"]++ }
    elseif ($cpu -lt 80) { $cpuBuckets["60-80"]++ }
    else                 { $cpuBuckets["80-100"]++ }
  }
  $cpuBucketMax = ($cpuBuckets.Values | Measure-Object -Maximum).Maximum
  if ($cpuBucketMax -eq 0) { $cpuBucketMax = 1 }
  
  # Extract KQL results by label — filter to rows that have actual data properties
  $profileLoadData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ProfileLoadPerformance" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "SessionHostName" })
  $connQualityData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionQuality" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "ClientOS" })
  $connQualityByRegionData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionQualityByRegion" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "GatewayRegion" })
  $connErrorData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionErrors" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "CodeSymbolic" })
  $disconnectData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_Disconnects" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "SessionHostName" })
  $disconnectReasonData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_DisconnectReasons" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "DisconnectCategory" })
  $disconnectsByHostData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_DisconnectsByHost" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "SessionHostName" })
  $concurrencyData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_HourlyConcurrency" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "HourOfDay" -and $_.HourOfDay -ne $null -and $_.HourOfDay -ne "" })
  $crossRegionRawData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_CrossRegionConnections" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "GatewayRegion" })
  
  # -------------------------------------------------------
  # Cross-Region Connection Analysis
  # -------------------------------------------------------
  # Azure region geographic coordinates and continent mapping
  $azureRegionGeo = @{
    # North America
    "eastus"             = @{ Lat = 37.37; Lon = -79.13; Continent = "NA"; FriendlyName = "East US" }
    "eastus2"            = @{ Lat = 36.67; Lon = -78.93; Continent = "NA"; FriendlyName = "East US 2" }
    "centralus"          = @{ Lat = 41.59; Lon = -93.62; Continent = "NA"; FriendlyName = "Central US" }
    "northcentralus"     = @{ Lat = 41.88; Lon = -87.63; Continent = "NA"; FriendlyName = "North Central US" }
    "southcentralus"     = @{ Lat = 29.43; Lon = -98.49; Continent = "NA"; FriendlyName = "South Central US" }
    "westcentralus"      = @{ Lat = 40.89; Lon = -110.23; Continent = "NA"; FriendlyName = "West Central US" }
    "westus"             = @{ Lat = 37.78; Lon = -122.42; Continent = "NA"; FriendlyName = "West US" }
    "westus2"            = @{ Lat = 47.23; Lon = -119.85; Continent = "NA"; FriendlyName = "West US 2" }
    "westus3"            = @{ Lat = 33.45; Lon = -112.07; Continent = "NA"; FriendlyName = "West US 3" }
    "canadacentral"      = @{ Lat = 43.65; Lon = -79.38; Continent = "NA"; FriendlyName = "Canada Central" }
    "canadaeast"         = @{ Lat = 46.82; Lon = -71.25; Continent = "NA"; FriendlyName = "Canada East" }
    # Europe
    "northeurope"        = @{ Lat = 53.35; Lon = -6.26; Continent = "EU"; FriendlyName = "North Europe" }
    "westeurope"         = @{ Lat = 52.37; Lon = 4.90; Continent = "EU"; FriendlyName = "West Europe" }
    "uksouth"            = @{ Lat = 51.51; Lon = -0.13; Continent = "EU"; FriendlyName = "UK South" }
    "ukwest"             = @{ Lat = 51.46; Lon = -3.18; Continent = "EU"; FriendlyName = "UK West" }
    "francecentral"      = @{ Lat = 46.82; Lon = 2.35; Continent = "EU"; FriendlyName = "France Central" }
    "francesouth"        = @{ Lat = 43.30; Lon = 3.13; Continent = "EU"; FriendlyName = "France South" }
    "germanywestcentral" = @{ Lat = 50.11; Lon = 8.68; Continent = "EU"; FriendlyName = "Germany West Central" }
    "germanynorth"       = @{ Lat = 53.07; Lon = 8.81; Continent = "EU"; FriendlyName = "Germany North" }
    "switzerlandnorth"   = @{ Lat = 47.45; Lon = 8.56; Continent = "EU"; FriendlyName = "Switzerland North" }
    "switzerlandwest"    = @{ Lat = 46.52; Lon = 6.14; Continent = "EU"; FriendlyName = "Switzerland West" }
    "norwayeast"         = @{ Lat = 59.91; Lon = 10.75; Continent = "EU"; FriendlyName = "Norway East" }
    "norwaywest"         = @{ Lat = 58.97; Lon = 5.73; Continent = "EU"; FriendlyName = "Norway West" }
    "swedencentral"      = @{ Lat = 60.67; Lon = 17.14; Continent = "EU"; FriendlyName = "Sweden Central" }
    "polandcentral"      = @{ Lat = 52.23; Lon = 21.01; Continent = "EU"; FriendlyName = "Poland Central" }
    "italynorth"         = @{ Lat = 45.47; Lon = 9.19; Continent = "EU"; FriendlyName = "Italy North" }
    "spaincentral"       = @{ Lat = 40.42; Lon = -3.70; Continent = "EU"; FriendlyName = "Spain Central" }
    # Asia Pacific
    "eastasia"           = @{ Lat = 22.27; Lon = 114.17; Continent = "APAC"; FriendlyName = "East Asia" }
    "southeastasia"      = @{ Lat = 1.35; Lon = 103.82; Continent = "APAC"; FriendlyName = "Southeast Asia" }
    "japaneast"          = @{ Lat = 35.68; Lon = 139.77; Continent = "APAC"; FriendlyName = "Japan East" }
    "japanwest"          = @{ Lat = 34.69; Lon = 135.50; Continent = "APAC"; FriendlyName = "Japan West" }
    "australiaeast"      = @{ Lat = -33.87; Lon = 151.21; Continent = "APAC"; FriendlyName = "Australia East" }
    "australiasoutheast" = @{ Lat = -37.81; Lon = 144.96; Continent = "APAC"; FriendlyName = "Australia Southeast" }
    "australiacentral"   = @{ Lat = -35.28; Lon = 149.13; Continent = "APAC"; FriendlyName = "Australia Central" }
    "koreacentral"       = @{ Lat = 37.57; Lon = 126.98; Continent = "APAC"; FriendlyName = "Korea Central" }
    "koreasouth"         = @{ Lat = 35.18; Lon = 129.08; Continent = "APAC"; FriendlyName = "Korea South" }
    "centralindia"       = @{ Lat = 18.97; Lon = 72.82; Continent = "APAC"; FriendlyName = "Central India" }
    "southindia"         = @{ Lat = 12.97; Lon = 80.18; Continent = "APAC"; FriendlyName = "South India" }
    "westindia"          = @{ Lat = 19.08; Lon = 72.88; Continent = "APAC"; FriendlyName = "West India" }
    "indonesiacentral"   = @{ Lat = -6.21; Lon = 106.85; Continent = "APAC"; FriendlyName = "Indonesia Central" }
    "malaysiawest"       = @{ Lat = 3.14; Lon = 101.69; Continent = "APAC"; FriendlyName = "Malaysia West" }
    "newzealandnorth"    = @{ Lat = -36.85; Lon = 174.76; Continent = "APAC"; FriendlyName = "New Zealand North" }
    "taiwannorth"        = @{ Lat = 25.03; Lon = 121.57; Continent = "APAC"; FriendlyName = "Taiwan North" }
    # Middle East & Africa
    "uaenorth"           = @{ Lat = 25.27; Lon = 55.30; Continent = "MEA"; FriendlyName = "UAE North" }
    "uaecentral"         = @{ Lat = 24.45; Lon = 54.65; Continent = "MEA"; FriendlyName = "UAE Central" }
    "southafricanorth"   = @{ Lat = -25.73; Lon = 28.22; Continent = "MEA"; FriendlyName = "South Africa North" }
    "southafricawest"    = @{ Lat = -33.93; Lon = 18.42; Continent = "MEA"; FriendlyName = "South Africa West" }
    "qatarcentral"       = @{ Lat = 25.29; Lon = 51.53; Continent = "MEA"; FriendlyName = "Qatar Central" }
    "israelcentral"      = @{ Lat = 31.78; Lon = 35.22; Continent = "MEA"; FriendlyName = "Israel Central" }
    # South America
    "brazilsouth"        = @{ Lat = -23.55; Lon = -46.63; Continent = "SA"; FriendlyName = "Brazil South" }
    "brazilsoutheast"    = @{ Lat = -22.91; Lon = -43.17; Continent = "SA"; FriendlyName = "Brazil Southeast" }
    "mexicocentral"      = @{ Lat = 20.59; Lon = -100.39; Continent = "SA"; FriendlyName = "Mexico Central" }
    "chilecentral"       = @{ Lat = -33.45; Lon = -70.67; Continent = "SA"; FriendlyName = "Chile Central" }
  }
  
  $continentNames = @{
    "NA"   = "North America"
    "EU"   = "Europe"
    "APAC" = "Asia Pacific"
    "MEA"  = "Middle East & Africa"
    "SA"   = "South America"
  }
  
  # Haversine distance function (km)
  function Get-GeoDistanceKm ([double]$lat1, [double]$lon1, [double]$lat2, [double]$lon2) {
    $R = 6371.0 # Earth radius km
    $dLat = [math]::PI / 180.0 * ($lat2 - $lat1)
    $dLon = [math]::PI / 180.0 * ($lon2 - $lon1)
    $a = [math]::Sin($dLat / 2) * [math]::Sin($dLat / 2) +
         [math]::Cos([math]::PI / 180.0 * $lat1) * [math]::Cos([math]::PI / 180.0 * $lat2) *
         [math]::Sin($dLon / 2) * [math]::Sin($dLon / 2)
    $c = 2 * [math]::Atan2([math]::Sqrt($a), [math]::Sqrt(1 - $a))
    return [math]::Round($R * $c, 0)
  }
  
  # Estimate baseline RTT from distance: ~0.01ms per km (fiber speed ~200km/ms) + overhead
  # Formula: baselineRTT = (distance_km / 100) + 10ms overhead
  # This is conservative — real-world is often worse due to routing hops
  function Get-ExpectedBaselineRTTms ([double]$distanceKm) {
    return [math]::Round(($distanceKm / 100.0) + 10, 0)
  }
  
  # Build host pool region lookup: SessionHostName → host pool region
  $sessionHostRegionMap = @{}
  foreach ($sh in $sessionHosts) {
    $hostRegion = $null
    # Try to get region from VM data first (more precise)
    $vmMatch = $vms | Where-Object { $_.SessionHostName -eq $sh.SessionHostName } | Select-Object -First 1
    if ($vmMatch -and $vmMatch.Region) {
      $hostRegion = $vmMatch.Region.ToLower().Replace(" ", "")
    } else {
      # Fall back to host pool location
      $hpMatch = $hostPools | Where-Object { $_.HostPoolName -eq $sh.HostPoolName } | Select-Object -First 1
      if ($hpMatch -and $hpMatch.Location) {
        $hostRegion = $hpMatch.Location.ToLower().Replace(" ", "")
      }
    }
    if ($hostRegion) {
      $sessionHostRegionMap[$sh.SessionHostName] = $hostRegion
    }
  }
  
  # Enrich cross-region data with geographic analysis
  $crossRegionAnalysis = [System.Collections.Generic.List[object]]::new()
  foreach ($cr in $crossRegionRawData) {
    $gwRegion = ($cr.GatewayRegion -replace '\s', '').ToLower()
    $hostName = $cr.SessionHostName
    $hostRegion = $sessionHostRegionMap[$hostName]
    
    # Try to match gateway region to our geo map (KQL returns mixed casing)
    $gwGeo = $azureRegionGeo[$gwRegion]
    $hostGeo = if ($hostRegion) { $azureRegionGeo[$hostRegion] } else { $null }
    
    $distanceKm = 0
    $expectedRTT = 0
    $isCrossRegion = $false
    $isCrossContinent = $false
    $gwContinent = "Unknown"
    $hostContinent = "Unknown"
    $gwFriendly = $cr.GatewayRegion
    $hostFriendly = if ($hostRegion) { $hostRegion } else { "Unknown" }
    
    if ($gwGeo) {
      $gwContinent = $continentNames[$gwGeo.Continent] ?? $gwGeo.Continent
      $gwFriendly = $gwGeo.FriendlyName
    }
    if ($hostGeo) {
      $hostContinent = $continentNames[$hostGeo.Continent] ?? $hostGeo.Continent
      $hostFriendly = $hostGeo.FriendlyName
    }
    
    if ($gwGeo -and $hostGeo) {
      $distanceKm = Get-GeoDistanceKm $gwGeo.Lat $gwGeo.Lon $hostGeo.Lat $hostGeo.Lon
      $expectedRTT = Get-ExpectedBaselineRTTms $distanceKm
      $isCrossRegion = ($gwRegion -ne $hostRegion)
      $isCrossContinent = ($gwGeo.Continent -ne $hostGeo.Continent)
    } elseif ($hostRegion) {
      $isCrossRegion = ($gwRegion -ne $hostRegion)
    }
    
    $actualP95 = [double]($cr.P95RTTms)
    $rttExcess = if ($expectedRTT -gt 0) { [math]::Round($actualP95 - $expectedRTT, 0) } else { 0 }
    $rttRating = if ($actualP95 -le 50) { "Excellent" }
                 elseif ($actualP95 -le 100) { "Good" }
                 elseif ($actualP95 -le 150) { "Acceptable" }
                 elseif ($actualP95 -le 250) { "Poor" }
                 else { "Critical" }
    
    $crossRegionAnalysis.Add([PSCustomObject]@{
      GatewayRegion     = $gwFriendly
      GatewayContinent  = $gwContinent
      SessionHostName   = $hostName
      HostRegion        = $hostFriendly
      HostContinent     = $hostContinent
      DistanceKm        = $distanceKm
      DistanceMi        = [math]::Round($distanceKm * 0.621371, 0)
      ExpectedRTTms     = $expectedRTT
      AvgRTTms          = $cr.AvgRTTms
      P50RTTms          = $cr.P50RTTms
      P95RTTms          = $cr.P95RTTms
      MaxRTTms          = $cr.MaxRTTms
      RTTExcessMs       = $rttExcess
      RTTRating         = $rttRating
      AvgBandwidthKBps  = $cr.AvgBandwidthKBps
      MinBandwidthKBps  = $cr.MinBandwidthKBps
      Connections       = $cr.Connections
      DistinctUsers     = $cr.DistinctUsers
      IsCrossRegion     = $isCrossRegion
      IsCrossContinent  = $isCrossContinent
    })
  }
  
  $crossContinentPaths = @($crossRegionAnalysis | Where-Object { $_.IsCrossContinent })
  $crossRegionPaths = @($crossRegionAnalysis | Where-Object { $_.IsCrossRegion -and -not $_.IsCrossContinent })
  $sameRegionPaths = @($crossRegionAnalysis | Where-Object { -not $_.IsCrossRegion })
  
  $hasKqlData = ($profileLoadData.Count + $connQualityData.Count + $connErrorData.Count + $disconnectData.Count + $disconnectReasonData.Count + $crossRegionRawData.Count) -gt 0
  
  # Check which queries returned data rows but no matching properties (schema mismatch)
  # Don't flag queries that failed entirely (table doesn't exist) — that's a different issue
  $kqlNoData = @()
  $kqlFailed = @()
  if ($profileLoadData.Count -eq 0) {
    $meta = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ProfileLoadPerformance" -and $_.QueryName -eq "Meta" })
    if ($meta.Count -gt 0 -and $meta[0].PSObject.Properties['Status'] -and $meta[0].Status -eq "QueryFailed") { $kqlFailed += "Profile Load" }
    elseif (@($laResults | Where-Object { $_.Label -eq "CurrentWindow_ProfileLoadPerformance" }).Count -gt 0) { $kqlNoData += "Profile Load" }
  }
  if ($connQualityData.Count -eq 0) {
    $meta = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionQuality" -and $_.QueryName -eq "Meta" })
    if ($meta.Count -gt 0 -and $meta[0].PSObject.Properties['Status'] -and $meta[0].Status -eq "QueryFailed") { $kqlFailed += "Connection Quality" }
    elseif (@($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionQuality" }).Count -gt 0) { $kqlNoData += "Connection Quality" }
  }
  if ($connErrorData.Count -eq 0) {
    $meta = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionErrors" -and $_.QueryName -eq "Meta" })
    if ($meta.Count -gt 0 -and $meta[0].PSObject.Properties['Status'] -and $meta[0].Status -eq "QueryFailed") { $kqlFailed += "Connection Errors" }
    elseif (@($laResults | Where-Object { $_.Label -eq "CurrentWindow_ConnectionErrors" }).Count -gt 0) { $kqlNoData += "Connection Errors" }
  }
  if ($disconnectData.Count -eq 0) {
    $meta = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_Disconnects" -and $_.QueryName -eq "Meta" })
    if ($meta.Count -gt 0 -and $meta[0].PSObject.Properties['Status'] -and $meta[0].Status -eq "QueryFailed") { $kqlFailed += "Disconnects" }
    elseif (@($laResults | Where-Object { $_.Label -eq "CurrentWindow_Disconnects" }).Count -gt 0) { $kqlNoData += "Disconnects" }
  }
  
  # Cost per VM for cost breakdown table
  $costBreakdown = @($vmRightSizing | Where-Object { $_.EstimatedMonthlySavings -ne "N/A" -and [double]$_.EstimatedMonthlySavings -ne 0 } | 
    Sort-Object { [math]::Abs([double]$_.EstimatedMonthlySavings) } -Descending |
    Select-Object -First 30)
  
  $htmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AVD Analysis Report - $(Get-Date -Format 'yyyy-MM-dd')</title>
    <style>
        :root { --blue: #0078d4; --green: #28a745; --yellow: #ffc107; --red: #dc3545; --orange: #fd7e14; --gray: #6c757d; --light: #f8f9fa; --white: #fff; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: #eef2f7; color: #333; line-height: 1.5; }
        .container { max-width: 1400px; margin: 0 auto; padding: 24px; }
        
        /* Header */
        .header { background: linear-gradient(135deg, #0078d4, #005a9e); color: white; padding: 32px; border-radius: 12px; margin-bottom: 24px; }
        .header h1 { font-size: 28px; font-weight: 600; margin-bottom: 8px; }
        .header-meta { display: flex; gap: 24px; font-size: 14px; opacity: 0.9; flex-wrap: wrap; }
        
        /* Navigation */
        .nav { display: flex; gap: 4px; background: var(--white); border-radius: 10px; padding: 6px; margin-bottom: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); overflow-x: auto; }
        .nav-btn { padding: 10px 16px; border: none; background: none; cursor: pointer; border-radius: 8px; font-size: 13px; font-weight: 500; color: var(--gray); white-space: nowrap; transition: all 0.2s; }
        .nav-btn:hover { background: var(--light); color: #333; }
        .nav-btn.active { background: var(--blue); color: white; }
        
        /* Sections */
        .section { display: none; }
        .section.active { display: block; }
        
        /* Cards */
        .card-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }
        .card { background: var(--white); border-radius: 10px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
        .card-label { font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--gray); margin-bottom: 4px; }
        .card-value { font-size: 28px; font-weight: 700; }
        .card-sub { font-size: 12px; color: var(--gray); margin-top: 4px; }
        .card-value.blue { color: var(--blue); }
        .card-value.green { color: var(--green); }
        .card-value.yellow { color: #d4a017; }
        .card-value.red { color: var(--red); }
        .card-value.orange { color: var(--orange); }
        
        /* Alert boxes */
        .alert { padding: 16px 20px; border-radius: 8px; margin-bottom: 16px; font-size: 14px; display: flex; align-items: center; gap: 10px; }
        .alert-danger { background: #fdecea; border-left: 4px solid var(--red); }
        .alert-warning { background: #fff8e1; border-left: 4px solid var(--yellow); }
        .alert-success { background: #e8f5e9; border-left: 4px solid var(--green); }
        .alert-info { background: #e3f2fd; border-left: 4px solid var(--blue); }
        
        /* Tables */
        .table-wrap { background: var(--white); border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); overflow: hidden; margin-bottom: 24px; }
        .table-title { padding: 16px 20px; font-weight: 600; font-size: 16px; border-bottom: 1px solid #eee; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th { background: #f8f9fb; padding: 10px 16px; text-align: left; font-weight: 600; color: #555; border-bottom: 2px solid #eee; position: sticky; top: 0; cursor: pointer; user-select: none; }
        th:hover { background: #eef1f5; }
        th.sorted-asc::after { content: ' ▲'; font-size: 10px; }
        th.sorted-desc::after { content: ' ▼'; font-size: 10px; }
        td { padding: 10px 16px; border-bottom: 1px solid #f0f0f0; }
        tr:hover { background: #f8fafc; }
        .table-scroll { max-height: 500px; overflow-y: auto; }
        
        /* Badges */
        .badge { padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; display: inline-block; }
        .b-red { background: #fdecea; color: #c62828; }
        .b-orange { background: #fff3e0; color: #e65100; }
        .b-yellow { background: #fff8e1; color: #f57f17; }
        .b-green { background: #e8f5e9; color: #2e7d32; }
        .b-blue { background: #e3f2fd; color: #0d47a1; }
        .b-gray { background: #f5f5f5; color: #616161; }
        .b-blue { background: #e3f2fd; color: #0d47a1; }
        .b-gray { background: #f5f5f5; color: #666; }
        
        /* Evidence bar */
        .evidence-bar { display: inline-flex; align-items: center; gap: 6px; }
        .evidence-fill { height: 8px; border-radius: 4px; background: #e0e0e0; width: 60px; overflow: hidden; display: inline-block; vertical-align: middle; }
        .evidence-fill-inner { height: 100%; border-radius: 4px; transition: width 0.3s; }
        
        /* Search & Filter */
        .toolbar { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
        .search-input { padding: 8px 14px; border: 1px solid #ddd; border-radius: 8px; font-size: 13px; width: 260px; }
        .search-input:focus { outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px rgba(0,120,212,0.1); }
        .filter-btn { padding: 6px 14px; border: 1px solid #ddd; border-radius: 8px; background: white; cursor: pointer; font-size: 12px; }
        .filter-btn.active { background: var(--blue); color: white; border-color: var(--blue); }
        .count-badge { font-size: 11px; background: rgba(0,0,0,0.1); padding: 1px 6px; border-radius: 8px; margin-left: 4px; }
        
        /* Two column */
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }
        @media (max-width: 768px) { .two-col { grid-template-columns: 1fr; } .card-grid { grid-template-columns: 1fr 1fr; } }
        @media print {
            .nav, .toolbar, .disclaimer { display: none !important; }
            .section { display: block !important; page-break-inside: avoid; }
            .header { background: #fff !important; color: #333 !important; border-bottom: 2px solid #0078d4; }
            .header h1, .header div { color: #333 !important; }
            .header-meta span { color: #666 !important; }
            .container { max-width: 100%; padding: 0; }
            body { font-size: 11px; }
            .print-btn { display: none !important; }
        }
        
        /* Disclaimer */
        .disclaimer { background: #fff3cd; border: 1px solid #ffc107; padding: 14px 18px; border-radius: 8px; font-size: 13px; margin-bottom: 24px; }
        .footer { text-align: center; padding: 24px; color: var(--gray); font-size: 12px; }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <div style="display:flex;align-items:center;gap:20px;justify-content:space-between;flex-wrap:wrap">
            <div style="display:flex;align-items:center;gap:16px">
$(if ($logoDataUri) { "                <img src='$logoDataUri' alt='Logo' style='height:48px;width:auto;border-radius:6px'>" })
                <div>
                    <h1 style="margin:0">Azure Virtual Desktop — Environment Analysis</h1>
$(if ($brandName) { "                    <div style='font-size:14px;color:#a0c4ff;margin-top:4px'>Prepared by $brandName$(if ($brandAnalyst) { " · $brandAnalyst" })</div>" })
                </div>
            </div>
$(if ($brandName) { "            <div style='text-align:right'><div style='font-size:11px;color:#8ab4f8;opacity:0.8'>CONFIDENTIAL<br>Prepared for client review</div><button onclick='window.print()' class='print-btn' style='margin-top:8px;padding:6px 16px;background:#fff;color:#0078d4;border:1px solid #0078d4;border-radius:4px;cursor:pointer;font-size:12px'>📄 Export PDF</button></div>" })
        </div>
        <div class="header-meta">
            <span>📅 $(Get-Date -Format 'yyyy-MM-dd HH:mm') UTC</span>
            <span>📊 $MetricsLookbackDays-day analysis window</span>
            <span>🖥️ $($summary.VMs) VMs across $($summary.HostPools) host pools</span>
        </div>
    </div>
    
    <div class="disclaimer">
        <strong>⚠️ Cost Disclaimer:</strong> All cost estimates use Azure retail pricing. Actual costs vary by EA/CSP agreements, reserved instances, and contract rates. Use percentages for comparison; validate dollar amounts against billing.
    </div>
$(if ($ScrubPII) {
@"
    <div style="background:#fff3e0;border:2px solid #ff9800;border-radius:8px;padding:12px 16px;margin:8px 0">
        <strong>🔒 PII Scrubbed Report</strong> — All usernames, VM names, host pool names, subscription IDs, resource groups, and IP addresses have been anonymized. Consistent hashing ensures the same entity always maps to the same anonymous ID across all tabs and CSV files, so cross-referencing remains possible.
    </div>
"@
})
    
    <div class="nav" id="nav">
        <button class="nav-btn active" onclick="showSection('overview')">Overview</button>
        <button class="nav-btn" onclick="showSection('rightsizing')">Right-Sizing</button>
        <button class="nav-btn" onclick="showSection('health')">Host Health</button>
        <button class="nav-btn" onclick="showSection('network')">Network & Connectivity</button>
        <button class="nav-btn" onclick="showSection('storage')">Storage & Disks</button>
        <button class="nav-btn" onclick="showSection('resiliency')">Zone Resiliency</button>
        <button class="nav-btn" onclick="showSection('images')">Images</button>
        <button class="nav-btn" onclick="showSection('costs')">Cost Breakdown</button>
        <button class="nav-btn" onclick="showSection('priorities')">Priority Matrix</button>
$(if ($IncludeReservationAnalysis -and (SafeCount $reservationAnalysis) -gt 0) { '        <button class="nav-btn" onclick="showSection(''reservations'')">Reservations</button>' })
$(if ($hasKqlData) { '        <button class="nav-btn" onclick="showSection(''kql'')">Connection &amp; Logins</button>' })
$(if ($concurrencyData.Count -gt 0 -or $scalingPlanSchedules.Count -gt 0) { '        <button class="nav-btn" onclick="showSection(''scaling'')">Scaling &amp; Autoscale</button>' })
$(if ($IncludeAzureAdvisor -and (SafeCount $advisorRecommendations) -gt 0) { '        <button class="nav-btn" onclick="showSection(''advisor'')">Azure Advisor</button>' })
$(if ($w365Analysis.Count -gt 0) { '        <button class="nav-btn" onclick="showSection(''w365'')">W365 Readiness</button>' })
$(if ($securityPosture.Count -gt 0) { '        <button class="nav-btn" onclick="showSection(''security'')">Security Posture</button>' })
$(if ($uxScores.Count -gt 0) { '        <button class="nav-btn" onclick="showSection(''ux'')">UX Scores</button>' })
$(if ($orphanedResources.Count -gt 0) { '        <button class="nav-btn" onclick="showSection(''orphans'')">Orphaned Resources</button>' })
$(if ($summary.IncidentWindowAnalysis.Enabled) { '        <button class="nav-btn" onclick="showSection(''incident'')">Incident Analysis</button>' })
    </div>

    <!-- ========== OVERVIEW ========== -->
    <div class="section active" id="sec-overview">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Total VMs</div>
                <div class="card-value blue">$($summary.VMs)</div>
                <div class="card-sub">$($summary.SessionHosts) session hosts</div>
            </div>
            <div class="card">
                <div class="card-label">Host Pools</div>
                <div class="card-value blue">$($summary.HostPools)</div>
                <div class="card-sub">$($summary.ScalingPlans) scaling plans</div>
            </div>
            <div class="card">
                <div class="card-label">Est. Monthly Savings</div>
                <div class="card-value green">~`$$($summary.RightSizingAnalysis.PotentialMonthlySavings)</div>
                <div class="card-sub">$($summary.RightSizingAnalysis.SavingsPercentage) of current spend</div>
            </div>
            <div class="card">
                <div class="card-label">Resiliency Score</div>
                <div class="card-value $(if ($summary.ZoneResiliencyAnalysis.AverageResiliencyScore -ge 75) { 'green' } elseif ($summary.ZoneResiliencyAnalysis.AverageResiliencyScore -ge 40) { 'yellow' } else { 'red' })">$($summary.ZoneResiliencyAnalysis.AverageResiliencyScore)/100</div>
                <div class="card-sub">$($summary.ZoneResiliencyAnalysis.NonZoneVMs) VMs not zone-enabled</div>
            </div>
$(if ($uxScores.Count -gt 0) { @"
            <div class="card">
                <div class="card-label">Avg UX Score</div>
                <div class="card-value $(if ($summary.UserExperience.AvgUXScore -ge 75) { 'green' } elseif ($summary.UserExperience.AvgUXScore -ge 60) { 'yellow' } else { 'red' })">$($summary.UserExperience.AvgUXScore)/100</div>
                <div class="card-sub">$($summary.UserExperience.PoolsBelowGradeC) pool(s) below C</div>
            </div>
"@ })
$(if ($securityPosture.Count -gt 0) { @"
            <div class="card">
                <div class="card-label">Avg Security Score</div>
                <div class="card-value $(if ($summary.SecurityPosture.AvgSecurityScore -ge 75) { 'green' } elseif ($summary.SecurityPosture.AvgSecurityScore -ge 60) { 'yellow' } else { 'red' })">$($summary.SecurityPosture.AvgSecurityScore)/100</div>
                <div class="card-sub">$($summary.SecurityPosture.PoolsBelowGradeC) pool(s) below C</div>
            </div>
"@ })
        </div>
        
        <!-- Quick Findings -->
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Downsize Candidates</div>
                <div class="card-value $(if ($downsizeCandidates -gt 0) { 'yellow' } else { 'green' })">$downsizeCandidates</div>
            </div>
            <div class="card">
                <div class="card-label">Upsize Candidates</div>
                <div class="card-value $(if ($upsizeCandidates -gt 0) { 'red' } else { 'green' })">$upsizeCandidates</div>
            </div>
            <div class="card">
                <div class="card-label">Hosts With Issues</div>
                <div class="card-value $(if ($unavailableHosts -gt 0 -or $stuckHosts -gt 0) { 'red' } else { 'green' })">$($unavailableHosts + $stuckHosts)</div>
                <div class="card-sub">$stuckHosts stuck in drain</div>
            </div>
            <div class="card">
                <div class="card-label">AccelNet Gaps</div>
                <div class="card-value $(if ($eligibleNotEnabled -gt 0) { 'orange' } else { 'green' })">$eligibleNotEnabled</div>
                <div class="card-sub">eligible but not enabled</div>
            </div>
$(if ($orphanedResources.Count -gt 0) { @"
            <div class="card">
                <div class="card-label">Orphaned Resources</div>
                <div class="card-value orange">$($orphanedResources.Count)</div>
                <div class="card-sub">~`$$orphanedWaste/mo wasted</div>
            </div>
"@ })
        </div>
        
        <!-- Executive Summary -->
"@

    # Build executive summary prose
    $execLines = [System.Collections.Generic.List[string]]::new()

    # Environment overview
    $personalPools = @($hostPools | Where-Object { $_.HostPoolType -eq "Personal" })
    $pooledPools = @($hostPools | Where-Object { $_.HostPoolType -eq "Pooled" })
    $remoteAppPools = @($hostPools | Where-Object { $_.PreferredAppGroupType -eq "RailApplications" })
    $desktopPooledPools = @($pooledPools | Where-Object { $_.PreferredAppGroupType -ne "RailApplications" })
    $totalRunning = @($vms | Where-Object { $_.PowerState -eq "VM running" }).Count
    $totalDealloc = @($vms | Where-Object { $_.PowerState -ne "VM running" }).Count
    $runPctOverall = if ($vms.Count -gt 0) { [math]::Round($totalRunning / $vms.Count * 100, 0) } else { 0 }

    $envParts = @()
    if ($personalPools.Count -gt 0) { $envParts += "$($personalPools.Count) personal desktop" }
    if ($desktopPooledPools.Count -gt 0) { $envParts += "$($desktopPooledPools.Count) pooled desktop" }
    if ($remoteAppPools.Count -gt 0) { $envParts += "$($remoteAppPools.Count) RemoteApp" }
    $poolBreakdown = $envParts -join ", "

    $execLines.Add("This environment spans <strong>$($summary.HostPools) host pools</strong> ($poolBreakdown) with <strong>$($summary.VMs) session host VMs</strong> across $($summary.Subscriptions) subscription(s). At time of collection, $totalRunning VMs ($runPctOverall%) were running and $totalDealloc were deallocated.")

    # Top findings (pick the 3 most impactful)
    $topFindings = [System.Collections.Generic.List[string]]::new()

    # Cost findings
    if ($summary.RightSizingAnalysis.PotentialMonthlySavings -and [int]$summary.RightSizingAnalysis.PotentialMonthlySavings -gt 0) {
      $topFindings.Add("<strong>`$$($summary.RightSizingAnalysis.PotentialMonthlySavings)/mo</strong> in right-sizing savings identified across $downsizeCandidates downsize candidates")
    }
    if ($orphanedResources.Count -gt 0 -and $orphanedWaste -gt 0) {
      $topFindings.Add("<strong>`$$orphanedWaste/mo</strong> in orphaned resources ($($orphanedResources.Count) unattached disks/NICs/IPs)")
    }

    # Security
    $lowSecCount = @($securityPosture | Where-Object { $_.SecurityScore -lt 60 }).Count
    if ($lowSecCount -gt 0) {
      $topFindings.Add("<strong>$lowSecCount pool(s)</strong> scored below C on security posture (missing Trusted Launch, Secure Boot, or vTPM)")
    }

    # UX
    $poorUxCount = @($uxScores | Where-Object { $_.UXScore -lt 60 }).Count
    if ($poorUxCount -gt 0) {
      $topFindings.Add("<strong>$poorUxCount pool(s)</strong> have poor user experience scores driven by profile load times, RTT, or disconnect rates")
    }

    # Scaling
    $execPoolsWithPlans = @{}
    foreach ($a in $scalingPlanAssignments) {
      $pn = if ($a.HostPoolName) { $a.HostPoolName } elseif ($a.HostPoolArmId) { ($a.HostPoolArmId -split '/')[-1] } else { "" }
      if ($pn) { $execPoolsWithPlans[$pn] = $true }
    }
    $execAllPools = @($vms | Select-Object -ExpandProperty HostPoolName -Unique)
    $execNoPlanPools = @($execAllPools | Where-Object { -not $execPoolsWithPlans.ContainsKey($_) })
    $execNoPlanRunning = 0
    foreach ($np in $execNoPlanPools) {
      $execNoPlanRunning += @($vms | Where-Object { $_.HostPoolName -eq $np -and $_.PowerState -eq "VM running" }).Count
    }
    $noScalingCount = $execNoPlanPools.Count
    $noPlanRunning = $execNoPlanRunning
    if ($noScalingCount -gt 0 -and $noPlanRunning -gt 0) {
      $topFindings.Add("<strong>$noScalingCount pool(s)</strong> have no scaling plan with $noPlanRunning VMs running unmanaged")
    }

    # Hosts with issues
    if ($unavailableHosts -gt 0 -or $stuckHosts -gt 0) {
      $topFindings.Add("<strong>$($unavailableHosts + $stuckHosts) session host(s)</strong> need attention ($stuckHosts stuck in drain, $unavailableHosts unavailable)")
    }

    # Upsizing
    if ($upsizeCandidates -gt 0) {
      $topFindings.Add("<strong>$upsizeCandidates VM(s)</strong> are under-provisioned and recommended for upsizing")
    }

    # Session duration licensing signal
    $execSessionDur = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_SessionDuration" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "UserName" -and $_.AvgDuration })
    if ($execSessionDur.Count -gt 0) {
      $execShortUsers = @($execSessionDur | Where-Object { [double]$_.AvgDuration -lt 120 }).Count
      $execShortPct = [math]::Round($execShortUsers / $execSessionDur.Count * 100, 0)
      if ($execShortPct -ge 40) {
        $topFindings.Add("<strong>$execShortPct% of users</strong> average under 2-hour sessions — strong Frontline/shared licensing opportunity")
      }
    }

    if ($topFindings.Count -gt 0) {
      $findingsLimit = [math]::Min($topFindings.Count, 4)
      $execLines.Add("<strong>Key findings:</strong> " + (($topFindings | Select-Object -First $findingsLimit | ForEach-Object { $_ }) -join " · ") + ".")
    }

    # Quick wins — count from known quick-win patterns already computed
    $quickWinCount = 0
    if ($orphanedResources.Count -gt 0) { $quickWinCount++ }
    if ($noScalingCount -gt 0 -and $noPlanRunning -gt 0) { $quickWinCount++ }
    $disabledPlansExec = @($scalingPlanAssignments | Where-Object { $_.IsEnabled -eq $false -or $_.IsEnabled -eq "False" })
    if ($disabledPlansExec.Count -gt 0) { $quickWinCount++ }
    $highRampDownExec = @($scalingPlanSchedules | Where-Object { $_.ScheduleName -eq "Weekdays" -and [int]$_.RampDownCapacity -gt 70 })
    if ($highRampDownExec.Count -gt 2) { $quickWinCount++ }
    if ($eligibleNotEnabled -gt 0) { $quickWinCount++ }
    if ($quickWinCount -gt 0) {
      $execLines.Add("At least <strong>$quickWinCount quick win(s)</strong> identified — low-effort actions like removing orphaned resources, enabling scaling plans, or turning on Accelerated Networking. See the Priority Matrix tab for the full list.")
    }

    $htmlReport += @"
        <div style="background:linear-gradient(135deg,#f8fafc,#eef2f7);border-left:4px solid #0078d4;border-radius:4px;padding:20px 24px;margin-bottom:24px">
            <h3 style="margin:0 0 12px 0;font-size:15px;color:#1a1a2e">📋 Executive Summary</h3>
            <div style="font-size:13px;line-height:1.7;color:#333">
                $(($execLines | ForEach-Object { "<p style='margin:0 0 8px 0'>$_</p>" }) -join "`n                ")
            </div>
        </div>
        
        <!-- CPU Distribution Chart -->
        <div class="two-col">
            <div class="table-wrap">
                <div class="table-title">📊 CPU Utilization Distribution</div>
                <div style="padding:20px;">
                    $(foreach ($bucket in @("0-20","20-40","40-60","60-80","80-100")) {
                      $count = $cpuBuckets[$bucket]
                      $pct = if ((SafeCount $vmRightSizing) -gt 0) { [math]::Round($count / (SafeCount $vmRightSizing) * 100, 0) } else { 0 }
                      $width = [math]::Round($count / $cpuBucketMax * 100, 0)
                      $color = switch ($bucket) {
                        "0-20"   { "#28a745" }
                        "20-40"  { "#5cb85c" }
                        "40-60"  { "#ffc107" }
                        "60-80"  { "#fd7e14" }
                        "80-100" { "#dc3545" }
                      }
                      "<div style='display:flex;align-items:center;gap:12px;margin-bottom:10px;'>" +
                      "<span style='width:55px;font-size:13px;font-weight:500;text-align:right;color:#555'>$bucket%</span>" +
                      "<div style='flex:1;background:#f0f0f0;border-radius:6px;height:28px;overflow:hidden'>" +
                      "<div style='width:${width}%;height:100%;background:$color;border-radius:6px;min-width:$(if ($count -gt 0) {'2px'} else {'0'})'></div>" +
                      "</div>" +
                      "<span style='width:60px;font-size:13px;color:#555'>$count VMs</span>" +
                      "<span style='width:35px;font-size:11px;color:#999'>$pct%</span></div>"
                    })
                </div>
            </div>
            <div class="table-wrap">
                <div class="table-title">📋 Quick Assessment</div>
                <div style="padding:20px;font-size:14px;line-height:2;">
                    $(
                      $overProvPct = if ((SafeCount $vmRightSizing) -gt 0) { [math]::Round(($cpuBuckets["0-20"] + $cpuBuckets["20-40"]) / (SafeCount $vmRightSizing) * 100, 0) } else { 0 }
                      $stressedPct = if ((SafeCount $vmRightSizing) -gt 0) { [math]::Round($cpuBuckets["80-100"] / (SafeCount $vmRightSizing) * 100, 0) } else { 0 }
                      $lines = @()
                      if ($overProvPct -ge 50) { $lines += "<div>🔵 <strong>$overProvPct%</strong> of VMs under 40% avg CPU — fleet is over-provisioned</div>" }
                      elseif ($overProvPct -ge 30) { $lines += "<div>🟢 <strong>$overProvPct%</strong> of VMs under 40% avg CPU — some optimization headroom</div>" }
                      if ($stressedPct -gt 0) { $lines += "<div>🔴 <strong>$stressedPct%</strong> of VMs above 80% avg CPU — investigate capacity</div>" }
                      if ($stuckHosts -gt 0) { $lines += "<div>🔴 <strong>$stuckHosts</strong> host(s) stuck in drain — immediate attention needed</div>" }
                      if ($eligibleNotEnabled -gt 0) { $lines += "<div>🟡 <strong>$eligibleNotEnabled</strong> VM(s) can enable AccelNet — quick win for latency</div>" }
                      if ($multiVersionImages -gt 0) { $lines += "<div>🟡 <strong>$multiVersionImages</strong> image group(s) have version drift</div>" }
                      if ($lines.Count -eq 0) { $lines += "<div>🟢 Environment looks healthy across all checks</div>" }
                      $lines -join ""
                    )
                </div>
            </div>
        </div>
        
        <!-- Top Recommendations -->
        $(if (@($summary.TopRecommendations).Count -gt 0) {
          $recsHtml = "<div class='table-wrap'><div class='table-title'>🎯 Top Recommendations</div><div style='padding:16px;'>"
          foreach ($rec in $summary.TopRecommendations) {
            $icon = if ($rec -match 'upsize|critical|overloaded|stuck|issues') { '🔴' } 
                    elseif ($rec -match 'downsize|savings|Premium|ephemeral|AccelNet|drift') { '🟡' }
                    else { '🔵' }
            $recsHtml += "<div style='padding:8px 0; border-bottom:1px solid #f0f0f0;'>$icon $rec</div>"
          }
          $recsHtml += "</div></div>"
          $recsHtml
        })
    </div>

    <!-- ========== RIGHT-SIZING ========== -->
    <div class="section" id="sec-rightsizing">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">VMs Analyzed</div>
                <div class="card-value blue">$($summary.RightSizingAnalysis.TotalVMsAnalyzed)</div>
            </div>
            <div class="card">
                <div class="card-label">Downsize</div>
                <div class="card-value yellow">$downsizeCandidates</div>
            </div>
            <div class="card">
                <div class="card-label">Upsize</div>
                <div class="card-value red">$upsizeCandidates</div>
            </div>
            <div class="card">
                <div class="card-label">Appropriately Sized</div>
                <div class="card-value green">$appropriatelySized</div>
            </div>
        </div>
        
        <div class="toolbar">
            <input type="text" class="search-input" id="rs-search" placeholder="Search VMs..." onkeyup="filterTable('rs-table', this.value)">
            <button class="filter-btn active" onclick="filterRs(this, 'all')">All</button>
            <button class="filter-btn" onclick="filterRs(this, 'upsize')">Upsize</button>
            <button class="filter-btn" onclick="filterRs(this, 'downsize')">Downsize</button>
            <button class="filter-btn" onclick="filterRs(this, 'keep')">Keep Current</button>
        </div>
        
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="rs-table">
                    <thead><tr>
                        <th onclick="sortTable('rs-table',0)">VM Name</th>
                        <th onclick="sortTable('rs-table',1)">Current</th>
                        <th onclick="sortTable('rs-table',2)">Recommended</th>
                        <th onclick="sortTable('rs-table',3)">Avg CPU</th>
                        <th onclick="sortTable('rs-table',4)">Peak CPU</th>
                        <th onclick="sortTable('rs-table',5)">Peak Sessions</th>
                        <th onclick="sortTable('rs-table',6)">Cost Impact</th>
                        <th onclick="sortTable('rs-table',7)">Evidence</th>
                    </tr></thead>
                    <tbody>
"@

  # Build right-sizing table rows
  $sortedRs = $vmRightSizing | Sort-Object { 
    if ($_.RecommendedSize -ne "Keep Current" -and $_.RecommendedSize -ne "Unknown") { 
      if ($_.EstimatedMonthlySavings -ne "N/A" -and [double]$_.EstimatedMonthlySavings -le 0) { 0 }  # Upsize first
      else { 1 }  # Downsize second
    } else { 2 }  # Keep current last
  }, { if ($_.EstimatedMonthlySavings -ne "N/A") { [math]::Abs([double]$_.EstimatedMonthlySavings) } else { 0 } } -Descending

  foreach ($rec in $sortedRs) {
    $action = if ($rec.RecommendedSize -eq "Keep Current" -or $rec.RecommendedSize -eq "Unknown") { "keep" }
              elseif ($rec.EstimatedMonthlySavings -ne "N/A" -and [double]$rec.EstimatedMonthlySavings -gt 0) { "downsize" }
              else { "upsize" }
    
    $actionBadge = switch ($action) {
      "upsize" { "<span class='badge b-red'>UPSIZE</span>" }
      "downsize" { "<span class='badge b-green'>DOWNSIZE</span>" }
      "keep" { "<span class='badge b-gray'>KEEP</span>" }
    }
    
    $costDisplay = if ($rec.EstimatedMonthlySavings -eq "N/A") { "<span style='color:#999'>—</span>" }
                   elseif ([double]$rec.EstimatedMonthlySavings -gt 0) { "<span style='color:#2e7d32;font-weight:600'>-`$$($rec.EstimatedMonthlySavings)/mo</span>" }
                   elseif ([double]$rec.EstimatedMonthlySavings -lt 0) { "<span style='color:#c62828;font-weight:600'>+`$$([math]::Abs([double]$rec.EstimatedMonthlySavings))/mo</span>" }
                   else { "<span style='color:#999'>—</span>" }
    
    $evScore = if ($rec.EvidenceScore) { $rec.EvidenceScore } else { 0 }
    $evColor = if ($evScore -ge 85) { "#28a745" } elseif ($evScore -ge 50) { "#ffc107" } else { "#dc3545" }
    $evSignals = if ($rec.EvidenceSignals) { $rec.EvidenceSignals } else { "None" }
    
    $recDisplay = if ($rec.RecommendedSize -eq "Keep Current") { $rec.RecommendedSize }
                  elseif ($rec.RecommendedSize -eq "Unknown") { $rec.RecommendedSize }
                  else { "$actionBadge $($rec.RecommendedSize)" }

    $reasonEscaped = ($rec.Reason -replace '"','&quot;' -replace "'","&#39;") -replace ';','<br>'
    $workloadBadge = if ($rec.AppGroupType -eq "RemoteApp") { "<span class='badge' style='background:#e3f2fd;color:#1565c0;font-size:9px'>RemoteApp</span>" } elseif ($rec.HostPoolType -eq "Personal") { "<span class='badge' style='background:#fce4ec;color:#c62828;font-size:9px'>Personal</span>" } else { "" }
    $htmlReport += @"
                    <tr data-action="$action" onclick="toggleDetail(this)" style="cursor:pointer">
                        <td><strong>$(Scrub-VMName $rec.VMName)</strong><br><span style="font-size:11px;color:#999">$(Scrub-HostPoolName $rec.HostPoolName)</span> $workloadBadge</td>
                        <td>$($rec.CurrentSize)</td>
                        <td>$recDisplay</td>
                        <td>$($rec.AvgCPU)%</td>
                        <td style="$(if ($rec.PeakCPU -gt 80) { 'color:#c62828;font-weight:600' })">$($rec.PeakCPU)%</td>
                        <td>$(if ($rec.PeakSessions) { $rec.PeakSessions } else { '—' })</td>
                        <td>$costDisplay</td>
                        <td title="$evSignals"><div class="evidence-bar"><div class="evidence-fill"><div class="evidence-fill-inner" style="width:${evScore}%;background:$evColor"></div></div><span style="font-size:11px">$evScore</span></div></td>
                    </tr>
                    <tr class="detail-row" style="display:none">
                        <td colspan="8" style="background:#f8fafc;padding:12px 20px;font-size:12px;border-left:3px solid $evColor">
                            <div style="display:flex;gap:24px;flex-wrap:wrap">
                                <div style="flex:1;min-width:300px"><strong>Reason:</strong><br>$reasonEscaped</div>
                                <div><strong>Evidence Signals:</strong><br>$evSignals<br><br><strong>Sessions/vCPU:</strong> $(if ($rec.SessionsPerVCPU) { $rec.SessionsPerVCPU } else { '—' }) &nbsp; <strong>Memory/User:</strong> $(if ($rec.MemoryPerSessionGB) { "$($rec.MemoryPerSessionGB) GB" } else { '—' })</div>
                            </div>
                        </td>
                    </tr>
"@
  }

  $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
        <div class="alert alert-info">
            <strong>💡 Tip:</strong> Click any row to expand and see the full recommendation reason, evidence signals, and session density metrics. Evidence Score: 0-100 based on available data (CPU=20, Memory=15, Sessions=30, Lookback≥7d=15, Pattern=10, Known SKU=10).
        </div>

        <!-- SKU Diversity & Allocation Resilience -->
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">🎲 SKU Diversity &amp; Allocation Resilience</h3>
        <div class="alert alert-info">
            Azure regional capacity constraints can prevent VM allocation for specific SKU families. Diversifying SKUs across families (e.g., D-series + E-series) ensures that an allocation failure on one family doesn't prevent scaling your entire host pool.
        </div>
$(if ($skuDiversityAnalysis.Count -gt 0) {
  $skuCards = @"
        <div class="card-grid">
            <div class="card">
                <div class="card-label">High Risk Pools</div>
                <div class="card-value $(if ($highRiskPools.Count -gt 0) { 'red' } else { 'green' })">$($highRiskPools.Count)</div>
                <div class="card-sub">single SKU family</div>
            </div>
            <div class="card">
                <div class="card-label">Medium Risk Pools</div>
                <div class="card-value $(if ($medRiskPools.Count -gt 0) { 'yellow' } else { 'green' })">$($medRiskPools.Count)</div>
                <div class="card-sub">limited diversity</div>
            </div>
            <div class="card">
                <div class="card-label">Single Region</div>
                <div class="card-value $(if ($singleRegionPools.Count -gt 0) { 'yellow' } else { 'green' })">$($singleRegionPools.Count)</div>
                <div class="card-sub">no geo-redundancy</div>
            </div>
        </div>
"@
  $skuCards

  $skuTableHtml = @"
        <div class="table-wrap">
            <div class="table-title">Host Pool Allocation Risk Assessment</div>
            <div class="table-scroll">
                <table id="sku-table">
                    <thead><tr>
                        <th onclick="sortTable('sku-table',0)">Host Pool</th>
                        <th onclick="sortTable('sku-table',1)">VMs</th>
                        <th onclick="sortTable('sku-table',2)">SKU Families</th>
                        <th onclick="sortTable('sku-table',3)">Dominant SKU</th>
                        <th onclick="sortTable('sku-table',4)">Concentration</th>
                        <th onclick="sortTable('sku-table',5)">Regions</th>
                        <th onclick="sortTable('sku-table',6)">Zones</th>
                        <th onclick="sortTable('sku-table',7)">SKU Risk</th>
                        <th onclick="sortTable('sku-table',8)">Region Risk</th>
                        <th>Recommendation</th>
                    </tr></thead>
                    <tbody>
"@
  $skuTableHtml
  foreach ($sd in ($skuDiversityAnalysis | Sort-Object { switch ($_.OverallRisk) { "High" { 0 } "Medium" { 1 } default { 2 } } })) {
    $skuRiskColor = switch ($sd.SkuRisk) { "High" { "color:#c62828;font-weight:700" } "Medium" { "color:#e65100;font-weight:600" } default { "color:#2e7d32" } }
    $regRiskColor = switch ($sd.RegionRisk) { "High" { "color:#c62828;font-weight:700" } "Medium" { "color:#e65100;font-weight:600" } default { "color:#2e7d32" } }
    $concColor = if ($sd.DominantSkuPct -ge 90) { "color:#c62828;font-weight:600" } elseif ($sd.DominantSkuPct -ge 70) { "color:#e65100" } else { "" }
    $recs = if ($sd.Recommendations) { $sd.Recommendations } else { "No action needed" }
    "                    <tr>"
    "                        <td><strong>$(Scrub-HostPoolName $sd.HostPoolName)</strong><br><span style='font-size:11px;color:#888'>$($sd.HostPoolType)</span></td>"
    "                        <td>$($sd.VMCount)</td>"
    "                        <td>$($sd.FamilyList)<br><span style='font-size:11px;color:#888'>$($sd.SeriesList)</span></td>"
    "                        <td>$($sd.DominantSku)</td>"
    "                        <td style='$concColor'>$($sd.DominantSkuPct)%</td>"
    "                        <td>$($sd.RegionList)</td>"
    "                        <td>$(if ($sd.ZoneList -eq 'No zones') { '<span style=''color:#999''>No zones</span>' } else { $sd.ZoneList })</td>"
    "                        <td style='$skuRiskColor'>$($sd.SkuRisk)</td>"
    "                        <td style='$regRiskColor'>$($sd.RegionRisk)</td>"
    "                        <td style='font-size:12px;max-width:350px'>$recs</td>"
    "                    </tr>"
  }
  "                    </tbody></table></div></div>"
} else {
  "        <div style='color:#999;padding:20px'>No host pool data available for SKU diversity analysis.</div>"
})
    </div>

    <!-- ========== HOST HEALTH ========== -->
    <div class="section" id="sec-health">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Healthy</div>
                <div class="card-value green">$($summary.SessionHostHealth.Healthy)</div>
                <div class="card-sub">of $($summary.SessionHostHealth.TotalHosts) total</div>
            </div>
            <div class="card">
                <div class="card-label">Drain Mode</div>
                <div class="card-value $(if ($drainedHosts -gt 0) { 'yellow' } else { 'green' })">$drainedHosts</div>
            </div>
            <div class="card">
                <div class="card-label">Stuck in Drain</div>
                <div class="card-value $(if ($stuckHosts -gt 0) { 'red' } else { 'green' })">$stuckHosts</div>
            </div>
            <div class="card">
                <div class="card-label">Unavailable</div>
                <div class="card-value $(if ($summary.SessionHostHealth.Unavailable -gt 0) { 'red' } else { 'green' })">$($summary.SessionHostHealth.Unavailable)</div>
            </div>
        </div>
        
$(if ($stuckHosts -gt 0) {
  "<div class='alert alert-danger'><strong>Action Required:</strong> $stuckHosts session host(s) appear stuck in drain mode with stale heartbeats. These hosts are not serving users and may need manual recovery or reimaging.</div>"
})
        
        <div class="table-wrap">
            <div class="table-title">Session Host Status</div>
            <div class="table-scroll">
                <table id="health-table">
                    <thead><tr>
                        <th onclick="sortTable('health-table',0)">Session Host</th>
                        <th onclick="sortTable('health-table',1)">Host Pool</th>
                        <th onclick="sortTable('health-table',2)">Health</th>
                        <th onclick="sortTable('health-table',3)">AVD Status</th>
                        <th onclick="sortTable('health-table',4)">Sessions</th>
                        <th onclick="sortTable('health-table',5)">Drain Mode</th>
                        <th onclick="sortTable('health-table',6)">Heartbeat Age</th>
                        <th onclick="sortTable('health-table',7)">Finding</th>
                        <th>Remediation</th>
                    </tr></thead>
                    <tbody>
"@

  # Only show hosts with issues first, then healthy
  $healthSorted = $sessionHostHealth | Sort-Object { switch ($_.Severity) { "High" { 0 } "Medium" { 1 } default { 2 } } }, SessionHostName
  foreach ($sh in $healthSorted) {
    $sevBadge = switch ($sh.Severity) {
      "High"   { "<span class='badge b-red'>Critical</span>" }
      "Medium" { "<span class='badge b-yellow'>Warning</span>" }
      default  { "<span class='badge b-green'>Healthy</span>" }
    }
    $drainDisplay = if ($sh.DrainMode) { "<span class='badge b-yellow'>Yes</span>" } else { "No" }
    $hbDisplay = if ($sh.HeartbeatAgeHrs) { "$([math]::Round($sh.HeartbeatAgeHrs, 1))h" } else { "—" }
    
    $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-VMName $sh.SessionHostName)</strong></td>
                        <td>$(Scrub-HostPoolName $sh.HostPoolName)</td>
                        <td>$sevBadge</td>
                        <td>$($sh.Status)</td>
                        <td>$($sh.ActiveSessions)</td>
                        <td>$drainDisplay</td>
                        <td>$hbDisplay</td>
                        <td style="font-size:12px">$($sh.Finding)</td>
                        <td style="font-size:11px;max-width:400px;color:#555">$(if ($sh.Remediation) { $sh.Remediation } else { "—" })</td>
                    </tr>
"@
  }

  $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- ========== NETWORK ========== -->
    <div class="section" id="sec-network">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">AccelNet Enabled</div>
                <div class="card-value green">$($summary.AcceleratedNetworking.Enabled)</div>
            </div>
            <div class="card">
                <div class="card-label">AccelNet Gaps</div>
                <div class="card-value $(if ($eligibleNotEnabled -gt 0) { 'orange' } else { 'green' })">$eligibleNotEnabled</div>
                <div class="card-sub">eligible but not enabled</div>
            </div>
            <div class="card">
                <div class="card-label">Not Eligible</div>
                <div class="card-value">$($summary.AcceleratedNetworking.NotEligible)</div>
                <div class="card-sub">&lt; 4 vCPU</div>
            </div>
        </div>
        
$(if ($eligibleNotEnabled -gt 0) {
  "<div class='alert alert-warning'><strong>Quick Win:</strong> $eligibleNotEnabled VM(s) support Accelerated Networking but don't have it enabled. Enabling AccelNet reduces network latency and jitter — directly improving RDP session quality. Requires a VM restart.</div>"
})
        
        <div class="table-wrap">
            <div class="table-title">Accelerated Networking Status</div>
            <div class="table-scroll">
                <table id="accel-table">
                    <thead><tr>
                        <th onclick="sortTable('accel-table',0)">VM Name</th>
                        <th onclick="sortTable('accel-table',1)">VM Size</th>
                        <th onclick="sortTable('accel-table',2)">vCPUs</th>
                        <th onclick="sortTable('accel-table',3)">AccelNet</th>
                        <th onclick="sortTable('accel-table',4)">Finding</th>
                    </tr></thead>
                    <tbody>
"@

  # Show gaps first
  $accelSorted = $accelNetFindings | Sort-Object { if ($_.Eligible -and -not $_.AccelNetEnabled) { 0 } else { 1 } }, VMName
  foreach ($an in $accelSorted) {
    $anBadge = if ($an.AccelNetEnabled) { "<span class='badge b-green'>Enabled</span>" }
               elseif ($an.Eligible) { "<span class='badge b-orange'>Not Enabled</span>" }
               else { "<span class='badge b-gray'>N/A</span>" }
    
    $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-VMName $an.VMName)</strong></td>
                        <td>$($an.VMSize)</td>
                        <td>$($an.vCPUs)</td>
                        <td>$anBadge</td>
                        <td style="font-size:12px">$($an.Finding)</td>
                    </tr>
"@
  }

  $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- Network Readiness Assessment (v4.0.0) -->
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">🔌 Network Readiness Assessment</h3>
"@

  # RDP Shortpath card
  $spBadge = if ($shortpathSummary.ShortpathPct -ge 80) { "<span class='badge b-green'>Healthy</span>" }
             elseif ($shortpathEnabled) { "<span class='badge b-orange'>Partial</span>" }
             else { "<span class='badge b-red'>Not Enabled</span>" }
  
  $htmlReport += @"
        <div class="card-grid">
            <div class="card">
                <div class="card-label">RDP Shortpath (UDP)</div>
                <div class="card-value $(if ($shortpathSummary.ShortpathPct -ge 80) { 'green' } elseif ($shortpathEnabled) { 'yellow' } else { 'red' })">$($shortpathSummary.ShortpathPct)%</div>
                <div class="card-sub">$($shortpathSummary.UdpConnections) / $($shortpathSummary.TotalConnections) connections</div>
            </div>
            <div class="card">
                <div class="card-label">Subnets Analyzed</div>
                <div class="card-value blue">$($subnetAnalysis.Count)</div>
                <div class="card-sub">$(($subnetAnalysis | Where-Object { $_.UsagePct -ge 70 } | Measure-Object).Count) at >70% capacity</div>
            </div>
            <div class="card">
                <div class="card-label">Private Endpoints</div>
                <div class="card-value $(if ($hpWithoutPE.Count -gt 0) { 'yellow' } else { 'green' })">$(($privateEndpointFindings | Where-Object { $_.HasPrivateEndpoint -eq $true } | Measure-Object).Count) / $(($privateEndpointFindings | Measure-Object).Count)</div>
                <div class="card-sub">host pools with PE</div>
            </div>
            <div class="card">
                <div class="card-label">Network Issues</div>
                <div class="card-value $(if ($criticalNetFindings.Count -gt 0) { 'red' } elseif ($warningNetFindings.Count -gt 0) { 'yellow' } else { 'green' })">$($criticalNetFindings.Count + $warningNetFindings.Count)</div>
                <div class="card-sub">$($criticalNetFindings.Count) critical, $($warningNetFindings.Count) warnings</div>
            </div>
        </div>
"@

  # Network findings table
  if ($networkFindings.Count -gt 0) {
    $htmlReport += @"
        <div class="table-wrap">
            <div class="table-title">Network Readiness Checks</div>
            <div class="table-scroll">
                <table id="netready-table">
                    <thead><tr>
                        <th onclick="sortTable('netready-table',0)">Check</th>
                        <th onclick="sortTable('netready-table',1)">Status</th>
                        <th>Detail</th>
                        <th onclick="sortTable('netready-table',3)">Impact</th>
                        <th>Recommendation</th>
                    </tr></thead>
                    <tbody>
"@
    foreach ($nf in ($networkFindings | Sort-Object { switch ($_.Impact) { "High" { 0 } "Medium" { 1 } default { 2 } } })) {
      $impactColor = switch ($nf.Impact) { "High" { "color:#c62828;font-weight:700" } "Medium" { "color:#e65100;font-weight:600" } default { "color:#2e7d32" } }
      $statusBadge = switch -Wildcard ($nf.Status) {
        "Critical"      { "<span class='badge b-red'>$($nf.Status)</span>" }
        "Not*"          { "<span class='badge b-red'>$($nf.Status)</span>" }
        "Missing"       { "<span class='badge b-red'>$($nf.Status)</span>" }
        "Disconnected"  { "<span class='badge b-red'>$($nf.Status)</span>" }
        "Unprotected"   { "<span class='badge b-red'>$($nf.Status)</span>" }
        "Warning"       { "<span class='badge b-orange'>$($nf.Status)</span>" }
        "Partial"       { "<span class='badge b-orange'>$($nf.Status)</span>" }
        "Custom*"       { "<span class='badge b-blue'>$($nf.Status)</span>" }
        "Azure*"        { "<span class='badge b-blue'>$($nf.Status)</span>" }
        "Good"          { "<span class='badge b-green'>$($nf.Status)</span>" }
        default         { "<span class='badge b-gray'>$($nf.Status)</span>" }
      }
      $htmlReport += @"
                    <tr>
                        <td><strong>$($nf.Check)</strong></td>
                        <td>$statusBadge</td>
                        <td style="font-size:12px;max-width:350px">$($nf.Detail)</td>
                        <td style="$impactColor">$($nf.Impact)</td>
                        <td style="font-size:12px;max-width:400px">$($nf.Recommendation)</td>
                    </tr>
"@
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
  }

  # Subnet capacity table
  if ($subnetAnalysis.Count -gt 0) {
    $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">Subnet Capacity</h3>
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="subnet-table">
                    <thead><tr>
                        <th onclick="sortTable('subnet-table',0)">Subnet</th>
                        <th onclick="sortTable('subnet-table',1)">VNet</th>
                        <th onclick="sortTable('subnet-table',2)">CIDR</th>
                        <th onclick="sortTable('subnet-table',3)">Usable IPs</th>
                        <th onclick="sortTable('subnet-table',4)">Used</th>
                        <th onclick="sortTable('subnet-table',5)">Available</th>
                        <th onclick="sortTable('subnet-table',6)">Usage %</th>
                        <th>Session Hosts</th>
                        <th><span title="Network Security Group — controls inbound/outbound traffic rules for the subnet">NSG</span></th>
                        <th><span title="User Defined Route (Azure Route Table) — custom routing rules that override default Azure routing, often used to force traffic through a firewall or NVA">UDR</span></th>
                    </tr></thead>
                    <tbody>
"@
    foreach ($sn in ($subnetAnalysis | Sort-Object UsagePct -Descending)) {
      $usageColor = if ($sn.UsagePct -ge 90) { "color:#c62828;font-weight:700" } elseif ($sn.UsagePct -ge 70) { "color:#e65100;font-weight:600" } else { "" }
      $usageBar = "<div style='width:80px;background:#f0f0f0;border-radius:4px;height:14px;overflow:hidden;display:inline-block;vertical-align:middle'><div style='width:$($sn.UsagePct)%;height:100%;background:$(if ($sn.UsagePct -ge 90) { '#c62828' } elseif ($sn.UsagePct -ge 70) { '#e65100' } else { '#2e7d32' });border-radius:4px'></div></div>"
      $nsgBadge = if ($sn.HasNSG) { "<span class='badge b-green' title='NSG attached — network traffic rules are enforced'>✓</span>" } else { "<span class='badge b-red' title='No NSG — subnet has no network security rules'>✗</span>" }
      $udrBadge = if ($sn.HasRouteTable) { "<span class='badge b-blue' title='Custom route table attached — traffic may be routed through a firewall or NVA'>✓</span>" } else { "<span class='badge b-gray' title='Using Azure default routing — no custom routes'>—</span>" }
      
      $htmlReport += @"
                    <tr>
                        <td><strong>$($sn.SubnetName)</strong></td>
                        <td>$($sn.VNetName)</td>
                        <td>$($sn.AddressPrefix)</td>
                        <td>$($sn.UsableIPs)</td>
                        <td>$($sn.UsedIPs)</td>
                        <td style="$usageColor">$($sn.AvailableIPs)</td>
                        <td>$usageBar <span style="$usageColor;font-size:12px">$($sn.UsagePct)%</span></td>
                        <td>$($sn.SessionHostVMs)</td>
                        <td>$nsgBadge</td>
                        <td>$udrBadge</td>
                    </tr>
"@
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
  }

  $htmlReport += @"
        <div class="alert alert-info">
            <strong>Connection quality data:</strong> Detailed profile load times, RTT/bandwidth metrics, connection errors, and disconnect rates are available on the <strong>Connection &amp; Logins</strong> tab.
$(if ($kqlNoData.Count -gt 0) {
  "            <br><br><strong>⚠️ No data returned for:</strong> $($kqlNoData -join ', '). This typically means too few user sessions in the $MetricsLookbackDays-day window. Try running with <code>-MetricsLookbackDays 30</code> for more data."
})
$(if ($kqlFailed.Count -gt 0) {
  "            <br><br><strong>ℹ️ Tables not available:</strong> $($kqlFailed -join ', '). These diagnostic tables may not be enabled on the host pool. This is normal and does not affect other analysis."
})
        </div>
    </div>

    <!-- ========== STORAGE ========== -->
    <div class="section" id="sec-storage">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Premium on Pooled</div>
                <div class="card-value $(if ($premiumOnPooled -gt 0) { 'yellow' } else { 'green' })">$premiumOnPooled</div>
                <div class="card-sub">cost savings available</div>
            </div>
            <div class="card">
                <div class="card-label">Non-Ephemeral Pooled</div>
                <div class="card-value $(if ($nonEphemeral -gt 0) { 'yellow' } else { 'green' })">$nonEphemeral</div>
                <div class="card-sub">perf + cost improvement</div>
            </div>
            <div class="card">
                <div class="card-label">Standard HDD</div>
                <div class="card-value $(if ($summary.StorageOptimization.StandardHDD -gt 0) { 'red' } else { 'green' })">$($summary.StorageOptimization.StandardHDD)</div>
                <div class="card-sub">too slow for AVD</div>
            </div>
        </div>

$(if ($premiumOnPooled -gt 0 -or $nonEphemeral -gt 0) {
  "<div class='alert alert-warning'><strong>Storage Optimization:</strong> Pooled session hosts benefit from Standard SSD (cost) or Ephemeral OS disks (cost + performance). Premium SSD is typically only justified for personal desktops with persistent user data.</div>"
})
        
        <div class="table-wrap">
            <div class="table-title">OS Disk Configuration</div>
            <div class="table-scroll">
                <table id="storage-table">
                    <thead><tr>
                        <th onclick="sortTable('storage-table',0)">VM Name</th>
                        <th onclick="sortTable('storage-table',1)">Host Pool</th>
                        <th onclick="sortTable('storage-table',2)">Pool Type</th>
                        <th onclick="sortTable('storage-table',3)"><span title="Desktop or RemoteApp">Workload</span></th>
                        <th onclick="sortTable('storage-table',4)">OS Disk</th>
                        <th onclick="sortTable('storage-table',5)">Ephemeral</th>
                        <th onclick="sortTable('storage-table',6)">Findings</th>
                    </tr></thead>
                    <tbody>
"@

  $storageSorted = $storageFindingsList | Sort-Object { if ($_.Findings -ne "Optimal") { 0 } else { 1 } }, VMName
  foreach ($sf in $storageSorted) {
    $findBadge = if ($sf.Findings -eq "Optimal") { "<span class='badge b-green'>Optimal</span>" } else { "<span class='badge b-yellow'>Action</span>" }
    $ephBadge = if ($sf.OSDiskEphemeral) { "<span class='badge b-green'>Yes</span>" } else { "No" }
    
    $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-VMName $sf.VMName)</strong></td>
                        <td>$(Scrub-HostPoolName $sf.HostPoolName)</td>
                        <td>$($sf.HostPoolType)</td>
                        <td>$(if ($sf.AppGroupType -eq 'RemoteApp') { "<span class='badge' style='background:#e3f2fd;color:#1565c0;font-size:10px'>RemoteApp</span>" } elseif ($sf.AppGroupType -eq 'Desktop') { 'Desktop' } else { $sf.AppGroupType })</td>
                        <td>$($sf.OSDiskType)</td>
                        <td>$ephBadge</td>
                        <td>$findBadge <span style="font-size:12px">$($sf.Findings)</span></td>
                    </tr>
"@
  }

  $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- ========== RESILIENCY ========== -->
    <div class="section" id="sec-resiliency">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Avg Resiliency Score</div>
                <div class="card-value $(if ($summary.ZoneResiliencyAnalysis.AverageResiliencyScore -ge 75) { 'green' } elseif ($summary.ZoneResiliencyAnalysis.AverageResiliencyScore -ge 40) { 'yellow' } else { 'red' })">$($summary.ZoneResiliencyAnalysis.AverageResiliencyScore)/100</div>
            </div>
            <div class="card">
                <div class="card-label">High (75-100)</div>
                <div class="card-value green">$($summary.ZoneResiliencyAnalysis.HighResiliency)</div>
            </div>
            <div class="card">
                <div class="card-label">Medium (40-74)</div>
                <div class="card-value yellow">$($summary.ZoneResiliencyAnalysis.MediumResiliency)</div>
            </div>
            <div class="card">
                <div class="card-label">Low (0-39)</div>
                <div class="card-value red">$($summary.ZoneResiliencyAnalysis.LowResiliency)</div>
            </div>
        </div>
        
        <div class="table-wrap">
            <div class="table-title">Zone Resiliency by Host Pool</div>
            <div class="table-scroll">
                <table id="zone-table">
                    <thead><tr>
                        <th onclick="sortTable('zone-table',0)">Host Pool</th>
                        <th onclick="sortTable('zone-table',1)">Score</th>
                        <th onclick="sortTable('zone-table',2)">Zone Distribution</th>
                        <th onclick="sortTable('zone-table',3)">Recommendations</th>
                    </tr></thead>
                    <tbody>
"@

  foreach ($zr in ($zoneResiliency | Sort-Object ResiliencyScore)) {
    $zrColor = if ($zr.ResiliencyScore -ge 75) { "b-green" } elseif ($zr.ResiliencyScore -ge 40) { "b-yellow" } else { "b-red" }
    
    $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $zr.HostPoolName)</strong></td>
                        <td><span class="badge $zrColor">$($zr.ResiliencyScore)/100</span></td>
                        <td style="font-size:12px">$($zr.ZoneDistribution)</td>
                        <td style="font-size:12px">$($zr.Recommendations)</td>
                    </tr>
"@
  }

  $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- ========== IMAGES ========== -->
    <div class="section" id="sec-images">
        <!-- Golden Image Maturity Score -->
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Golden Image Grade</div>
                <div class="card-value $(if ($goldenImageScore -ge 60) { 'green' } elseif ($goldenImageScore -ge 40) { 'yellow' } else { 'red' })">$goldenImageGrade</div>
                <div class="card-sub">$goldenImageScore / 100</div>
            </div>
            <div class="card">
                <div class="card-label">Image Sources</div>
                <div class="card-value blue">$((@($imageAnalysis | Select-Object -ExpandProperty ImageSource -Unique)).Count)</div>
                <div class="card-sub">$($galleryVms.Count) Gallery, $($marketplaceVms.Count) Marketplace</div>
            </div>
            <div class="card">
                <div class="card-label">Version Drift</div>
                <div class="card-value $(if ($multiVersionImages -gt 0) { 'orange' } else { 'green' })">$multiVersionImages</div>
                <div class="card-sub">image groups with drift</div>
            </div>
            <div class="card">
                <div class="card-label">Stale Images</div>
                <div class="card-value $(if (@($galleryAnalysis | Where-Object { $_.AgeDays -and $_.AgeDays -gt 90 }).Count -gt 0) { 'red' } else { 'green' })">$(($galleryAnalysis | Where-Object { $_.AgeDays -and $_.AgeDays -gt 90 } | Measure-Object).Count)</div>
                <div class="card-sub">&gt; 90 days since publish</div>
            </div>
        </div>
        
        <!-- Golden Image Maturity Assessment -->
        <div class="table-wrap" style="margin-bottom:16px">
            <div class="table-title">🏅 Golden Image Maturity Assessment</div>
            <div style="padding:16px">
$(foreach ($gf in $goldenImageFindings) {
  "                <div style='padding:6px 0;font-size:13px;border-bottom:1px solid #f0f0f0'>$gf</div>"
})
            </div>
        </div>

$(if ($multiVersionImages -gt 0 -or $marketplaceVms.Count -gt 0) {
  $alertMsg = @()
  if ($multiVersionImages -gt 0) { $alertMsg += "$multiVersionImages image group(s) have version drift" }
  if ($marketplaceVms.Count -gt 0) { $alertMsg += "$($marketplaceVms.Count) VM(s) using raw marketplace images without a golden image pipeline" }
  "<div class='alert alert-warning'><strong>Image Findings:</strong> $($alertMsg -join '. '). Standardize on Azure Compute Gallery with a monthly image refresh cycle.</div>"
})

        <!-- Notes from the Field -->
        <div class="table-wrap" style="margin-bottom:16px">
            <div class="table-title">📋 Notes from the Field — Image Best Practices</div>
            <div style="padding:16px;font-size:13px;line-height:1.7;color:#444">
"@

  # Build context-aware tips based on what we actually found
  $fieldNotes = [System.Collections.Generic.List[string]]::new()

  # Marketplace vs Golden Image
  if ($marketplaceVms.Count -gt 0 -and $galleryVms.Count -eq 0) {
    $fieldNotes.Add("<p><strong>🔧 Build a Golden Image Pipeline.</strong> Every VM in this environment is running a raw marketplace image. That means each session host boots up without your apps pre-installed, GPOs aren't baked in, and FSLogix has to do more heavy lifting on first login. In practice, this adds 2–5 minutes to a user's first session and creates inconsistency between hosts. The recommended approach: stand up an Azure Compute Gallery, build a golden image with your apps and configurations using Azure Image Builder or Packer, and publish new versions monthly. The upfront investment is a day or two — the payoff is faster logins, consistent user experience, and dramatically simpler troubleshooting when every host starts from the same baseline.</p>")
  }
  elseif ($marketplaceVms.Count -gt 0 -and $galleryVms.Count -gt 0) {
    $fieldNotes.Add("<p><strong>🔧 Finish the Migration to Gallery Images.</strong> You've already got some hosts on gallery images — that's the right direction. The remaining $($marketplaceVms.Count) marketplace VM(s) should be migrated to match. Mixed sources make troubleshooting harder because you're comparing hosts with different baselines. A common pattern we see: a team starts with marketplace images for a quick proof-of-concept, then builds a golden image pipeline but never goes back to reimage the originals. Now's a good time to clean that up.</p>")
  }

  # Version drift
  if ($multiVersionImages -gt 0) {
    $fieldNotes.Add("<p><strong>🔄 Tackle Version Drift with Drain-and-Reimage.</strong> Multiple image versions in the same pool means some hosts have different patches, apps, or configurations than others. Users notice — 'it works on one machine but not another' is almost always image drift. The safe rollout pattern: drain old-version hosts (stop new sessions), wait for existing sessions to end, reimage to the latest version, then undrain. Do this in batches — never reimage more than 25% of a pool at once. For pooled desktops, this is seamless since users reconnect to a different host. For personal desktops, schedule during a maintenance window.</p>")
  }

  # Stale gallery images
  $staleGallery = @($galleryAnalysis | Where-Object { $_.AgeDays -and $_.AgeDays -gt 90 })
  if ($staleGallery.Count -gt 0) {
    $worstAge = ($staleGallery | Sort-Object AgeDays -Descending | Select-Object -First 1).AgeDays
    $fieldNotes.Add("<p><strong>📅 Refresh Images Monthly.</strong> The oldest gallery image in this environment is $worstAge days old. Microsoft releases security patches on the second Tuesday of every month (Patch Tuesday). Every month your golden image goes un-updated, your session hosts boot with a known-vulnerable baseline — and FSLogix profile load on first login gets heavier as Windows Update has to catch up. Best practice is to automate image builds on the Wednesday after Patch Tuesday: build the image, run validation tests in a staging host pool, then promote to production. Azure Image Builder or a DevOps pipeline with Packer makes this a hands-off process once set up.</p>")
  }

  # AVD-optimized images
  $nonAvdOptimized = @($imageAnalysis | Where-Object { $_.ImageSource -eq "Marketplace" -and $_.IsAVDOptimized -ne $true -and $_.OSGeneration -match "Windows 1[01]" })
  if ($nonAvdOptimized.Count -gt 0) {
    $fieldNotes.Add("<p><strong>⚡ Use AVD-Optimized Image SKUs.</strong> Some hosts are using standard Windows desktop images rather than the AVD-optimized variants (SKU names containing 'avd'). The AVD-optimized images include pre-configured settings for multi-session environments: reduced disk I/O from telemetry, disabled unnecessary services, optimized visual effects, and pre-installed Teams media optimization. These small tweaks collectively improve login times and session density. When building your next golden image, start from a <code>win11-24h2-avd</code> base instead of the standard desktop SKU.</p>")
  }

  # Legacy OS
  $legacyImages = @($imageAnalysis | Where-Object { $_.OSAge -eq "Legacy" -or $_.OSAge -eq "End of Life" })
  if ($legacyImages.Count -gt 0) {
    $eolImages = @($legacyImages | Where-Object { $_.OSAge -eq "End of Life" })
    if ($eolImages.Count -gt 0) {
      $fieldNotes.Add("<p><strong>🚨 End-of-Life OS Detected.</strong> Some session hosts are running an OS version that no longer receives security updates. This is a compliance risk in most regulated environments and should be treated as urgent. Plan a migration to Windows 11 24H2 multi-session or Server 2022 — both are supported through at least 2027. If application compatibility is blocking the upgrade, consider running the legacy apps in RemoteApp while migrating the base desktop to a current OS.</p>")
    } else {
      $fieldNotes.Add("<p><strong>⏳ Plan Your OS Upgrade.</strong> Some hosts are running older OS builds (21H2 or earlier) that are approaching end of extended support. These still receive patches today, but Microsoft's cadence of feature updates means newer builds get better AVD integration, improved Teams performance, and lower resource overhead. Build your next golden image on Windows 11 24H2 — it includes the latest GPU and multimedia redirect improvements that directly affect user experience.</p>")
    }
  }

  # All good
  if ($goldenImageScore -ge 80) {
    $fieldNotes.Add("<p><strong>✅ Strong Image Hygiene.</strong> This environment follows golden image best practices — gallery-based images, consistent versions across pools, and current OS builds. To maintain this, consider adding automated validation: after each image build, deploy 2–3 session hosts to a staging pool, run synthetic login tests (Azure Load Testing or a simple PowerShell script that launches a remote session), and only promote to production after validation passes. This catches broken apps or misconfigured GPOs before they affect real users.</p>")
  }

  # Multi-session vs single-session advice
  $singleSessionDesktop = @($imageAnalysis | Where-Object { $_.IsMultiSession -ne $true -and $_.OSGeneration -match "Windows 1[01]" -and $_.ImageSource -eq "Marketplace" })
  if ($singleSessionDesktop.Count -gt 0) {
    $singleSessionVmCount = ($singleSessionDesktop | Measure-Object -Property VMCount -Sum).Sum
    if ($singleSessionVmCount -gt 5) {
      $fieldNotes.Add("<p><strong>👥 Consider Multi-Session.</strong> $singleSessionVmCount VM(s) appear to be running single-session Windows desktop images. If these are pooled desktops with similar user workloads, Windows 11 Enterprise multi-session can host 4–8 users per VM (depending on workload), reducing your VM count and cost by 60–75%. Multi-session is not appropriate for every workload — GPU-heavy, developer, or compliance-isolated scenarios still warrant dedicated VMs — but for general knowledge workers, it's a significant cost optimization.</p>")
    }
  }

  foreach ($fn in $fieldNotes) {
    $htmlReport += $fn
  }

  if ($fieldNotes.Count -eq 0) {
    $htmlReport += "<p style='color:#888'>No specific image recommendations — environment is well-configured.</p>"
  }

  $htmlReport += @"
            </div>
        </div>
        
        <!-- Host Pool Image Consistency -->
        <div class="table-wrap">
            <div class="table-title">Host Pool Image Consistency</div>
            <div class="table-scroll">
                <table id="hp-image-table">
                    <thead><tr>
                        <th onclick="sortTable('hp-image-table',0)">Host Pool</th>
                        <th onclick="sortTable('hp-image-table',1)">VMs</th>
                        <th onclick="sortTable('hp-image-table',2)">Sources</th>
                        <th onclick="sortTable('hp-image-table',3)">SKUs</th>
                        <th onclick="sortTable('hp-image-table',4)">Versions</th>
                        <th onclick="sortTable('hp-image-table',5)">Consistency</th>
                        <th>Finding</th>
                    </tr></thead>
                    <tbody>
"@

  foreach ($hpImg in ($hpImageConsistency | Sort-Object { switch ($_.Consistency) { "Mixed Sources" { 0 } "Mixed SKUs" { 1 } "Version Drift" { 2 } default { 3 } } })) {
    $consBadge = switch ($hpImg.Consistency) {
      "Mixed Sources"  { "<span class='badge b-red'>Mixed Sources</span>" }
      "Mixed SKUs"     { "<span class='badge b-orange'>Mixed SKUs</span>" }
      "Version Drift"  { "<span class='badge b-yellow'>Version Drift</span>" }
      default          { "<span class='badge b-green'>Consistent</span>" }
    }
    
    $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $hpImg.HostPoolName)</strong></td>
                        <td>$($hpImg.VMCount)</td>
                        <td>$($hpImg.ImageSources)</td>
                        <td style="font-size:12px">$($hpImg.ImageSkus)</td>
                        <td style="font-size:12px">$($hpImg.ImageVersions)</td>
                        <td>$consBadge</td>
                        <td style="font-size:12px">$($hpImg.Finding)</td>
                    </tr>
"@
  }

  $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Detailed Image Groups -->
        <div class="table-wrap">
            <div class="table-title">Image Groups</div>
            <div class="table-scroll">
                <table id="image-table">
                    <thead><tr>
                        <th onclick="sortTable('image-table',0)">Source</th>
                        <th onclick="sortTable('image-table',1)">Image</th>
                        <th onclick="sortTable('image-table',2)">VMs</th>
                        <th onclick="sortTable('image-table',3)">OS</th>
                        <th onclick="sortTable('image-table',4)">Versions</th>
$(if ($galleryAnalysis.Count -gt 0) { "                        <th>Image Age</th>" })
                        <th>Finding</th>
                    </tr></thead>
                    <tbody>
"@

  foreach ($img in $imageAnalysis) {
    $srcBadge = switch ($img.ImageSource) {
      "Gallery"      { "<span class='badge b-green'>Gallery</span>" }
      "Marketplace"  { "<span class='badge b-blue'>Marketplace</span>" }
      "ManagedImage" { "<span class='badge b-yellow'>Managed</span>" }
      default        { "<span class='badge b-gray'>$($img.ImageSource)</span>" }
    }
    $imgName = if ($img.ImageSource -eq "Gallery") { "$($img.GalleryName) / $($img.GalleryImageDef)" }
               else { "$($img.ImagePublisher) / $($img.ImageOffer) / $($img.ImageSku)" }
    $osDisplay = if ($img.OSGeneration) { "$($img.OSGeneration) $($img.OSBuild)" } else { "—" }
    $ageDisplay = if ($img.VersionAge) {
      $ageColor = if ($img.VersionAge -gt 180) { "color:#c62828;font-weight:700" } elseif ($img.VersionAge -gt 90) { "color:#e65100;font-weight:600" } else { "" }
      "<span style='$ageColor'>$($img.VersionAge) days</span>"
    } else { "—" }
    $findingColor = if ($img.Finding -match "End of Life|severely outdated") { "color:#c62828" } elseif ($img.Finding -match "drift|Legacy|missing") { "color:#e65100" } else { "" }
    
    $htmlReport += @"
                    <tr>
                        <td>$srcBadge</td>
                        <td style="font-size:12px"><strong>$imgName</strong></td>
                        <td>$($img.VMCount)</td>
                        <td style="font-size:12px">$osDisplay</td>
                        <td style="font-size:12px">$($img.VersionsInUse)</td>
$(if ($galleryAnalysis.Count -gt 0) { "                        <td>$ageDisplay</td>" })
                        <td style="font-size:12px;$findingColor">$($img.Finding)</td>
                    </tr>
"@
  }

  $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
    </div>
"@

  # ========== COST BREAKDOWN ==========
  $htmlReport += @"
    
    <div class="section" id="sec-costs">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Current Monthly$(if ($costSourceLabel -eq 'Actual Billing') { ' (Actual)' } else { ' (Est.)' })</div>
                <div class="card-value blue">~`$$([math]::Round($totalCurrentCost, 0))</div>
                <div class="card-sub">Compute + Disks</div>
            </div>
$(if ($infraCostTotal -gt 0) { @"
            <div class="card">
                <div class="card-label">Infrastructure/mo (Actual)</div>
                <div class="card-value blue">~`$$([math]::Round($infraCostTotal, 0))</div>
                <div class="card-sub">Network + Storage + AVD Service</div>
            </div>
            <div class="card">
                <div class="card-label">Full AVD Monthly</div>
                <div class="card-value blue">~`$$([math]::Round($totalCurrentCost + $infraCostTotal, 0))</div>
                <div class="card-sub">Compute + Infra</div>
            </div>
"@
})
            <div class="card">
                <div class="card-label">Est. Monthly Savings</div>
                <div class="card-value green">~`$$($summary.RightSizingAnalysis.PotentialMonthlySavings)</div>
            </div>
            <div class="card">
                <div class="card-label">Est. Annual Savings</div>
                <div class="card-value green">~`$$($summary.RightSizingAnalysis.PotentialAnnualSavings)</div>
            </div>
            <div class="card">
                <div class="card-label">Cost Source</div>
                <div class="card-value $(if ($costSourceLabel -eq 'Actual Billing') { 'green' } else { 'orange' })">$costSourceLabel</div>
            </div>
        </div>
$(if ($infraCostByCategory.Count -gt 0) { @"
        <div class="table-wrap" style="margin-bottom:16px">
            <div class="table-title">Infrastructure Cost Breakdown — AVD Resource Groups (Last 30 Days Projected Monthly)</div>
            <div class="table-scroll">
                <table>
                    <thead><tr><th>Category</th><th>Monthly Cost</th><th>% of Infra</th></tr></thead>
                    <tbody>
"@
  foreach ($catEntry in ($infraCostByCategory.GetEnumerator() | Sort-Object Value -Descending)) {
    $catPct = if ($infraCostTotal -gt 0) { [math]::Round($catEntry.Value / $infraCostTotal * 100, 1) } else { 0 }
    "                    <tr><td><strong>$($catEntry.Key)</strong></td><td>`$$([math]::Round($catEntry.Value, 0))</td><td>$catPct%</td></tr>`n"
  }
@"
                    <tr style="font-weight:600;border-top:2px solid #333"><td>Total Infrastructure</td><td>`$$([math]::Round($infraCostTotal, 0))</td><td>100%</td></tr>
                    </tbody>
                </table>
            </div>
        </div>
"@
})
$(if ($hostPoolCosts.Count -gt 0) { @"
        <div class="table-wrap">
            <div class="table-title">Actual Cost by Host Pool (Last 30 Days Projected Monthly — Amortized)</div>
            <div class="table-scroll">
                <table id="hpcost-table">
                    <thead><tr>
                        <th onclick="sortTable('hpcost-table',0)">Host Pool</th>
                        <th onclick="sortTable('hpcost-table',1)">VMs</th>
                        <th onclick="sortTable('hpcost-table',2)">Total/mo</th>
                        <th onclick="sortTable('hpcost-table',3)">Compute/mo</th>
                        <th onclick="sortTable('hpcost-table',4)">Storage/mo</th>
                        <th onclick="sortTable('hpcost-table',5)">Avg/VM/mo</th>
                        <th onclick="sortTable('hpcost-table',6)">Pricing</th>
                    </tr></thead>
                    <tbody>
"@ })
"@

  # Render host pool cost rows
  if ($hostPoolCosts.Count -gt 0) {
    foreach ($hpn in ($hostPoolCosts.Keys | Sort-Object)) {
      $hpc = $hostPoolCosts[$hpn]
      # Determine dominant pricing model for this pool
      $poolCostEntries = @($actualCostData | Where-Object { $_.HostPoolName -eq $hpn })
      $riCount = @($poolCostEntries | Where-Object { $_.PricingModel -match 'Reserved|Savings' }).Count
      $totalCount = [math]::Max(1, $poolCostEntries.Count)
      $pricingBadge = if ($riCount -eq $totalCount -and $riCount -gt 0) { "<span class='badge b-green'>RI/SP</span>" }
                      elseif ($riCount -gt 0) { "<span class='badge b-orange'>Mixed ($riCount/$totalCount RI)</span>" }
                      else { "<span style='color:#888'>PAYG</span>" }
      $htmlReport += "                    <tr><td><strong>$(Scrub-HostPoolName $hpn)</strong></td><td>$($hpc.VMCount)</td><td>`$$([math]::Round($hpc.TotalMonthly, 0))</td><td>`$$([math]::Round($hpc.ComputeMonthly, 0))</td><td>`$$([math]::Round($hpc.StorageMonthly, 0))</td><td>`$$([math]::Round($hpc.TotalMonthly / [math]::Max(1, $hpc.VMCount), 0))</td><td>$pricingBadge</td></tr>`n"
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
  }

  $htmlReport += @"
        
        <div class="table-wrap">
            <div class="table-title">Top Cost Impact VMs (Savings &amp; Increases)</div>
            <div class="table-scroll">
                <table id="cost-table">
                    <thead><tr>
                        <th onclick="sortTable('cost-table',0)">VM Name</th>
                        <th onclick="sortTable('cost-table',1)">Host Pool</th>
                        <th onclick="sortTable('cost-table',2)">Current Size</th>
                        <th onclick="sortTable('cost-table',3)">Recommended</th>
                        <th onclick="sortTable('cost-table',4)">Monthly Impact</th>
                        <th onclick="sortTable('cost-table',5)">Confidence</th>
                    </tr></thead>
                    <tbody>
"@

  foreach ($cb in $costBreakdown) {
    $impactVal = [double]$cb.EstimatedMonthlySavings
    $impactColor = if ($impactVal -gt 0) { "#2e7d32" } else { "#c62828" }
    $impactSign = if ($impactVal -gt 0) { "-" } else { "+" }
    $impactDisplay = "$impactSign`$$([math]::Abs($impactVal))/mo"
    $confBadge = switch ($cb.Confidence) { "High" { "b-green" }; "Medium" { "b-yellow" }; default { "b-red" } }
    
    $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-VMName $cb.VMName)</strong></td>
                        <td>$(Scrub-HostPoolName $cb.HostPoolName)</td>
                        <td>$($cb.CurrentSize)</td>
                        <td>$($cb.RecommendedSize)</td>
                        <td style="color:$impactColor;font-weight:600">$impactDisplay</td>
                        <td><span class="badge $confBadge">$($cb.Confidence)</span></td>
                    </tr>
"@
  }

  $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
        <div class="alert alert-$(if ($costSourceLabel -eq 'Actual Billing') { 'info' } else { 'warning' })">
            $(if ($costSourceLabel -eq 'Actual Billing') { "<strong>📊 Cost Source:</strong> Current costs use actual Azure billing data (last 30 days). Savings estimates for recommended SKUs use Azure retail pricing since those VMs don't exist yet." } else { "<strong>⚠️ Reminder:</strong> All costs are estimates based on Azure retail pricing. Actual costs vary by EA/CSP agreements, reserved instances, and contract rates. Use <code>-SkipActualCosts</code> to disable Cost Management queries." })
        </div>
    </div>
"@

  # ========== PRIORITY MATRIX ==========
  # Categorize all findings into effort vs impact quadrants
  $priorityItems = [System.Collections.Generic.List[object]]::new()
  
  # Stuck-in-drain hosts — high impact, low effort (just undrain or restart)
  if ($stuckHosts -gt 0) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$stuckHosts host(s) stuck in drain mode"; Impact = "High"; Effort = "Low"; Category = "Host Health"; Action = "Undrain or restart affected hosts"; QuickWin = $true })
  }
  
  # Unavailable hosts
  if ($unavailableHosts -gt 0) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$unavailableHosts host(s) unavailable / needs assistance"; Impact = "High"; Effort = "Low"; Category = "Host Health"; Action = "Investigate and remediate"; QuickWin = $true })
  }
  
  # AccelNet gaps — medium impact, low effort (requires restart)
  if ($eligibleNotEnabled -gt 0) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$eligibleNotEnabled VM(s) eligible for AccelNet but not enabled"; Impact = "Medium"; Effort = "Low"; Category = "Network"; Action = "Enable AccelNet (requires VM restart)"; QuickWin = $true })
  }
  
  # Upsize candidates — high impact, medium effort
  $upsizeCount = (SafeCount $upsizeCandidates)
  if ($upsizeCount -gt 0) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$upsizeCount VM(s) need more resources (upsize)"; Impact = "High"; Effort = "Medium"; Category = "Right-Sizing"; Action = "Resize VMs during maintenance window"; QuickWin = $false })
  }
  
  # Downsize candidates — medium impact (cost), medium effort
  $downsizeCount = (SafeCount $downsizeCandidates)
  if ($downsizeCount -gt 0) {
    $savingsEst = $summary.RightSizingAnalysis.PotentialMonthlySavings
    $priorityItems.Add([PSCustomObject]@{ Finding = "$downsizeCount VM(s) can be downsized (~`$$savingsEst/mo est. savings)"; Impact = "Medium"; Effort = "Medium"; Category = "Cost"; Action = "Validate with users, then resize"; QuickWin = $false })
  }
  
  # Premium SSD on pooled — medium impact (cost), low effort
  if ($premiumOnPooled -gt 0) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$premiumOnPooled pooled host(s) using Premium SSD"; Impact = "Medium"; Effort = "Low"; Category = "Storage"; Action = "Switch to Standard SSD or Ephemeral"; QuickWin = $true })
  }
  
  # Non-ephemeral on pooled — medium impact, medium effort
  if ($nonEphemeral -gt 0) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$nonEphemeral pooled host(s) not using Ephemeral OS disk"; Impact = "Medium"; Effort = "Medium"; Category = "Storage"; Action = "Rebuild with ephemeral disk (faster reimage, lower cost)"; QuickWin = $false })
  }
  
  # Image & Golden Image findings (v4.0.0)
  if ($multiVersionImages -gt 0) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$multiVersionImages image group(s) have version drift — hosts running different image versions"; Impact = "Medium"; Effort = "Medium"; Category = "Images"; Action = "Reimage drifted hosts to latest golden image version"; QuickWin = $false })
  }
  if ($marketplaceVms.Count -gt 0 -and $totalVmCount -gt 0 -and $marketplacePct -ge 50) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$($marketplaceVms.Count) VM(s) ($marketplacePct%) using raw marketplace images — no golden image pipeline"; Impact = "High"; Effort = "High"; Category = "Images"; Action = "Implement Azure Compute Gallery with a golden image pipeline (Packer/DevOps/Image Builder) for consistent, pre-configured images"; QuickWin = $false })
  }
  $staleImages = @($galleryAnalysis | Where-Object { $_.AgeDays -and $_.AgeDays -gt 90 })
  if ($staleImages.Count -gt 0) {
    $worstAge = ($staleImages | Sort-Object AgeDays -Descending | Select-Object -First 1).AgeDays
    $priorityItems.Add([PSCustomObject]@{ Finding = "$($staleImages.Count) gallery image(s) older than 90 days (oldest: $worstAge days) — missing security patches"; Impact = "High"; Effort = "Medium"; Category = "Images"; Action = "Update golden images monthly; automate with Azure Image Builder or DevOps pipeline"; QuickWin = $false })
  }
  $mixedSourcePools = @($hpImageConsistency | Where-Object { $_.Consistency -eq "Mixed Sources" })
  if ($mixedSourcePools.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$($mixedSourcePools.Count) host pool(s) have mixed image sources (marketplace + custom) — inconsistent user experience"; Impact = "Medium"; Effort = "Medium"; Category = "Images"; Action = "Standardize all hosts in each pool on the same golden image"; QuickWin = $false })
  }
  $legacyOsImages = @($imageAnalysis | Where-Object { $_.OSAge -eq "Legacy" -or $_.OSAge -eq "End of Life" })
  if ($legacyOsImages.Count -gt 0) {
    $eolCount = @($legacyOsImages | Where-Object { $_.OSAge -eq "End of Life" }).Count
    $legacyCount = @($legacyOsImages | Where-Object { $_.OSAge -eq "Legacy" }).Count
    if ($eolCount -gt 0) {
      $priorityItems.Add([PSCustomObject]@{ Finding = "$eolCount image(s) on End of Life OS — no security updates"; Impact = "High"; Effort = "High"; Category = "Images"; Action = "Migrate to supported OS immediately (Windows 11 24H2 or Server 2022 recommended)"; QuickWin = $false })
    }
    if ($legacyCount -gt 0) {
      $priorityItems.Add([PSCustomObject]@{ Finding = "$legacyCount image(s) on legacy OS build — approaching end of support"; Impact = "Medium"; Effort = "High"; Category = "Images"; Action = "Plan migration to current OS build during next image refresh cycle"; QuickWin = $false })
    }
  }
  if ($goldenImageScore -lt 40) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "Golden Image Maturity score $goldenImageScore/100 (Grade: $goldenImageGrade) — ad-hoc image management"; Impact = "Medium"; Effort = "High"; Category = "Images"; Action = "Implement structured golden image pipeline: Azure Compute Gallery + Image Builder + monthly refresh cycle"; QuickWin = $false })
  }
  
  # Zone resiliency — high impact, high effort
  $avgResiliency = ($summary.ZoneResiliencyAnalysis.AverageResiliencyScore)
  if ($avgResiliency -and [double]$avgResiliency -lt 50) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "Average zone resiliency score: $avgResiliency/100"; Impact = "High"; Effort = "High"; Category = "Resiliency"; Action = "Redistribute VMs across availability zones"; QuickWin = $false })
  }
  
  # Reservation opportunities — high impact, low effort (buying an RI is a purchase, not a migration)
  if ($IncludeReservationAnalysis -and $totalUncoveredVMs -gt 0) {
    $alwaysOnUncovered = @($reservationAnalysis | Where-Object { $_.Status -eq "Uncovered (Always-On)" -and $_.UncoveredVMs -gt 0 })
    if ($alwaysOnUncovered.Count -gt 0) {
      $alwaysOnTotal = ($alwaysOnUncovered | Measure-Object -Property UncoveredVMs -Sum).Sum
      $alwaysOnSavings = ($alwaysOnUncovered | Measure-Object -Property TotalMonthlySavings3Y -Sum).Sum
      $priorityItems.Add([PSCustomObject]@{ Finding = "$alwaysOnTotal always-on personal desktop(s) without RI coverage (~`$$([math]::Round($alwaysOnSavings, 0))/mo savings)"; Impact = "High"; Effort = "Low"; Category = "Reservations"; Action = "Purchase 3-year RIs for stable personal desktops"; QuickWin = $true })
    }
    $pooledUncovered = @($reservationAnalysis | Where-Object { $_.Status -eq "Uncovered" -and $_.UncoveredVMs -gt 0 })
    if ($pooledUncovered.Count -gt 0) {
      $pooledTotal = ($pooledUncovered | Measure-Object -Property UncoveredVMs -Sum).Sum
      $pooledSavings = ($pooledUncovered | Measure-Object -Property TotalMonthlySavings3Y -Sum).Sum
      $priorityItems.Add([PSCustomObject]@{ Finding = "$pooledTotal pooled VM(s) without RI coverage (~`$$([math]::Round($pooledSavings, 0))/mo savings)"; Impact = "Medium"; Effort = "Low"; Category = "Reservations"; Action = "Evaluate fleet stability, then purchase 1yr or 3yr RIs"; QuickWin = $true })
    }
  }
  if ($IncludeReservationAnalysis -and $totalOverProvisioned -gt 0) {
    $priorityItems.Add([PSCustomObject]@{ Finding = "$totalOverProvisioned over-provisioned RI(s) — paying for unused reservations"; Impact = "Medium"; Effort = "Low"; Category = "Reservations"; Action = "Exchange for different SKU or let expire"; QuickWin = $true })
  }

  # Cross-Region / Cross-Continent connection findings
  if ($crossContinentPaths.Count -gt 0) {
    $ccUsers = ($crossContinentPaths | ForEach-Object { [int]$_.DistinctUsers } | Measure-Object -Sum).Sum
    $ccConns = ($crossContinentPaths | ForEach-Object { [int]$_.Connections } | Measure-Object -Sum).Sum
    $ccWorstP95 = ($crossContinentPaths | Sort-Object { [double]$_.P95RTTms } -Descending | Select-Object -First 1)
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($crossContinentPaths.Count) cross-continent connection path(s) detected ($ccUsers users, worst P95: $($ccWorstP95.P95RTTms)ms from $($ccWorstP95.GatewayRegion)→$($ccWorstP95.HostRegion))"
      Impact   = "High"
      Effort   = "High"
      Category = "Geo-Latency"
      Action   = "Deploy session hosts in regions closer to remote users; enable RDP Shortpath for direct UDP connections"
      QuickWin = $false
    })
  }
  $highRTTSameRegion = @($crossRegionAnalysis | Where-Object { -not $_.IsCrossRegion -and [double]$_.P95RTTms -gt 100 })
  if ($highRTTSameRegion.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($highRTTSameRegion.Count) same-region path(s) with P95 RTT > 100ms — possible network issues"
      Impact   = "Medium"
      Effort   = "Medium"
      Category = "Geo-Latency"
      Action   = "Investigate client-side network, ISP routing, or consider enabling RDP Shortpath"
      QuickWin = $false
    })
  }
  $highExcessRTT = @($crossRegionAnalysis | Where-Object { $_.RTTExcessMs -gt 100 })
  if ($highExcessRTT.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($highExcessRTT.Count) path(s) with RTT significantly above expected baseline (>100ms excess) — routing inefficiency"
      Impact   = "Medium"
      Effort   = "Low"
      Category = "Geo-Latency"
      Action   = "Enable RDP Shortpath to bypass gateway relay; check for VPN/proxy adding hops"
      QuickWin = $true
    })
  }

  # SKU Diversity & Allocation Resilience findings (v4.0.0)
  if ($highRiskPools.Count -gt 0) {
    $hpNames = ($highRiskPools | ForEach-Object { $_.HostPoolName }) -join ", "
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($highRiskPools.Count) host pool(s) at high risk of allocation failure — single SKU family ($hpNames)"
      Impact   = "High"
      Effort   = "Medium"
      Category = "Allocation Risk"
      Action   = "Diversify VM SKUs across 2+ families; mix D-series with E-series or different generations"
      QuickWin = $false
    })
  }
  if ($medRiskPools.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($medRiskPools.Count) host pool(s) with moderate SKU concentration risk"
      Impact   = "Medium"
      Effort   = "Medium"
      Category = "Allocation Risk"
      Action   = "Increase SKU diversity; consider adding VMs from alternate series during next scale-out"
      QuickWin = $false
    })
  }
  $allSingleRegion = @($skuDiversityAnalysis | Where-Object { $_.UniqueRegions -eq 1 -and $_.VMCount -gt 1 })
  if ($allSingleRegion.Count -eq $skuDiversityAnalysis.Count -and $skuDiversityAnalysis.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "All $($skuDiversityAnalysis.Count) host pool(s) deployed in a single region — no geo-redundancy for regional outages"
      Impact   = "High"
      Effort   = "High"
      Category = "Allocation Risk"
      Action   = "Deploy secondary host pool in a paired region for disaster recovery and allocation resilience"
      QuickWin = $false
    })
  }

  # Network Readiness findings (v4.0.0)
  if (-not $shortpathEnabled -and $shortpathSummary.TotalConnections -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "RDP Shortpath not enabled — all $($shortpathSummary.TotalConnections) connections using TCP relay through gateway"
      Impact   = "High"
      Effort   = "Low"
      Category = "Network"
      Action   = "Enable RDP Shortpath for managed/public networks to reduce latency via direct UDP"
      QuickWin = $true
    })
  }
  elseif ($shortpathEnabled -and $shortpathSummary.ShortpathPct -lt 50) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "Only $($shortpathSummary.ShortpathPct)% of connections using RDP Shortpath — most still on TCP relay"
      Impact   = "Medium"
      Effort   = "Low"
      Category = "Network"
      Action   = "Check client versions and firewall rules (UDP 3390 required); ensure Shortpath is enabled on host pool"
      QuickWin = $true
    })
  }
  $criticalSubnets = @($subnetAnalysis | Where-Object { $_.UsagePct -ge 90 })
  if ($criticalSubnets.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($criticalSubnets.Count) subnet(s) at >90% capacity — scaling may fail due to IP exhaustion"
      Impact   = "High"
      Effort   = "Medium"
      Category = "Network"
      Action   = "Expand subnets or migrate to larger address space before next scale-out event"
      QuickWin = $false
    })
  }
  $disconnectedPeers = @($vnetAnalysis | Where-Object { $_.DisconnectedPeers -gt 0 })
  if ($disconnectedPeers.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($disconnectedPeers.Count) VNet(s) with disconnected peering — may break FSLogix, domain join, or app access"
      Impact   = "High"
      Effort   = "Low"
      Category = "Network"
      Action   = "Reconnect VNet peerings immediately"
      QuickWin = $true
    })
  }
  if ($hpWithoutPE.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($hpWithoutPE.Count) host pool(s) without private endpoints"
      Impact   = "Medium"
      Effort   = "Medium"
      Category = "Network"
      Action   = "Configure private endpoints to keep AVD control plane traffic on the Microsoft backbone"
      QuickWin = $false
    })
  }
  if ($vmsWithNoNsgAtAll.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($vmsWithNoNsgAtAll.Count) session host(s) with no NSG at NIC or subnet level"
      Impact   = "Medium"
      Effort   = "Low"
      Category = "Network"
      Action   = "Attach an NSG to the session host subnet with appropriate AVD rules"
      QuickWin = $true
    })
  }
  
  # Disconnect Reason findings (v4.0.0)
  if ($disconnectReasonData.Count -gt 0) {
    $dcNormalCats = @("Normal Completion", "User Initiated", "No Completion Record", "Idle Timeout", "Auto-Reconnect", "Authentication")
    $dcAbnormal = @($disconnectReasonData | Where-Object { $_.DisconnectCategory -notin $dcNormalCats })
    $dcTotalAll = ($disconnectReasonData | ForEach-Object { if ($_.PSObject.Properties['SessionCount']) { [int]$_.SessionCount } else { 0 } } | Measure-Object -Sum).Sum
    $dcTotalAbnormal = ($dcAbnormal | ForEach-Object { if ($_.PSObject.Properties['SessionCount']) { [int]$_.SessionCount } else { 0 } } | Measure-Object -Sum).Sum
    $dcAbnormalPct = if ($dcTotalAll -gt 0) { [math]::Round(($dcTotalAbnormal / $dcTotalAll) * 100, 1) } else { 0 }
    
    if ($dcAbnormalPct -gt 10) {
      $topCause = ($dcAbnormal | Sort-Object { [int]$_.SessionCount } -Descending | Select-Object -First 1)
      $priorityItems.Add([PSCustomObject]@{
        Finding  = "$dcAbnormalPct% abnormal disconnect rate ($dcTotalAbnormal sessions) — top cause: $($topCause.DisconnectCategory)"
        Impact   = "High"
        Effort   = "Medium"
        Category = "Connections"
        Action   = "Investigate $($topCause.DisconnectCategory) disconnects. See Connections tab for per-host breakdown and remediation steps."
        QuickWin = $false
      })
    }
    
    $dcNetworkDrops = @($dcAbnormal | Where-Object { $_.DisconnectCategory -eq "Network Drop" })
    if ($dcNetworkDrops.Count -gt 0 -and ([int]$dcNetworkDrops[0].SessionCount) -gt 10) {
      $priorityItems.Add([PSCustomObject]@{
        Finding  = "$($dcNetworkDrops[0].SessionCount) network drop disconnects affecting $($dcNetworkDrops[0].DistinctUsers) users"
        Impact   = "High"
        Effort   = "Medium"
        Category = "Connections"
        Action   = "Check VPN stability, NIC health, NSG rules. Enable RDP Shortpath for UDP resilience."
        QuickWin = $false
      })
    }
    
    $dcResourceIssues = @($dcAbnormal | Where-Object { $_.DisconnectCategory -eq "Resource Exhaustion" })
    if ($dcResourceIssues.Count -gt 0 -and ([int]$dcResourceIssues[0].SessionCount) -gt 5) {
      $priorityItems.Add([PSCustomObject]@{
        Finding  = "$($dcResourceIssues[0].SessionCount) disconnects from resource exhaustion (memory/disk) on $($dcResourceIssues[0].DistinctHosts) host(s)"
        Impact   = "High"
        Effort   = "Medium"
        Category = "Connections"
        Action   = "Review memory and disk metrics on affected hosts. Consider upsizing or reducing session density."
        QuickWin = $false
      })
    }
  }
  
  # Security Posture findings (v4.0.0)
  $lowSecPools_pm = @($securityPosture | Where-Object { $_.SecurityScore -lt 40 })
  if ($lowSecPools_pm.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($lowSecPools_pm.Count) host pool(s) have security grade D/F — missing Trusted Launch, Secure Boot, or vTPM"
      Impact   = "High"
      Effort   = "Medium"
      Category = "Security"
      Action   = "Redeploy session hosts with Trusted Launch enabled. Requires Gen2 VM images."
      QuickWin = $false
    })
  }
  
  # Orphaned resources findings (v4.0.0)
  if ($orphanedResources.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($orphanedResources.Count) orphaned resource(s) found — ~`$$orphanedWaste/mo wasted"
      Impact   = if ($orphanedWaste -gt 100) { "Medium" } else { "Low" }
      Effort   = "Low"
      Category = "Cost"
      Action   = "Delete unattached disks, NICs, and public IPs. Review ENHANCED-Orphaned-Resources.csv for full list."
      QuickWin = $true
    })
  }
  
  # Profile health findings (v4.0.0)
  $critProfiles = @($profileHealth | Where-Object { $_.Severity -eq "Critical" })
  if ($critProfiles.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($critProfiles.Count) session host(s) have critical profile load times (P95 >= 60s)"
      Impact   = "High"
      Effort   = "Medium"
      Category = "Performance"
      Action   = "Check FSLogix profile container storage IOPS/throughput. Consider Azure Files Premium or ANF. Review ENHANCED-Profile-Health.csv."
      QuickWin = $false
    })
  }
  
  # UX Score findings (v4.0.0)
  $poorUxPools = @($uxScores | Where-Object { $_.UXScore -lt 60 })
  if ($poorUxPools.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($poorUxPools.Count) host pool(s) have poor user experience (UX Score < 60 / grade D-F)"
      Impact   = "High"
      Effort   = "Medium"
      Category = "User Experience"
      Action   = "Review UX Score breakdown: profile loads, RTT, disconnect rates, connection errors. See ENHANCED-UX-Scores.csv."
      QuickWin = $false
    })
  }

  # W365 Cloud PC findings (v4.0.0)
  $w365StrongPools = @($w365Analysis | Where-Object { $_.Recommendation -eq "Strong W365 Candidate" })
  if ($w365StrongPools.Count -gt 0) {
    $w365TotalEstSavings = ($w365StrongPools | Where-Object { $_.CostDelta -and $_.CostDelta -lt 0 } | ForEach-Object { [math]::Abs($_.CostDelta) } | Measure-Object -Sum).Sum
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($w365StrongPools.Count) host pool(s) are strong W365 Cloud PC candidates$(if ($w365TotalEstSavings -gt 0) { " — estimated ~`$$w365TotalEstSavings/mo savings" })"
      Impact   = "High"
      Effort   = "Medium"
      Category = "Cost/Strategy"
      Action   = "Evaluate W365 Enterprise for personal/low-density pools. See W365 Readiness tab for per-pool analysis and cost comparison."
      QuickWin = $false
    })
  }

  # Scaling plan findings (v4.0.0)
  # Pools with no scaling plan that have running VMs
  $allPoolNamesForPM = @($vms | Select-Object -ExpandProperty HostPoolName -Unique)
  $poolsWithPlansPM = @{}
  foreach ($a in $scalingPlanAssignments) {
    $pn = if ($a.HostPoolName) { $a.HostPoolName } elseif ($a.HostPoolArmId) { ($a.HostPoolArmId -split '/')[-1] } else { "" }
    if ($pn) { $poolsWithPlansPM[$pn] = $true }
  }
  $noPlanPoolsPM = @($allPoolNamesForPM | Where-Object { -not $poolsWithPlansPM.ContainsKey($_) })
  $noPlanRunning = 0
  foreach ($np in $noPlanPoolsPM) {
    $noPlanRunning += @($vms | Where-Object { $_.HostPoolName -eq $np -and $_.PowerState -eq "VM running" }).Count
  }
  if ($noPlanPoolsPM.Count -gt 0 -and $noPlanRunning -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($noPlanPoolsPM.Count) host pool(s) have no scaling plan — $noPlanRunning VMs running with no autoscale"
      Impact   = "Medium"
      Effort   = "Low"
      Category = "Cost"
      Action   = "Create and assign scaling plans. See Scaling tab for coverage details."
      QuickWin = $true
    })
  }

  # Disabled scaling plan assignments
  $disabledPlansPM = @($scalingPlanAssignments | Where-Object { $_.IsEnabled -eq $false -or $_.IsEnabled -eq "False" })
  if ($disabledPlansPM.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($disabledPlansPM.Count) scaling plan assignment(s) are disabled — pools may be running 24/7"
      Impact   = "Medium"
      Effort   = "Low"
      Category = "Cost"
      Action   = "Re-enable or investigate why plans were disabled. See Scaling tab."
      QuickWin = $true
    })
  }

  # High ramp-down capacity
  $highRampDownPM = @($scalingPlanSchedules | Where-Object { $_.ScheduleName -eq "Weekdays" -and [int]$_.RampDownCapacity -gt 70 })
  if ($highRampDownPM.Count -gt 2) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($highRampDownPM.Count) scaling plans have weekday ramp-down capacity >70% — VMs stay running past business hours"
      Impact   = "Medium"
      Effort   = "Low"
      Category = "Cost"
      Action   = "Reduce ramp-down capacity to 20-40% for low-usage pools. See Scaling tab for schedule details."
      QuickWin = $true
    })
  }

  # StartVMOnConnect disabled
  $noStartVmPM = @($hostPools | Where-Object { $_.HostPoolName -and (-not $_.StartVMOnConnect -or $_.StartVMOnConnect -eq "False") })
  if ($noStartVmPM.Count -gt 0) {
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$($noStartVmPM.Count) host pool(s) have Start VM on Connect disabled — users hit cold hosts or must wait for pre-started VMs"
      Impact   = "Medium"
      Effort   = "Low"
      Category = "User Experience"
      Action   = "Enable Start VM on Connect in host pool properties. Requires Contributor role on the subscription and Desktop Virtualization Power On Off Contributor on the VMs."
      QuickWin = $true
    })
  }

  # Autoscale failures (from Log Analytics)
  $autoscaleFailedPM = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_WVDAutoscaleDetailed" -and $_.QueryName -eq "AVD" -and $_.PoolName -and $_.PoolName -ne "NoTable" -and [int]$_.Failed -gt 0 })
  if ($autoscaleFailedPM.Count -gt 0) {
    $totalFailedPM = ($autoscaleFailedPM | ForEach-Object { [int]$_.Failed } | Measure-Object -Sum).Sum
    $priorityItems.Add([PSCustomObject]@{
      Finding  = "$totalFailedPM autoscale evaluation(s) failed across $($autoscaleFailedPM.Count) pool(s) — VMs may not be starting/stopping as expected"
      Impact   = "High"
      Effort   = "Medium"
      Category = "Scaling"
      Action   = "Check WVDAutoscaleEvaluationPooled in Log Analytics for error details. Common causes: SKU allocation failures, insufficient RBAC permissions, or Azure throttling. See Scaling tab."
      QuickWin = $false
    })
  }

  $htmlReport += @"

    <!-- ========== PRIORITY MATRIX ========== -->
    <div class="section" id="sec-priorities">
        <h3 style="margin:0 0 20px 0;font-size:16px;color:#333">Recommendations Priority Matrix</h3>
        
        <!-- 2x2 Grid -->
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:2px;background:#ddd;border-radius:8px;overflow:hidden;margin-bottom:24px;">
            <!-- High Impact / Low Effort = DO FIRST -->
            <div style="background:#e8f5e9;padding:20px;">
                <div style="font-weight:700;color:#2e7d32;font-size:14px;margin-bottom:12px">🎯 DO FIRST — High Impact, Low Effort</div>
$(
  $q1 = @($priorityItems | Where-Object { $_.Impact -eq "High" -and $_.Effort -eq "Low" })
  if ($q1.Count -gt 0) {
    foreach ($item in $q1) {
      "                <div style='padding:8px 0;border-bottom:1px solid #c8e6c9;font-size:13px'><span class='badge b-green'>$($item.Category)</span> $($item.Finding)<br><span style='color:#666;font-size:11px'>→ $($item.Action)</span></div>"
    }
  } else {
    "                <div style='color:#999;font-size:13px'>No items in this quadrant ✓</div>"
  }
)
            </div>
            <!-- High Impact / High Effort = PLAN -->
            <div style="background:#fff3e0;padding:20px;">
                <div style="font-weight:700;color:#e65100;font-size:14px;margin-bottom:12px">📋 PLAN — High Impact, High Effort</div>
$(
  $q2 = @($priorityItems | Where-Object { $_.Impact -eq "High" -and ($_.Effort -eq "Medium" -or $_.Effort -eq "High") })
  if ($q2.Count -gt 0) {
    foreach ($item in $q2) {
      "                <div style='padding:8px 0;border-bottom:1px solid #ffe0b2;font-size:13px'><span class='badge b-yellow'>$($item.Category)</span> $($item.Finding)<br><span style='color:#666;font-size:11px'>→ $($item.Action)</span></div>"
    }
  } else {
    "                <div style='color:#999;font-size:13px'>No items in this quadrant ✓</div>"
  }
)
            </div>
            <!-- Medium/Low Impact / Low Effort = QUICK WINS -->
            <div style="background:#e3f2fd;padding:20px;">
                <div style="font-weight:700;color:#0d47a1;font-size:14px;margin-bottom:12px">⚡ QUICK WINS — Moderate Impact, Low Effort</div>
$(
  $q3 = @($priorityItems | Where-Object { ($_.Impact -eq "Medium" -or $_.Impact -eq "Low") -and $_.Effort -eq "Low" })
  if ($q3.Count -gt 0) {
    foreach ($item in $q3) {
      "                <div style='padding:8px 0;border-bottom:1px solid #bbdefb;font-size:13px'><span class='badge b-blue'>$($item.Category)</span> $($item.Finding)<br><span style='color:#666;font-size:11px'>→ $($item.Action)</span></div>"
    }
  } else {
    "                <div style='color:#999;font-size:13px'>No items in this quadrant ✓</div>"
  }
)
            </div>
            <!-- Medium/Low Impact / High Effort = CONSIDER -->
            <div style="background:#fafafa;padding:20px;">
                <div style="font-weight:700;color:#616161;font-size:14px;margin-bottom:12px">📝 CONSIDER — Moderate Impact, More Effort</div>
$(
  $q4 = @($priorityItems | Where-Object { ($_.Impact -eq "Medium" -or $_.Impact -eq "Low") -and ($_.Effort -eq "Medium" -or $_.Effort -eq "High") })
  if ($q4.Count -gt 0) {
    foreach ($item in $q4) {
      "                <div style='padding:8px 0;border-bottom:1px solid #e0e0e0;font-size:13px'><span class='badge b-gray'>$($item.Category)</span> $($item.Finding)<br><span style='color:#666;font-size:11px'>→ $($item.Action)</span></div>"
    }
  } else {
    "                <div style='color:#999;font-size:13px'>No items in this quadrant ✓</div>"
  }
)
            </div>
        </div>

        <!-- Summary table -->
        <div class="table-wrap">
            <div class="table-title">All Findings — Sorted by Priority</div>
            <div class="table-scroll">
                <table id="priority-table">
                    <thead><tr>
                        <th onclick="sortTable('priority-table',0)">Category</th>
                        <th>Finding</th>
                        <th onclick="sortTable('priority-table',2)">Impact</th>
                        <th onclick="sortTable('priority-table',3)">Effort</th>
                        <th>Recommended Action</th>
                        <th onclick="sortTable('priority-table',5)">Quick Win?</th>
                    </tr></thead>
                    <tbody>
"@

  # Sort: quick wins first, then by impact desc
  $sortedPriority = @($priorityItems | Sort-Object @{Expression={$_.QuickWin};Descending=$true}, @{Expression={switch($_.Impact){"High"{3}"Medium"{2}default{1}}};Descending=$true})
  
  foreach ($pi in $sortedPriority) {
    $impactBadge = switch ($pi.Impact) { "High" { "b-red" }; "Medium" { "b-yellow" }; default { "b-gray" } }
    $effortBadge = switch ($pi.Effort) { "Low" { "b-green" }; "Medium" { "b-yellow" }; default { "b-red" } }
    
    $htmlReport += @"
                    <tr>
                        <td><span class="badge b-blue">$($pi.Category)</span></td>
                        <td>$($pi.Finding)</td>
                        <td><span class="badge $impactBadge">$($pi.Impact)</span></td>
                        <td><span class="badge $effortBadge">$($pi.Effort)</span></td>
                        <td>$($pi.Action)</td>
                        <td>$(if ($pi.QuickWin) { '<span class="badge b-green">Yes</span>' } else { '—' })</td>
                    </tr>
"@
  }

  if ($sortedPriority.Count -eq 0) {
    $htmlReport += @"
                    <tr><td colspan="6" style="text-align:center;color:#2e7d32;padding:24px">✅ No findings — environment is in good shape across all checks!</td></tr>
"@
  }

  $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
    </div>
"@

  # ========== CONNECTION & LOGINS (KQL Data) ==========
  if ($IncludeReservationAnalysis -and (SafeCount $reservationAnalysis) -gt 0) {
    $htmlReport += @"
    
    <!-- ========== RESERVATIONS ========== -->
    <div class="section" id="sec-reservations">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Uncovered VMs</div>
                <div class="card-value $(if ($totalUncoveredVMs -gt 0) { 'orange' } else { 'green' })">$totalUncoveredVMs</div>
                <div class="card-sub">paying PAYG rates</div>
            </div>
            <div class="card">
                <div class="card-label">Potential Savings (1yr RI)</div>
                <div class="card-value green">~`$$([math]::Round($totalRI1ySavings, 0))/mo</div>
                <div class="card-sub">`$$([math]::Round($totalRI1ySavings * 12, 0))/yr</div>
            </div>
            <div class="card">
                <div class="card-label">Potential Savings (3yr RI)</div>
                <div class="card-value green">~`$$([math]::Round($totalRI3ySavings, 0))/mo</div>
                <div class="card-sub">`$$([math]::Round($totalRI3ySavings * 12, 0))/yr</div>
            </div>
            <div class="card">
                <div class="card-label">Existing Reservations</div>
                <div class="card-value blue">$(SafeCount $existingReservations)</div>
                <div class="card-sub">$(if ($totalOverProvisioned -gt 0) { "$totalOverProvisioned over-provisioned" } else { "none over-provisioned" })</div>
            </div>
        </div>
"@

    # Expiring reservations warning
    $expiringRIcount = @($existingReservations | Where-Object { $_.DaysUntilExpiry -ne "Unknown" -and [int]$_.DaysUntilExpiry -le 90 }).Count
    if ($expiringRIcount -gt 0) {
      $htmlReport += @"
        <div class="alert alert-danger">
            <strong>⚠️ $expiringRIcount reservation(s) expiring within 90 days.</strong> Review and renew to avoid reverting to PAYG rates.
        </div>
"@
    }

    $htmlReport += @"
        <div class="table-wrap">
            <div class="table-title">RI Coverage by SKU &amp; Region</div>
            <div class="table-scroll">
                <table id="ri-table">
                    <thead><tr>
                        <th onclick="sortTable('ri-table',0)">VM Size</th>
                        <th onclick="sortTable('ri-table',1)">Region</th>
                        <th onclick="sortTable('ri-table',2)">Deployed</th>
                        <th onclick="sortTable('ri-table',3)">Reserved</th>
                        <th onclick="sortTable('ri-table',4)">Uncovered</th>
                        <th onclick="sortTable('ri-table',5)">Status</th>
                        <th onclick="sortTable('ri-table',6)">PAYG /mo</th>
                        <th onclick="sortTable('ri-table',7)">1yr RI /mo</th>
                        <th onclick="sortTable('ri-table',8)">3yr RI /mo</th>
                        <th onclick="sortTable('ri-table',9)">Savings (3yr)</th>
                        <th onclick="sortTable('ri-table',10)">Priority</th>
                    </tr></thead>
                    <tbody>
"@

    foreach ($ra in ($reservationAnalysis | Sort-Object @{Expression={$_.TotalMonthlySavings3Y};Descending=$true})) {
      $statusBadge = switch ($ra.Status) {
        "Fully Covered"          { "b-green" }
        "Over-Provisioned"       { "b-yellow" }
        "Uncovered (Always-On)"  { "b-red" }
        "Uncovered"              { "b-red" }
        default                  { "b-gray" }
      }
      $priBadge = switch ($ra.Priority) { "High" { "b-red" }; "Medium" { "b-yellow" }; "Low" { "b-gray" }; default { "b-green" } }
      $savingsDisplay = if ($ra.TotalMonthlySavings3Y -gt 0) { "<span style='color:#2e7d32;font-weight:600'>~`$$($ra.TotalMonthlySavings3Y)/mo</span>" } else { "—" }
      
      $htmlReport += @"
                    <tr onclick="toggleDetail(this)" style="cursor:pointer">
                        <td><strong>$($ra.VMSize)</strong><br><span style="font-size:11px;color:#999">$($ra.HostPools)</span></td>
                        <td>$($ra.Region)</td>
                        <td>$($ra.DeployedVMs)</td>
                        <td>$($ra.ReservedQty)</td>
                        <td style="$(if ($ra.UncoveredVMs -gt 0) { 'color:#c62828;font-weight:600' })">$($ra.UncoveredVMs)</td>
                        <td><span class="badge $statusBadge">$($ra.Status)</span></td>
                        <td>`$$($ra.PAYGMonthlyPerVM)</td>
                        <td>`$$($ra.RI1YMonthlyPerVM)</td>
                        <td>`$$($ra.RI3YMonthlyPerVM)</td>
                        <td>$savingsDisplay</td>
                        <td><span class="badge $priBadge">$($ra.Priority)</span></td>
                    </tr>
                    <tr class="detail-row" style="display:none">
                        <td colspan="11" style="background:#f8fafc;padding:12px 20px;font-size:12px;border-left:3px solid #0078d4">
                            <strong>Recommendation:</strong> $($ra.Recommendation)$(if ($ra.ExpiryWarning) { "<br><br>$($ra.ExpiryWarning)" })
                            <br><br><strong>Per-VM comparison:</strong> PAYG `$$($ra.PAYGMonthlyPerVM)/mo → 1yr RI `$$($ra.RI1YMonthlyPerVM)/mo ($($ra.Savings1YPct) off) → 3yr RI `$$($ra.RI3YMonthlyPerVM)/mo ($($ra.Savings3YPct) off)
                            $(if ($ra.IsAlwaysOn) { "<br><strong>⚡ Personal desktop (always-on)</strong> — highest RI ROI" })
                        </td>
                    </tr>
"@
    }

    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@

    # Existing reservations table — filtered to AVD-relevant SKUs
    if ($existingReservations.Count -gt 0) {
      # Build set of AVD VM SKUs + regions for relevance matching
      $avdSkuSet = @{}
      foreach ($v in $vms) {
        if ($v.VMSize -and $v.Region) { $avdSkuSet["$($v.VMSize)|$($v.Region)".ToLower()] = $true }
        if ($v.VMSize) { $avdSkuSet["$($v.VMSize)".ToLower()] = $true }  # SKU-only match for shared-scope RIs
      }
      
      # Filter: show RIs that match AVD VM SKUs (by name)
      $avdRelevantRIs = @($existingReservations | Where-Object {
        $riSku = $_.SKU
        $riLoc = $_.Location
        $avdSkuSet.ContainsKey("$riSku|$riLoc".ToLower()) -or 
        $avdSkuSet.ContainsKey("$riSku".ToLower())
      })
      $nonAvdCount = $existingReservations.Count - $avdRelevantRIs.Count
      
      if ($avdRelevantRIs.Count -gt 0) {
        $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">Existing Reservations — AVD Relevant ($($avdRelevantRIs.Count) of $($existingReservations.Count))</h3>
$(if ($nonAvdCount -gt 0) { "        <p style='font-size:12px;color:#666;margin:-8px 0 12px 0'>$nonAvdCount non-AVD reservation(s) filtered out (SKU not matching any AVD session host)</p>" })
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="existing-ri-table">
                    <thead><tr>
                        <th onclick="sortTable('existing-ri-table',0)">Reservation Name</th>
                        <th onclick="sortTable('existing-ri-table',1)">SKU</th>
                        <th onclick="sortTable('existing-ri-table',2)">Location</th>
                        <th onclick="sortTable('existing-ri-table',3)">Qty</th>
                        <th onclick="sortTable('existing-ri-table',4)">AVD VMs</th>
                        <th onclick="sortTable('existing-ri-table',5)">Coverage</th>
                        <th onclick="sortTable('existing-ri-table',6)">Term</th>
                        <th onclick="sortTable('existing-ri-table',7)">Status</th>
                        <th onclick="sortTable('existing-ri-table',8)">Expiry</th>
                        <th onclick="sortTable('existing-ri-table',9)">Days Left</th>
                    </tr></thead>
                    <tbody>
"@
        foreach ($eri in ($avdRelevantRIs | Sort-Object DaysUntilExpiry)) {
          $expiryStyle = ""
          if ($eri.DaysUntilExpiry -ne "Unknown" -and [int]$eri.DaysUntilExpiry -le 90) { $expiryStyle = "color:#c62828;font-weight:600" }
          elseif ($eri.DaysUntilExpiry -ne "Unknown" -and [int]$eri.DaysUntilExpiry -le 180) { $expiryStyle = "color:#e65100;font-weight:600" }
          $statusBadge = if ($eri.Status -eq "Active") { "b-green" } else { "b-yellow" }
          
          # Count AVD VMs this RI covers
          $matchingAvdVms = @($vms | Where-Object { $_.VMSize -eq $eri.SKU -and $_.Region -eq $eri.Location }).Count
          if ($matchingAvdVms -eq 0) { $matchingAvdVms = @($vms | Where-Object { $_.VMSize -eq $eri.SKU }).Count }
          $coverageBadge = if ($eri.Quantity -ge $matchingAvdVms -and $matchingAvdVms -gt 0) { "<span class='badge b-green'>Covered</span>" }
                           elseif ($eri.Quantity -gt 0 -and $matchingAvdVms -gt 0) { "<span class='badge b-orange'>Partial ($($eri.Quantity)/$matchingAvdVms)</span>" }
                           else { "<span class='badge b-gray'>—</span>" }
          
          $htmlReport += @"
                    <tr>
                        <td><strong>$($eri.ReservationName)</strong></td>
                        <td>$($eri.SKU)</td>
                        <td>$($eri.Location)</td>
                        <td>$($eri.Quantity)</td>
                        <td>$matchingAvdVms</td>
                        <td>$coverageBadge</td>
                        <td>$($eri.Term)</td>
                        <td><span class="badge $statusBadge">$($eri.Status)</span></td>
                        <td>$(if ($eri.ExpiryDate) { $eri.ExpiryDate.ToString('yyyy-MM-dd') } else { 'N/A' })</td>
                        <td style="$expiryStyle">$($eri.DaysUntilExpiry)</td>
                    </tr>
"@
        }
        $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
      } elseif ($existingReservations.Count -gt 0) {
        $htmlReport += @"
        <div class="alert alert-info"><strong>ℹ️</strong> Found $($existingReservations.Count) existing reservation(s) in the tenant, but none match AVD session host SKUs. All AVD VMs are running at PAYG or Savings Plan rates.</div>
"@
      }
    }

    $htmlReport += @"
        <div class="alert alert-warning">
            <strong>⚠️ Note:</strong> RI pricing shown is estimated East US retail. Actual savings vary by region, EA/CSP agreements, and negotiated rates. Always validate with Azure Cost Management before purchasing.
$(if (-not $hasAzReservations) { "            <br><br><strong>ℹ️</strong> Existing reservations could not be checked — install <code>Az.Reservations</code> module and ensure Reservations Reader role for full coverage analysis." })
        </div>
    </div>
"@
  }

  # ========== CONNECTION & LOGINS (KQL Data) ==========
  if ($hasKqlData) {
    $htmlReport += @"
    
    <div class="section" id="sec-kql">
$(if ($kqlNoData.Count -gt 0) {
  "        <div class='alert alert-warning'><strong>Limited Data:</strong> The following queries returned no rows: $($kqlNoData -join ', '). Try <code>-MetricsLookbackDays 30</code> for more data.</div>"
})
$(if ($kqlFailed.Count -gt 0) {
  "        <div class='alert alert-warning'><strong>Note:</strong> These diagnostic tables are not enabled: $($kqlFailed -join ', '). Enable them in the host pool diagnostic settings to collect this data.</div>"
})
"@

    # --- Login Time Analysis (v4.1) ---
    if ($loginTimeData.Count -gt 0) {
      $htmlReport += @"
        <h3 style="margin:0 0 16px 0;font-size:16px;color:#333">🔑 Login Time — Time to Desktop</h3>
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Pools Measured</div>
                <div class="card-value blue">$($loginTimeData.Count)</div>
            </div>
            <div class="card">
                <div class="card-label">Avg Login Time</div>
                <div class="card-value $(if (($loginTimeData | ForEach-Object { [double]$_.AvgLoginSec } | Measure-Object -Average).Average -gt 45) { 'orange' } else { 'green' })">$([math]::Round(($loginTimeData | ForEach-Object { [double]$_.AvgLoginSec } | Measure-Object -Average).Average, 0))s</div>
            </div>
            <div class="card">
                <div class="card-label">P95 Login Time</div>
                <div class="card-value $(if (($loginTimeData | ForEach-Object { [double]$_.P95LoginSec } | Measure-Object -Average).Average -gt 60) { 'red' } elseif (($loginTimeData | ForEach-Object { [double]$_.P95LoginSec } | Measure-Object -Average).Average -gt 30) { 'orange' } else { 'green' })">$([math]::Round(($loginTimeData | ForEach-Object { [double]$_.P95LoginSec } | Measure-Object -Average).Average, 0))s</div>
            </div>
            <div class="card">
                <div class="card-label">Worst Pool P95</div>
                <div class="card-value $(if (($loginTimeData | ForEach-Object { [double]$_.P95LoginSec } | Measure-Object -Maximum).Maximum -gt 90) { 'red' } else { 'orange' })">$([math]::Round(($loginTimeData | ForEach-Object { [double]$_.P95LoginSec } | Measure-Object -Maximum).Maximum, 0))s</div>
            </div>
        </div>
        <div class="table-wrap">
            <div class="table-title">Login Time by Host Pool (session start to desktop ready)</div>
            <div class="table-scroll">
                <table id="login-time-table">
                    <thead><tr>
                        <th onclick="sortTable('login-time-table',0)">Host Pool</th>
                        <th onclick="sortTable('login-time-table',1)">Avg (s)</th>
                        <th onclick="sortTable('login-time-table',2)">Median (s)</th>
                        <th onclick="sortTable('login-time-table',3)">P95 (s)</th>
                        <th onclick="sortTable('login-time-table',4)">Max (s)</th>
                        <th onclick="sortTable('login-time-table',5)">Connections</th>
                        <th>Rating</th>
                    </tr></thead>
                    <tbody>
"@
      foreach ($lt in ($loginTimeData | Sort-Object { [double]$_.P95LoginSec } -Descending)) {
        $ltP95 = [double]$lt.P95LoginSec
        $ltRating = if ($ltP95 -le 15) { "<span class='badge b-green'>Excellent</span>" }
                    elseif ($ltP95 -le 30) { "<span class='badge b-green'>Good</span>" }
                    elseif ($ltP95 -le 60) { "<span class='badge b-orange'>Fair</span>" }
                    else { "<span class='badge b-red'>Poor</span>" }
        $ltStyle = if ($ltP95 -gt 60) { "color:#c62828;font-weight:600" } elseif ($ltP95 -gt 30) { "color:#e65100" } else { "" }
        $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $lt.HostPool)</strong></td>
                        <td>$($lt.AvgLoginSec)</td>
                        <td>$($lt.P50LoginSec)</td>
                        <td style="$ltStyle">$($lt.P95LoginSec)</td>
                        <td>$($lt.MaxLoginSec)</td>
                        <td>$($lt.TotalConnections)</td>
                        <td>$ltRating</td>
                    </tr>
"@
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # --- Connection Success Rate (v4.1) ---
    if ($connSuccessData.Count -gt 0) {
      $totalAttempts = ($connSuccessData | ForEach-Object { [int]$_.TotalAttempts } | Measure-Object -Sum).Sum
      $totalFailed = ($connSuccessData | ForEach-Object { [int]$_.Failed } | Measure-Object -Sum).Sum
      $overallSuccessRate = if ($totalAttempts -gt 0) { [math]::Round(($totalAttempts - $totalFailed) / $totalAttempts * 100, 1) } else { 0 }
      $problematicPools = @($connSuccessData | Where-Object { [double]$_.FailureRate -gt 5 })
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">📊 Connection Success Rate</h3>
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Overall Success Rate</div>
                <div class="card-value $(if ($overallSuccessRate -ge 98) { 'green' } elseif ($overallSuccessRate -ge 95) { 'orange' } else { 'red' })">$overallSuccessRate%</div>
            </div>
            <div class="card">
                <div class="card-label">Total Connections</div>
                <div class="card-value blue">$totalAttempts</div>
            </div>
            <div class="card">
                <div class="card-label">Failed Connections</div>
                <div class="card-value $(if ($totalFailed -gt 0) { 'red' } else { 'green' })">$totalFailed</div>
            </div>
            <div class="card">
                <div class="card-label">Pools &gt;5% Failure</div>
                <div class="card-value $(if ($problematicPools.Count -gt 0) { 'red' } else { 'green' })">$($problematicPools.Count)</div>
            </div>
        </div>
$(if ($problematicPools.Count -gt 0) {
  $poolNames = ($problematicPools | ForEach-Object { "$(Scrub-HostPoolName $_.HostPool) ($($_.FailureRate)%)" }) -join ", "
  "        <div class='alert alert-warning'><strong>⚠️ High failure rate pools:</strong> $poolNames — investigate connection errors on these pools.</div>"
})
        <div class="table-wrap">
            <div class="table-title">Connection Success Rate by Pool</div>
            <div class="table-scroll">
                <table id="conn-rate-table">
                    <thead><tr>
                        <th onclick="sortTable('conn-rate-table',0)">Host Pool</th>
                        <th onclick="sortTable('conn-rate-table',1)">Attempts</th>
                        <th onclick="sortTable('conn-rate-table',2)">Succeeded</th>
                        <th onclick="sortTable('conn-rate-table',3)">Failed</th>
                        <th onclick="sortTable('conn-rate-table',4)">Success %</th>
                        <th onclick="sortTable('conn-rate-table',5)">Failure %</th>
                        <th>Status</th>
                    </tr></thead>
                    <tbody>
"@
      foreach ($cr in ($connSuccessData | Sort-Object { [double]$_.FailureRate } -Descending)) {
        $crStatus = if ([double]$cr.FailureRate -le 1) { "<span class='badge b-green'>Healthy</span>" }
                    elseif ([double]$cr.FailureRate -le 5) { "<span class='badge b-orange'>Monitor</span>" }
                    else { "<span class='badge b-red'>Investigate</span>" }
        $crStyle = if ([double]$cr.FailureRate -gt 5) { "color:#c62828;font-weight:600" } elseif ([double]$cr.FailureRate -gt 1) { "color:#e65100" } else { "" }
        $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $cr.HostPool)</strong></td>
                        <td>$($cr.TotalAttempts)</td>
                        <td>$($cr.Succeeded)</td>
                        <td style="$crStyle">$($cr.Failed)</td>
                        <td>$($cr.SuccessRate)%</td>
                        <td style="$crStyle">$($cr.FailureRate)%</td>
                        <td>$crStatus</td>
                    </tr>
"@
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # --- Drain Mode Alert (v4.1) ---
    if ($drainModeHosts.Count -gt 0) {
      $htmlReport += @"
        <div class="alert alert-warning" style="margin-top:16px">
            <strong>🚫 Drain Mode Hosts: $($drainModeHosts.Count) session host(s) are in drain mode</strong> (AllowNewSession = false) across $($drainByPool.Count) pool(s).
            Users cannot be routed to these hosts. Effective capacity is reduced.
            <div style="margin-top:8px;font-size:13px">
$(foreach ($dp in $drainByPool.GetEnumerator()) {
  $totalInPool = @($sessionHosts | Where-Object { $_.HostPoolName -eq $dp.Key }).Count
  $effective = $totalInPool - $dp.Value
  "                <div>$(Scrub-HostPoolName $dp.Key): <strong>$($dp.Value)</strong> of $totalInPool hosts in drain → effective capacity: <strong>$effective hosts</strong></div>`n"
})
            </div>
        </div>
"@
    }

    # --- Profile Load Performance ---
    if ($profileLoadData.Count -gt 0) {
      $htmlReport += @"
        <h3 style="margin:0 0 16px 0;font-size:16px;color:#333">Profile Load Performance</h3>
        <div class="table-wrap">
            <div class="table-title">Session Establishment Latency by Host (slowest first)</div>
            <div class="table-scroll">
                <table id="profile-table">
                    <thead><tr>
                        <th onclick="sortTable('profile-table',0)">Session Host</th>
                        <th onclick="sortTable('profile-table',1)">Avg (s)</th>
                        <th onclick="sortTable('profile-table',2)">P50 (s)</th>
                        <th onclick="sortTable('profile-table',3)">P95 (s)</th>
                        <th onclick="sortTable('profile-table',4)">Max (s)</th>
                        <th onclick="sortTable('profile-table',5)">Sessions</th>
                        <th onclick="sortTable('profile-table',6)">Slow (&gt;30s)</th>
                        <th onclick="sortTable('profile-table',7)">Slow %</th>
                    </tr></thead>
                    <tbody>
"@
      foreach ($pl in ($profileLoadData | Sort-Object { [double]($_.P95ProfileLoadSec) } -Descending | Select-Object -First 30)) {
        $p95val = [double]$pl.P95ProfileLoadSec
        $p95color = if ($p95val -gt 60) { "color:#c62828;font-weight:600" } elseif ($p95val -gt 30) { "color:#e65100;font-weight:600" } else { "" }
        $slowPct = if ($pl.PSObject.Properties['SlowLoginPct'] -and $pl.SlowLoginPct) { "$($pl.SlowLoginPct)%" } else { "0%" }
        
        $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-VMName $pl.SessionHostName)</strong></td>
                        <td>$($pl.AvgProfileLoadSec)</td>
                        <td>$($pl.P50ProfileLoadSec)</td>
                        <td style="$p95color">$($pl.P95ProfileLoadSec)</td>
                        <td>$($pl.MaxProfileLoadSec)</td>
                        <td>$($pl.TotalSessions)</td>
                        <td>$($pl.SlowLogins_Over30s)</td>
                        <td>$slowPct</td>
                    </tr>
"@
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # --- Session Duration Analysis ---
    $sessionDurationData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_SessionDuration" -and $_.QueryName -eq "AVD" -and $_.PSObject.Properties.Name -contains "UserName" -and $_.AvgDuration })
    if ($sessionDurationData.Count -gt 0) {
      # Build duration distribution buckets (by user average)
      $durBuckets = [ordered]@{
        "< 30 min"   = 0
        "30m - 1h"   = 0
        "1 - 2h"     = 0
        "2 - 4h"     = 0
        "4 - 6h"     = 0
        "6 - 8h"     = 0
        "8h+"        = 0
      }
      $allAvgDurations = @()
      foreach ($sd in $sessionDurationData) {
        $dur = [double]$sd.AvgDuration
        $allAvgDurations += $dur
        if ($dur -lt 30)      { $durBuckets["< 30 min"]++ }
        elseif ($dur -lt 60)  { $durBuckets["30m - 1h"]++ }
        elseif ($dur -lt 120) { $durBuckets["1 - 2h"]++ }
        elseif ($dur -lt 240) { $durBuckets["2 - 4h"]++ }
        elseif ($dur -lt 360) { $durBuckets["4 - 6h"]++ }
        elseif ($dur -lt 480) { $durBuckets["6 - 8h"]++ }
        else                  { $durBuckets["8h+"]++ }
      }

      $overallAvgMin = [math]::Round(($allAvgDurations | Measure-Object -Average).Average, 0)
      $overallMedian = [math]::Round(($allAvgDurations | Sort-Object)[[math]::Floor($allAvgDurations.Count / 2)], 0)
      $overallMaxMin = [math]::Round(($sessionDurationData | ForEach-Object { [double]$_.MaxDuration } | Measure-Object -Maximum).Maximum, 0)
      $totalUsers = $sessionDurationData.Count
      $shortSessionUsers = @($sessionDurationData | Where-Object { [double]$_.AvgDuration -lt 120 }).Count
      $fullDayUsers = @($sessionDurationData | Where-Object { [double]$_.AvgDuration -ge 360 }).Count
      $shortPct = if ($totalUsers -gt 0) { [math]::Round($shortSessionUsers / $totalUsers * 100, 0) } else { 0 }
      $fullDayPct = if ($totalUsers -gt 0) { [math]::Round($fullDayUsers / $totalUsers * 100, 0) } else { 0 }

      # Format avg as hours:minutes
      $avgHrs = [math]::Floor($overallAvgMin / 60)
      $avgMins = $overallAvgMin % 60
      $avgDisplay = if ($avgHrs -gt 0) { "${avgHrs}h ${avgMins}m" } else { "${avgMins}m" }
      $medHrs = [math]::Floor($overallMedian / 60)
      $medMins = $overallMedian % 60
      $medDisplay = if ($medHrs -gt 0) { "${medHrs}h ${medMins}m" } else { "${medMins}m" }

      # Licensing insight
      $licensingInsight = if ($shortPct -ge 60) {
        "<span style='color:#1565c0;font-weight:600'>$shortPct% of users average under 2 hours</span> — strong indicator for W365 Frontline (Shared mode) or aggressive autoscale with short ramp-down timers. These users don't need a dedicated VM running all day."
      } elseif ($shortPct -ge 30) {
        "<span style='color:#e65100;font-weight:600'>$shortPct% of users average under 2 hours</span> — a mixed fleet with both short-session and full-day users. Consider segmenting into separate pools: short-session users on Frontline/pooled with aggressive scaling, full-day users on dedicated/personal."
      } elseif ($fullDayPct -ge 60) {
        "<span style='color:#2e7d32'>$fullDayPct% of users average 6+ hour sessions</span> — full-day knowledge workers. Personal desktops or dedicated pooled with conservative scaling are appropriate. W365 Enterprise is viable for this pattern."
      } else {
        "Mixed session patterns — $shortPct% under 2 hours, $fullDayPct% over 6 hours. Consider pool segmentation by usage intensity."
      }

      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">⏱️ Session Duration Analysis</h3>
        <div class="card-grid" style="margin-bottom:16px">
            <div class="card">
                <div class="card-label">Active Users</div>
                <div class="card-value blue">$totalUsers</div>
                <div class="card-sub">with session data</div>
            </div>
            <div class="card">
                <div class="card-label">Avg Session</div>
                <div class="card-value">$avgDisplay</div>
                <div class="card-sub">median: $medDisplay</div>
            </div>
            <div class="card">
                <div class="card-label">&lt; 2 Hours</div>
                <div class="card-value $(if ($shortPct -ge 50) { 'orange' } else { 'blue' })">$shortPct%</div>
                <div class="card-sub">$shortSessionUsers users</div>
            </div>
            <div class="card">
                <div class="card-label">6+ Hours</div>
                <div class="card-value $(if ($fullDayPct -ge 50) { 'green' } else { 'blue' })">$fullDayPct%</div>
                <div class="card-sub">$fullDayUsers full-day users</div>
            </div>
        </div>

        <div class="two-col">
            <div class="table-wrap">
                <div class="table-title">User Session Duration Distribution</div>
                <div style="padding:20px;">
$(
  $maxBucket = ($durBuckets.Values | Measure-Object -Maximum).Maximum
  foreach ($bucket in $durBuckets.GetEnumerator()) {
    $pct = if ($totalUsers -gt 0) { [math]::Round($bucket.Value / $totalUsers * 100, 0) } else { 0 }
    $barW = if ($maxBucket -gt 0) { [math]::Round($bucket.Value / $maxBucket * 100, 0) } else { 0 }
    $barColor = switch ($bucket.Key) {
      "< 30 min" { "#ef5350" }
      "30m - 1h" { "#ff7043" }
      "1 - 2h"   { "#ffa726" }
      "2 - 4h"   { "#66bb6a" }
      "4 - 6h"   { "#42a5f5" }
      "6 - 8h"   { "#5c6bc0" }
      "8h+"      { "#7e57c2" }
    }
    "                    <div style='display:flex;align-items:center;gap:8px;margin:4px 0'>"
    "                        <span style='width:70px;font-size:12px;text-align:right;color:#666'>$($bucket.Key)</span>"
    "                        <div style='flex:1;background:#f0f0f0;border-radius:4px;height:22px'>"
    "                            <div style='width:${barW}%;background:${barColor};height:100%;border-radius:4px;min-width:2px'></div>"
    "                        </div>"
    "                        <span style='width:60px;font-size:12px;color:#333'>$($bucket.Value) ($pct%)</span>"
    "                    </div>"
  }
)
                </div>
            </div>
            <div class="table-wrap">
                <div class="table-title">💡 Licensing Insight</div>
                <div style="padding:16px 20px;font-size:13px;line-height:1.7;color:#333">
                    <p style="margin:0 0 12px 0">$licensingInsight</p>
                    <div style="background:#f8f9fa;border-radius:4px;padding:12px;font-size:12px;color:#555">
                        <div style="margin-bottom:4px">Users with avg &lt; 2h are shift/occasional workers who benefit from shared or Frontline licensing.</div>
                        <div style="margin-bottom:4px">Users with avg 2-6h are standard knowledge workers — pooled desktops with moderate scaling.</div>
                        <div>Users with avg 6h+ are power/full-day users — best served by personal desktops or dedicated pools.</div>
                    </div>
                </div>
            </div>
        </div>

        <div class="table-wrap" style="margin-top:16px">
            <div class="table-title">Top 20 Longest Average Sessions</div>
            <div class="table-scroll">
                <table id="session-dur-table">
                    <thead><tr>
                        <th onclick="sortTable('session-dur-table',0)">User</th>
                        <th onclick="sortTable('session-dur-table',1)">Avg Duration</th>
                        <th onclick="sortTable('session-dur-table',2)">Max Duration</th>
                        <th onclick="sortTable('session-dur-table',3)">Pattern</th>
                    </tr></thead>
                    <tbody>
"@
      $topUsers = $sessionDurationData | Sort-Object { [double]$_.AvgDuration } -Descending | Select-Object -First 20
      foreach ($su in $topUsers) {
        $suAvg = [double]$su.AvgDuration
        $suMax = [double]$su.MaxDuration
        $suAvgH = [math]::Floor($suAvg / 60)
        $suAvgM = [math]::Round($suAvg % 60, 0)
        $suMaxH = [math]::Floor($suMax / 60)
        $suMaxM = [math]::Round($suMax % 60, 0)
        $suAvgStr = if ($suAvgH -gt 0) { "${suAvgH}h ${suAvgM}m" } else { "${suAvgM}m" }
        $suMaxStr = if ($suMaxH -gt 0) { "${suMaxH}h ${suMaxM}m" } else { "${suMaxM}m" }
        $pattern = if ($suAvg -ge 360) { "<span class='badge' style='background:#e8eaf6;color:#3949ab'>Full Day</span>" }
                   elseif ($suAvg -ge 120) { "<span class='badge' style='background:#e8f5e9;color:#2e7d32'>Standard</span>" }
                   elseif ($suAvg -ge 30) { "<span class='badge' style='background:#fff3e0;color:#e65100'>Short Session</span>" }
                   else { "<span class='badge' style='background:#ffebee;color:#c62828'>Micro Session</span>" }
        # Mask username for privacy — show first part only
        $maskedUser = if ($ScrubPII) { Scrub-Username $su.UserName } elseif ($su.UserName -match '^([^@]+)@') { $matches[1] } else { $su.UserName }
        $htmlReport += @"
                    <tr>
                        <td>$maskedUser</td>
                        <td>$suAvgStr</td>
                        <td>$suMaxStr</td>
                        <td>$pattern</td>
                    </tr>
"@
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # --- Connection Quality by Client OS ---
    if ($connQualityData.Count -gt 0) {
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">Connection Quality by Client OS</h3>
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="connq-table">
                    <thead><tr>
                        <th onclick="sortTable('connq-table',0)">Client OS</th>
                        <th onclick="sortTable('connq-table',1)">Avg RTT (ms)</th>
                        <th onclick="sortTable('connq-table',2)">P95 RTT (ms)</th>
                        <th onclick="sortTable('connq-table',3)">Max RTT (ms)</th>
                        <th onclick="sortTable('connq-table',4)">Avg BW (KBps)</th>
                        <th onclick="sortTable('connq-table',5)">Min BW (KBps)</th>
                        <th onclick="sortTable('connq-table',6)">Connections</th>
                        <th onclick="sortTable('connq-table',7)">High Latency %</th>
                    </tr></thead>
                    <tbody>
"@
      foreach ($cq in ($connQualityData | Sort-Object { [double]($_.P95RTTms) } -Descending)) {
        $rttColor = if ([double]$cq.P95RTTms -gt 250) { "color:#c62828;font-weight:600" } elseif ([double]$cq.P95RTTms -gt 150) { "color:#e65100;font-weight:600" } else { "" }
        $hlPct = if ($cq.PSObject.Properties['HighLatencyPct'] -and $cq.HighLatencyPct) { "$($cq.HighLatencyPct)%" } else { "0%" }
        
        $htmlReport += @"
                    <tr>
                        <td><strong>$($cq.ClientOS)</strong></td>
                        <td>$($cq.AvgRTTms)</td>
                        <td style="$rttColor">$($cq.P95RTTms)</td>
                        <td>$($cq.MaxRTTms)</td>
                        <td>$($cq.AvgBandwidthKBps)</td>
                        <td>$($cq.MinBandwidthKBps)</td>
                        <td>$($cq.Connections)</td>
                        <td>$hlPct</td>
                    </tr>
"@
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # --- Connection Quality by Region ---
    if ($connQualityByRegionData.Count -gt 0) {
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">Connection Quality by Gateway Region</h3>
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="connr-table">
                    <thead><tr>
                        <th onclick="sortTable('connr-table',0)">Gateway Region</th>
                        <th onclick="sortTable('connr-table',1)">Avg RTT (ms)</th>
                        <th onclick="sortTable('connr-table',2)">P95 RTT (ms)</th>
                        <th onclick="sortTable('connr-table',3)">Avg BW (KBps)</th>
                        <th onclick="sortTable('connr-table',4)">Connections</th>
                        <th onclick="sortTable('connr-table',5)">High Latency %</th>
                    </tr></thead>
                    <tbody>
"@
      foreach ($cr in ($connQualityByRegionData | Sort-Object { [double]($_.P95RTTms) } -Descending)) {
        $rttColor = if ([double]$cr.P95RTTms -gt 250) { "color:#c62828;font-weight:600" } elseif ([double]$cr.P95RTTms -gt 150) { "color:#e65100;font-weight:600" } else { "" }
        
        $htmlReport += @"
                    <tr>
                        <td><strong>$($cr.GatewayRegion)</strong></td>
                        <td>$($cr.AvgRTTms)</td>
                        <td style="$rttColor">$($cr.P95RTTms)</td>
                        <td>$($cr.AvgBandwidthKBps)</td>
                        <td>$($cr.Connections)</td>
                        <td>$(if ($cr.PSObject.Properties['HighLatencyPct'] -and $cr.HighLatencyPct) { "$($cr.HighLatencyPct)%" } else { "0%" })</td>
                    </tr>
"@
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # --- Connection Errors ---
    if ($connErrorData.Count -gt 0) {
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">Top Connection Errors</h3>
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="err-table">
                    <thead><tr>
                        <th onclick="sortTable('err-table',0)">Error Code</th>
                        <th onclick="sortTable('err-table',1)">Count</th>
                        <th onclick="sortTable('err-table',2)">Users Affected</th>
                        <th onclick="sortTable('err-table',3)">Sessions Affected</th>
                        <th>Message</th>
                    </tr></thead>
                    <tbody>
"@
      foreach ($ce in ($connErrorData | Sort-Object { [int]($_.ErrorCount) } -Descending | Select-Object -First 15)) {
        $htmlReport += @"
                    <tr>
                        <td><strong>$($ce.CodeSymbolic)</strong></td>
                        <td style="font-weight:600">$($ce.ErrorCount)</td>
                        <td>$($ce.DistinctUsers)</td>
                        <td>$($ce.DistinctCorrelations)</td>
                        <td style="font-size:11px;max-width:400px;overflow:hidden;text-overflow:ellipsis">$($ce.Message)</td>
                    </tr>
"@
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # --- Disconnect Reason Analysis ---
    if ($disconnectReasonData.Count -gt 0 -or $disconnectData.Count -gt 0 -or $disconnectsByHostData.Count -gt 0) {
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">🔌 Disconnect Analysis</h3>
"@

      # Disconnect Reason Category Breakdown
      if ($disconnectReasonData.Count -gt 0) {
        # Separate normal from abnormal
        $normalCategories = @("Normal Completion", "User Initiated", "No Completion Record", "Idle Timeout", "Auto-Reconnect", "Authentication")
        $abnormalReasons = @($disconnectReasonData | Where-Object { $_.DisconnectCategory -notin $normalCategories })
        $normalReasons = @($disconnectReasonData | Where-Object { $_.DisconnectCategory -in $normalCategories })
        $totalAbnormalSessions = ($abnormalReasons | ForEach-Object { if ($_.PSObject.Properties['SessionCount']) { [int]$_.SessionCount } else { 0 } } | Measure-Object -Sum).Sum
        $totalAllSessions = ($disconnectReasonData | ForEach-Object { if ($_.PSObject.Properties['SessionCount']) { [int]$_.SessionCount } else { 0 } } | Measure-Object -Sum).Sum
        $abnormalPct = if ($totalAllSessions -gt 0) { [math]::Round(($totalAbnormalSessions / $totalAllSessions) * 100, 1) } else { 0 }

        if ($abnormalReasons.Count -gt 0) {
          $topAbnormal = ($abnormalReasons | Sort-Object { [int]$_.SessionCount } -Descending | Select-Object -First 1)
          $htmlReport += @"
        <div class="alert $(if ($abnormalPct -gt 10) { 'alert-danger' } elseif ($abnormalPct -gt 5) { 'alert-warning' } else { 'alert-info' })" style="line-height:1.6">
            <strong>$abnormalPct% of sessions ended abnormally</strong> ($totalAbnormalSessions of $totalAllSessions sessions).
            Top cause: <strong>$($topAbnormal.DisconnectCategory)</strong> ($($topAbnormal.SessionCount) sessions, $($topAbnormal.DistinctUsers) users affected).
        </div>
"@
        }

        $htmlReport += @"
        <div class="table-wrap">
            <div class="table-title">Disconnect Reasons</div>
            <div class="table-scroll">
                <table id="disc-reason-table">
                    <thead><tr>
                        <th onclick="sortTable('disc-reason-table',0)">Category</th>
                        <th onclick="sortTable('disc-reason-table',1)">Sessions</th>
                        <th onclick="sortTable('disc-reason-table',2)">% of Total</th>
                        <th onclick="sortTable('disc-reason-table',3)">Users Affected</th>
                        <th onclick="sortTable('disc-reason-table',4)">Hosts Affected</th>
                        <th>Sample Error</th>
                        <th>Remediation</th>
                    </tr></thead>
                    <tbody>
"@
        foreach ($dr in ($disconnectReasonData | Sort-Object { if ($_.PSObject.Properties['SessionCount']) { [int]$_.SessionCount } else { 0 } } -Descending)) {
          $cat = if ($dr.PSObject.Properties['DisconnectCategory']) { $dr.DisconnectCategory } else { "Unknown" }
          $isNormal = ($cat -in $normalCategories)
          
          # Clean up "Other: <RawCodeSymbolic>" categories for display
          $catDisplay = $cat
          $catTooltip = ""
          if ($cat -match '^Other:\s*(.+)') {
            $rawCode = $matches[1]
            $catTooltip = $rawCode
            # Try to derive a friendly name from the camelCase/PascalCase code
            $friendly = ($rawCode -creplace '([a-z])([A-Z])', '$1 $2' -creplace '([A-Z]+)([A-Z][a-z])', '$1 $2')
            if ($friendly.Length -gt 35) { $friendly = $friendly.Substring(0, 32) + "..." }
            $catDisplay = $friendly
          }
          
          $catBadge = switch -Wildcard ($cat) {
            "Normal Completion" { "<span class='badge b-green'>$cat</span>" }
            "User Initiated"   { "<span class='badge b-green'>$cat</span>" }
            "No Completion*"   { "<span class='badge b-green'>Still Active / In Progress</span>" }
            "Idle Timeout"     { "<span class='badge b-green'>$cat</span>" }
            "Auto-Reconnect"   { "<span class='badge b-green'>$cat</span>" }
            "Network Drop"     { "<span class='badge b-red'>$cat</span>" }
            "Resource*"        { "<span class='badge b-red'>$cat</span>" }
            "Server*"          { "<span class='badge b-red'>$cat</span>" }
            "Agent*"           { "<span class='badge b-red'>$cat</span>" }
            "Authentication"   { "<span class='badge b-orange'>$cat</span>" }
            "Gateway*"         { "<span class='badge b-orange'>$cat</span>" }
            "Licensing*"       { "<span class='badge b-orange'>$cat</span>" }
            "Short Session*"   { "<span class='badge b-yellow'>$cat</span>" }
            "Profile*"         { "<span class='badge b-red'>$cat</span>" }
            "Auto-Reconnect"   { "<span class='badge b-green'>$cat</span>" }
            default            { if ($cat -match '^Other:') { "<span class='badge b-orange' title='$catTooltip'>$catDisplay</span>" } else { "<span class='badge b-gray'>$catDisplay</span>" } }
          }
          $remediation = switch -Wildcard ($cat) {
            "Network Drop"     { "Check VPN stability, NIC health, NSG rules, and UDR routing. Enable RDP Shortpath for UDP resilience." }
            "Idle Timeout"     { "Review session time limits in host pool settings and GPO. Adjust idle disconnect timeout if too aggressive." }
            "Server Side"      { "Check for host reboots (Windows Update, scaling plan ramp-down), VM deallocations, or agent crashes." }
            "Authentication"   { "Verify domain join health, DNS resolution, Kerberos ticket renewal, and NLA/CredSSP settings." }
            "Licensing*"       { "Check RDS CAL server availability and license counts. Review host pool max session limits." }
            "Resource*"        { "Investigate memory pressure and disk I/O on affected hosts. Consider upsizing or reducing session density." }
            "Agent*"           { "Check RDAgent and RDAgentBootLoader services. Reinstall agent if persistent." }
            "Gateway*"         { "Usually transient. If recurring, check private endpoints and network path to AVD control plane." }
            "Short Session*"   { "Sessions <1 min with no error code. May indicate app crashes, profile load failures, or GPO login script issues." }
            "Profile*"         { "FSLogix profile container failed to attach. Check storage account connectivity, VHD permissions, and available disk space on the file share." }
            "Auto-Reconnect"   { "—" }
            "Normal Completion" { "—" }
            "No Completion*"   { "—" }
            "User Initiated"   { "—" }
            default            { "Review error codes in Log Analytics for specific troubleshooting guidance." }
          }
          $pctVal = if ($dr.PSObject.Properties['Pct']) { $dr.Pct } else { 0 }
          $sampleErr = if ($dr.PSObject.Properties['SampleError'] -and $dr.SampleError) { $dr.SampleError } else { "—" }
          # Truncate long sample errors for display
          $sampleErrDisplay = if ($sampleErr.Length -gt 120) { $sampleErr.Substring(0, 117) + "..." } else { $sampleErr }
          $sampleErrEscaped = $sampleErr -replace '"','&quot;' -replace "'","&#39;"
          
          $htmlReport += @"
                    <tr$(if (-not $isNormal) { " style='background:#fff8f0'" })>
                        <td>$catBadge</td>
                        <td>$(if ($dr.PSObject.Properties['SessionCount']) { $dr.SessionCount } else { 0 })</td>
                        <td>$pctVal%</td>
                        <td>$(if ($dr.PSObject.Properties['DistinctUsers']) { $dr.DistinctUsers } else { 0 })</td>
                        <td>$(if ($dr.PSObject.Properties['DistinctHosts']) { $dr.DistinctHosts } else { 0 })</td>
                        <td style="font-size:11px;max-width:250px;color:#666" title="$sampleErrEscaped">$sampleErrDisplay</td>
                        <td style="font-size:11px;max-width:350px;color:#555">$(if (-not $isNormal) { $remediation } else { "—" })</td>
                    </tr>
"@
        }
        $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
      }

      # Disconnect Breakdown by Host
      if ($disconnectsByHostData.Count -gt 0) {
        $htmlReport += @"
        <div class="table-wrap" style="margin-top:16px">
            <div class="table-title">Disconnect Breakdown by Host</div>
            <div class="table-scroll">
                <table id="disc-host-table">
                    <thead><tr>
                        <th onclick="sortTable('disc-host-table',0)">Session Host</th>
                        <th onclick="sortTable('disc-host-table',1)">Sessions</th>
                        <th onclick="sortTable('disc-host-table',2)">Abnormal %</th>
                        <th onclick="sortTable('disc-host-table',3)"><span title="Sessions ending due to network connectivity loss">Network</span></th>
                        <th onclick="sortTable('disc-host-table',4)"><span title="Sessions ended by idle timeout policy">Timeout</span></th>
                        <th onclick="sortTable('disc-host-table',5)"><span title="Server-initiated disconnects (reboots, scaling, agent)">Server</span></th>
                        <th onclick="sortTable('disc-host-table',6)"><span title="Kerberos, CredSSP, or NLA authentication failures">Auth</span></th>
                        <th onclick="sortTable('disc-host-table',7)"><span title="Out of memory, disk full, or resource exhaustion">Resource</span></th>
                        <th onclick="sortTable('disc-host-table',8)"><span title="Other error codes not matching known categories">Other</span></th>
                    </tr></thead>
                    <tbody>
"@
        foreach ($dh in ($disconnectsByHostData | Sort-Object { if ($_.PSObject.Properties['AbnormalPct']) { [double]$_.AbnormalPct } else { 0 } } -Descending | Select-Object -First 20)) {
          $abPct = if ($dh.PSObject.Properties['AbnormalPct']) { [double]$dh.AbnormalPct } else { 0 }
          $abColor = if ($abPct -gt 15) { "color:#c62828;font-weight:600" } elseif ($abPct -gt 5) { "color:#e65100;font-weight:600" } else { "" }
          $netDrops = if ($dh.PSObject.Properties['NetworkDrops']) { $dh.NetworkDrops } else { 0 }
          $timeouts = if ($dh.PSObject.Properties['Timeouts']) { $dh.Timeouts } else { 0 }
          $serverSide = if ($dh.PSObject.Properties['ServerSide']) { $dh.ServerSide } else { 0 }
          $authFail = if ($dh.PSObject.Properties['AuthFailures']) { $dh.AuthFailures } else { 0 }
          $resSide = if ($dh.PSObject.Properties['ResourceIssues']) { $dh.ResourceIssues } else { 0 }
          $otherErrs = if ($dh.PSObject.Properties['OtherErrors']) { $dh.OtherErrors } else { 0 }

          $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-VMName $dh.SessionHostName)</strong></td>
                        <td>$(if ($dh.PSObject.Properties['TotalSessions']) { $dh.TotalSessions } else { 0 })</td>
                        <td style="$abColor">$abPct%</td>
                        <td>$(if ([int]$netDrops -gt 0) { "<span style='color:#c62828'>$netDrops</span>" } else { $netDrops })</td>
                        <td>$(if ([int]$timeouts -gt 0) { "<span style='color:#e65100'>$timeouts</span>" } else { $timeouts })</td>
                        <td>$(if ([int]$serverSide -gt 0) { "<span style='color:#c62828'>$serverSide</span>" } else { $serverSide })</td>
                        <td>$(if ([int]$authFail -gt 0) { "<span style='color:#e65100'>$authFail</span>" } else { $authFail })</td>
                        <td>$(if ([int]$resSide -gt 0) { "<span style='color:#c62828'>$resSide</span>" } else { $resSide })</td>
                        <td>$(if ([int]$otherErrs -gt 0) { "<span style='color:#e65100'>$otherErrs</span>" } else { $otherErrs })</td>
                    </tr>
"@
        }
        $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
      }

      # Legacy: Simple disconnect rate table (fallback if new queries returned no data)
      if ($disconnectReasonData.Count -eq 0 -and $disconnectsByHostData.Count -eq 0 -and $disconnectData.Count -gt 0) {
        $highDisc = @($disconnectData | Where-Object { $_.PSObject.Properties['DisconnectPct'] -and [double]$_.DisconnectPct -gt 10 })
        if ($highDisc.Count -gt 0) {
          $htmlReport += @"
        <div class="alert alert-danger" style="margin-top:24px">
            <strong>⚠️ $($highDisc.Count) host(s) have &gt;10% unexpected disconnect rate.</strong> Sessions ending in under 60 seconds often indicate crashes, timeouts, or resource exhaustion.
        </div>
"@
        }
        $htmlReport += @"
        <div class="table-wrap">
            <div class="table-title">Unexpected Disconnect Rates by Host</div>
            <div class="table-scroll">
                <table id="disc-table">
                    <thead><tr>
                        <th onclick="sortTable('disc-table',0)">Session Host</th>
                        <th onclick="sortTable('disc-table',1)">Total Sessions</th>
                        <th onclick="sortTable('disc-table',2)">Disconnects</th>
                        <th onclick="sortTable('disc-table',3)">Disconnect %</th>
                        <th onclick="sortTable('disc-table',4)">Avg Session (min)</th>
                    </tr></thead>
                    <tbody>
"@
        foreach ($dc in ($disconnectData | Sort-Object { if ($_.PSObject.Properties['DisconnectPct']) { [double]($_.DisconnectPct) } else { 0 } } -Descending | Select-Object -First 20)) {
          $dcPct = if ($dc.PSObject.Properties['DisconnectPct']) { [double]$dc.DisconnectPct } else { 0 }
          $dcColor = if ($dcPct -gt 10) { "color:#c62828;font-weight:600" } elseif ($dcPct -gt 5) { "color:#e65100;font-weight:600" } else { "" }
          
          $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-VMName $dc.SessionHostName)</strong></td>
                        <td>$($dc.TotalSessions)</td>
                        <td>$($dc.UnexpectedDisconnects)</td>
                        <td style="$dcColor">$($dc.DisconnectPct)%</td>
                        <td>$($dc.AvgSessionMinutes)</td>
                    </tr>
"@
        }
        $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
      }
    }

    # --- Cross-Region Connection Analysis ---
    if ($crossRegionAnalysis.Count -gt 0) {
      # Summary cards
      $totalCrossContinent = ($crossContinentPaths | ForEach-Object { [int]$_.Connections } | Measure-Object -Sum).Sum
      $totalCrossRegion = ($crossRegionPaths | ForEach-Object { [int]$_.Connections } | Measure-Object -Sum).Sum
      $totalSameRegion = ($sameRegionPaths | ForEach-Object { [int]$_.Connections } | Measure-Object -Sum).Sum
      $totalAllPaths = ($crossRegionAnalysis | ForEach-Object { [int]$_.Connections } | Measure-Object -Sum).Sum
      if (-not $totalAllPaths) { $totalAllPaths = 1 }
      $crossContinentPct = [math]::Round(($totalCrossContinent / $totalAllPaths) * 100, 0)
      $worstPath = $crossRegionAnalysis | Sort-Object { [double]$_.P95RTTms } -Descending | Select-Object -First 1
      
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">🌍 Cross-Region Connection Analysis</h3>
        <div class="alert alert-info" style="line-height:1.6">
            Users connect through their nearest RD Gateway, but session hosts may be in a different region or continent.
            This analysis correlates gateway regions with session host locations to identify latency caused by geographic distance versus network routing issues.
        </div>
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Cross-Continent</div>
                <div class="card-value $(if ($crossContinentPaths.Count -gt 0) { 'red' } else { 'green' })">$($crossContinentPaths.Count) paths</div>
                <div class="card-sub">$totalCrossContinent connections ($crossContinentPct%)</div>
            </div>
            <div class="card">
                <div class="card-label">Cross-Region</div>
                <div class="card-value $(if ($crossRegionPaths.Count -gt 0) { 'yellow' } else { 'green' })">$($crossRegionPaths.Count) paths</div>
                <div class="card-sub">$totalCrossRegion connections (same continent)</div>
            </div>
            <div class="card">
                <div class="card-label">Same Region</div>
                <div class="card-value green">$($sameRegionPaths.Count) paths</div>
                <div class="card-sub">$totalSameRegion connections (optimal)</div>
            </div>
            <div class="card">
                <div class="card-label">Worst P95 RTT</div>
                <div class="card-value $(if ([double]$worstPath.P95RTTms -gt 150) { 'red' } elseif ([double]$worstPath.P95RTTms -gt 100) { 'yellow' } else { 'green' })">$($worstPath.P95RTTms) ms</div>
                <div class="card-sub">$($worstPath.GatewayRegion) → $($worstPath.HostRegion)</div>
            </div>
        </div>
"@
      
      # Cross-region detail table
      $htmlReport += @"
        <div class="table-wrap">
            <div class="table-title">Gateway → Session Host Connection Paths</div>
            <div class="table-scroll">
                <table id="crossregion-table">
                    <thead><tr>
                        <th onclick="sortTable('crossregion-table',0)">Gateway Region</th>
                        <th>→</th>
                        <th onclick="sortTable('crossregion-table',2)">Host Region</th>
                        <th onclick="sortTable('crossregion-table',3)">Distance</th>
                        <th onclick="sortTable('crossregion-table',4)">Expected RTT</th>
                        <th onclick="sortTable('crossregion-table',5)">Avg RTT</th>
                        <th onclick="sortTable('crossregion-table',6)">P95 RTT</th>
                        <th onclick="sortTable('crossregion-table',7)">Excess RTT</th>
                        <th onclick="sortTable('crossregion-table',8)">Rating</th>
                        <th onclick="sortTable('crossregion-table',9)">Users</th>
                        <th onclick="sortTable('crossregion-table',10)">Sessions</th>
                    </tr></thead>
                    <tbody>
"@
      foreach ($path in ($crossRegionAnalysis | Sort-Object { [double]$_.P95RTTms } -Descending)) {
        $pathIcon = if ($path.IsCrossContinent) { "🌍" } elseif ($path.IsCrossRegion) { "🔀" } else { "✅" }
        $ratingColor = switch ($path.RTTRating) {
          "Critical"   { "color:#c62828;font-weight:700" }
          "Poor"       { "color:#e65100;font-weight:600" }
          "Acceptable" { "color:#f57f17;font-weight:500" }
          "Good"       { "color:#2e7d32" }
          "Excellent"  { "color:#1b5e20" }
          default      { "" }
        }
        $distStr = if ($path.DistanceKm -gt 0) { "$($path.DistanceKm.ToString('N0')) km / $($path.DistanceMi.ToString('N0')) mi" } else { "—" }
        $expectedStr = if ($path.ExpectedRTTms -gt 0) { "~$($path.ExpectedRTTms) ms" } else { "—" }
        $excessStr = if ($path.RTTExcessMs -gt 0) { "+$($path.RTTExcessMs) ms" } elseif ($path.ExpectedRTTms -gt 0) { "On target" } else { "—" }
        $excessColor = if ($path.RTTExcessMs -gt 100) { "color:#c62828;font-weight:600" } elseif ($path.RTTExcessMs -gt 50) { "color:#e65100" } else { "" }
        $p95Color = if ([double]$path.P95RTTms -gt 250) { "color:#c62828;font-weight:600" } elseif ([double]$path.P95RTTms -gt 150) { "color:#e65100;font-weight:600" } else { "" }
        
        $htmlReport += @"
                    <tr>
                        <td>$pathIcon <strong>$($path.GatewayRegion)</strong><br><span style="font-size:11px;color:#888">$($path.GatewayContinent)</span></td>
                        <td style="font-size:18px;color:#999">→</td>
                        <td><strong>$($path.HostRegion)</strong><br><span style="font-size:11px;color:#888">$($path.HostContinent)</span></td>
                        <td style="font-size:12px">$distStr</td>
                        <td style="font-size:12px">$expectedStr</td>
                        <td>$($path.AvgRTTms) ms</td>
                        <td style="$p95Color">$($path.P95RTTms) ms</td>
                        <td style="$excessColor">$excessStr</td>
                        <td style="$ratingColor">$($path.RTTRating)</td>
                        <td>$($path.DistinctUsers)</td>
                        <td>$($path.Connections)</td>
                    </tr>
"@
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
      
      # Actionable callouts for cross-continent connections
      if ($crossContinentPaths.Count -gt 0) {
        $htmlReport += @"
        <div class="alert alert-warning" style="margin-top:16px">
            <strong>⚠️ Cross-Continent Connections Detected</strong><br>
            <span style="font-size:13px">
            Users are connecting through gateways on a different continent than the session hosts. This adds unavoidable network latency due to geographic distance.<br><br>
            <strong>Recommendations:</strong><br>
"@
        # Build unique continent pairs
        $continentPairs = @{}
        foreach ($cp in $crossContinentPaths) {
          $pairKey = "$($cp.GatewayContinent) → $($cp.HostContinent)"
          if (-not $continentPairs[$pairKey]) {
            $continentPairs[$pairKey] = @{ Users = 0; Connections = 0; AvgP95 = [System.Collections.Generic.List[double]]::new() }
          }
          $continentPairs[$pairKey].Users += [int]$cp.DistinctUsers
          $continentPairs[$pairKey].Connections += [int]$cp.Connections
          $continentPairs[$pairKey].AvgP95.Add([double]$cp.P95RTTms)
        }
        
        foreach ($pair in $continentPairs.GetEnumerator()) {
          $avgP95 = [math]::Round(($pair.Value.AvgP95 | Measure-Object -Average).Average, 0)
          $htmlReport += "            • <strong>$($pair.Key)</strong>: $($pair.Value.Users) users, $($pair.Value.Connections) sessions, ~$($avgP95) ms P95 RTT<br>`n"
        }
        
        $htmlReport += @"
            <br>
            • <strong>Deploy session hosts closer to users:</strong> Consider deploying a host pool in the region where your users are located (e.g., Southeast Asia, Central India) to eliminate cross-continent latency<br>
            • <strong>Enable RDP Shortpath:</strong> Reduces latency by using a direct UDP connection between client and session host, bypassing the gateway relay<br>
            • <strong>Review user-to-host affinity:</strong> If only a subset of users are remote, consider a multi-region deployment with geo-based load balancing<br>
            • <strong>Expected baseline RTT</strong> is estimated from fiber-optic distance at ~100 km/ms + 10 ms overhead. Actual RTT significantly above baseline suggests network routing issues beyond simple distance
            </span>
        </div>
"@
      }
    }

    $htmlReport += @"
    </div>
"@
  }

  # ========== SCALING & AUTOSCALE ANALYSIS ==========
  if ($concurrencyData.Count -gt 0 -or $scalingPlanSchedules.Count -gt 0) {

    $htmlReport += @"

    <div class="section" id="sec-scaling">
"@

    # ── Scaling Plan Configuration Table ──
    if ($scalingPlanSchedules.Count -gt 0) {

      # Build assignment lookup: which pools have plans and are they enabled?
      $assignmentLookup = @{}
      foreach ($a in $scalingPlanAssignments) {
        $poolName = if ($a.HostPoolName) { $a.HostPoolName } elseif ($a.HostPoolArmId) { ($a.HostPoolArmId -split '/')[-1] } else { "" }
        if ($poolName) {
          if (-not $assignmentLookup.ContainsKey($a.ScalingPlanName)) { $assignmentLookup[$a.ScalingPlanName] = @() }
          $assignmentLookup[$a.ScalingPlanName] += [PSCustomObject]@{ PoolName = $poolName; Enabled = $a.IsEnabled }
        }
      }

      # Count VMs per host pool and power states
      $poolVmCounts = @{}
      $poolRunning = @{}
      foreach ($v in $vms) {
        $p = $v.HostPoolName
        if ($p) {
          if (-not $poolVmCounts.ContainsKey($p)) { $poolVmCounts[$p] = 0; $poolRunning[$p] = 0 }
          $poolVmCounts[$p]++
          if ($v.PowerState -eq "VM running") { $poolRunning[$p]++ }
        }
      }

      # Identify host pools WITHOUT any scaling plan
      $poolsWithPlans = @{}
      foreach ($a in $scalingPlanAssignments) {
        $poolName = if ($a.HostPoolName) { $a.HostPoolName } elseif ($a.HostPoolArmId) { ($a.HostPoolArmId -split '/')[-1] } else { "" }
        if ($poolName) { $poolsWithPlans[$poolName] = $true }
      }
      $allPoolNames = @($vms | Select-Object -ExpandProperty HostPoolName -Unique)
      $poolsWithoutPlans = @($allPoolNames | Where-Object { -not $poolsWithPlans.ContainsKey($_) })
      $disabledAssignments = @($scalingPlanAssignments | Where-Object { $_.IsEnabled -eq $false -or $_.IsEnabled -eq "False" })

      # Scaling Plan Findings
      $scalingFindings = [System.Collections.Generic.List[string]]::new()
      if ($poolsWithoutPlans.Count -gt 0) {
        foreach ($np in $poolsWithoutPlans) {
          $running = if ($poolRunning.ContainsKey($np)) { $poolRunning[$np] } else { 0 }
          $total = if ($poolVmCounts.ContainsKey($np)) { $poolVmCounts[$np] } else { 0 }
          $scalingFindings.Add("$np has no scaling plan ($running of $total VMs running)")
        }
      }
      if ($disabledAssignments.Count -gt 0) {
        foreach ($da in $disabledAssignments) {
          $dPool = if ($da.HostPoolName) { $da.HostPoolName } elseif ($da.HostPoolArmId) { ($da.HostPoolArmId -split '/')[-1] } else { $da.ScalingPlanName }
          $scalingFindings.Add("$($da.ScalingPlanName) → $dPool is DISABLED")
        }
      }

      # Check for identical schedules (cookie-cutter detection)
      $scheduleSignatures = @{}
      foreach ($sch in $scalingPlanSchedules) {
        $sig = "$($sch.ScheduleName)|$($sch.RampUpStartTime)|$($sch.PeakStartTime)|$($sch.RampDownStartTime)|$($sch.OffPeakStartTime)|$($sch.RampUpCapacity)|$($sch.RampDownCapacity)"
        if (-not $scheduleSignatures.ContainsKey($sig)) { $scheduleSignatures[$sig] = @() }
        $scheduleSignatures[$sig] += $sch.ScalingPlanName
      }
      $identicalGroups = @($scheduleSignatures.GetEnumerator() | Where-Object { @($_.Value | Select-Object -Unique).Count -gt 2 })
      if ($identicalGroups.Count -gt 0) {
        $identicalCount = ($identicalGroups | ForEach-Object { $_.Value } | Select-Object -Unique).Count
        $scalingFindings.Add("$identicalCount scaling plans share identical schedules — consider tuning per workload")
      }

      # High ramp-down capacity detection
      $highRampDown = @($scalingPlanSchedules | Where-Object {
        $_.ScheduleName -eq "Weekdays" -and [int]$_.RampDownCapacity -gt 70
      })
      if ($highRampDown.Count -gt 0) {
        $scalingFindings.Add("$($highRampDown.Count) plans have weekday ramp-down capacity >70% — VMs stay running well past business hours")
      }

      # Weekend capacity detection
      $highWeekend = @($scalingPlanSchedules | Where-Object {
        $_.ScheduleName -eq "Weekend" -and [int]$_.RampUpCapacity -gt 50
      })
      if ($highWeekend.Count -gt 0) {
        $scalingFindings.Add("$($highWeekend.Count) plans have weekend capacity >50% — high for typical business workloads")
      }

      # BreadthFirst on RemoteApp with low usage — may be keeping more hosts warm than needed
      foreach ($pn in $allPoolNames) {
        $pnAppGroup = $hpAppGroupLookup[$pn]
        $pnLB = $hpLoadBalancerLookup[$pn]
        $pnType = $hpTypeLookup[$pn]
        if ($pnAppGroup -eq "RailApplications" -and $pnLB -eq "BreadthFirst" -and $pnType -eq "Pooled") {
          $pnRunning = if ($poolRunning.ContainsKey($pn)) { $poolRunning[$pn] } else { 0 }
          $pnTotal = if ($poolVmCounts.ContainsKey($pn)) { $poolVmCounts[$pn] } else { 0 }
          if ($pnRunning -gt 2 -and $pnTotal -gt 3) {
            $shortPn = ($pn -split '-' | Select-Object -Last 3) -join '-'
            $scalingFindings.Add("$shortPn uses BreadthFirst on RemoteApp ($pnRunning/$pnTotal running) — DepthFirst would consolidate sessions and allow more hosts to deallocate")
          }
        }
      }

      # Personal pools running 100% with no apparent scaling
      foreach ($pn in $allPoolNames) {
        $pnType = $hpTypeLookup[$pn]
        if ($pnType -eq "Personal") {
          $pnRunning = if ($poolRunning.ContainsKey($pn)) { $poolRunning[$pn] } else { 0 }
          $pnTotal = if ($poolVmCounts.ContainsKey($pn)) { $poolVmCounts[$pn] } else { 0 }
          $pnHasPlan = $poolsWithPlans.ContainsKey($pn)
          if (-not $pnHasPlan -and $pnRunning -gt 5 -and $pnTotal -gt 0) {
            $shortPn = ($pn -split '-' | Select-Object -Last 3) -join '-'
            $scalingFindings.Add("Personal pool $shortPn has no scaling plan ($pnRunning/$pnTotal running) — Start VM on Connect + deallocate-on-disconnect saves compute for idle desktops")
          }
        }
      }

      # StartVMOnConnect disabled
      $noStartVmPools = @($hostPools | Where-Object { $_.HostPoolName -and (-not $_.StartVMOnConnect -or $_.StartVMOnConnect -eq "False") })
      if ($noStartVmPools.Count -gt 0) {
        $noStartVmNames = ($noStartVmPools | ForEach-Object { ($_.HostPoolName -split '-' | Select-Object -Last 3) -join '-' }) -join ", "
        $scalingFindings.Add("$($noStartVmPools.Count) pool(s) have Start VM on Connect disabled ($noStartVmNames) — users must wait for pre-started hosts or manual boot")
      }

      # Summary cards
      $totalPlans = $scalingPlans.Count
      $enabledAssignments = @($scalingPlanAssignments | Where-Object { $_.IsEnabled -eq $true -or $_.IsEnabled -eq "True" }).Count
      $findingsColor = if ($scalingFindings.Count -eq 0) { "#4caf50" } elseif ($scalingFindings.Count -le 2) { "#ff9800" } else { "#f44336" }

      $htmlReport += @"
        <div class="card-grid" style="margin-bottom:20px">
            <div class="card"><div class="card-value">$totalPlans</div><div class="card-sub">Scaling Plans</div></div>
            <div class="card"><div class="card-value">$enabledAssignments</div><div class="card-sub">Active Assignments</div></div>
            <div class="card"><div class="card-value">$($poolsWithoutPlans.Count)</div><div class="card-sub" style="color:$(if ($poolsWithoutPlans.Count -gt 0) { '#f44336' } else { '#4caf50' })">Pools Without Plans</div></div>
            <div class="card"><div class="card-value" style="color:$findingsColor">$($scalingFindings.Count)</div><div class="card-sub">Findings</div></div>
        </div>
"@

      # Findings alerts
      if ($scalingFindings.Count -gt 0) {
        $htmlReport += "        <div class=`"alert alert-warning`" style=`"margin-bottom:16px`"><strong>Scaling Findings:</strong><br>"
        foreach ($f in $scalingFindings) {
          $htmlReport += "• $f<br>"
        }
        $htmlReport += "        </div>`n"
      }

      # Scaling Plan Schedule Table
      $htmlReport += @"
        <div class="table-wrap" style="margin-bottom:24px">
            <div class="table-title">Scaling Plan Schedules</div>
            <div class="table-scroll">
                <table id="scaling-plan-table">
                    <thead><tr>
                        <th onclick="sortTable('scaling-plan-table',0)">Scaling Plan</th>
                        <th onclick="sortTable('scaling-plan-table',1)">Schedule</th>
                        <th onclick="sortTable('scaling-plan-table',2)"><span title="Time ramp-up phase begins">Ramp Up</span></th>
                        <th onclick="sortTable('scaling-plan-table',3)"><span title="Time peak phase begins">Peak</span></th>
                        <th onclick="sortTable('scaling-plan-table',4)"><span title="Time ramp-down phase begins">Ramp Down</span></th>
                        <th onclick="sortTable('scaling-plan-table',5)"><span title="Time off-peak phase begins">Off-Peak</span></th>
                        <th onclick="sortTable('scaling-plan-table',6)"><span title="% of hosts powered on during ramp-up">RampUp %</span></th>
                        <th onclick="sortTable('scaling-plan-table',7)"><span title="% capacity threshold during ramp-down">RampDn %</span></th>
                        <th onclick="sortTable('scaling-plan-table',8)"><span title="Force logoff during ramp-down?">Force Logoff</span></th>
                        <th onclick="sortTable('scaling-plan-table',9)">Assigned Pools</th>
                    </tr></thead>
                    <tbody>
"@

      # De-duplicate — group by plan+schedule
      $uniqueSchedules = @{}
      foreach ($sch in $scalingPlanSchedules) {
        $key = "$($sch.ScalingPlanName)|$($sch.ScheduleName)"
        $uniqueSchedules[$key] = $sch
      }

      foreach ($sch in ($uniqueSchedules.Values | Sort-Object ScalingPlanName, ScheduleName)) {
        $planName = $sch.ScalingPlanName
        $shortPlan = if ($planName.Length -gt 35) { $planName.Substring(0,32) + "..." } else { $planName }

        # Parse times — handle both "@{hour=X; minute=Y}" format and direct strings
        $rampUpTime = "$($sch.RampUpStartTime)" -replace '@\{hour=(\d+);\s*minute=(\d+)\}', '$1:$2'
        $peakTime = "$($sch.PeakStartTime)" -replace '@\{hour=(\d+);\s*minute=(\d+)\}', '$1:$2'
        $rampDnTime = "$($sch.RampDownStartTime)" -replace '@\{hour=(\d+);\s*minute=(\d+)\}', '$1:$2'
        $offPeakTime = "$($sch.OffPeakStartTime)" -replace '@\{hour=(\d+);\s*minute=(\d+)\}', '$1:$2'

        $rampUpCap = $sch.RampUpCapacity
        $rampDnCap = $sch.RampDownCapacity
        $rampDnBadge = if ([int]$rampDnCap -gt 70) { "b-yellow" } elseif ([int]$rampDnCap -gt 90) { "b-red" } else { "b-green" }
        $forceLogoff = if ($sch.RampDownForceLogoff -eq $true -or $sch.RampDownForceLogoff -eq "True") { "Yes ($($sch.RampDownLogoffTimeoutMinutes)min)" } else { "No" }

        # Get assigned pools
        $pools = if ($assignmentLookup.ContainsKey($planName)) {
          ($assignmentLookup[$planName] | ForEach-Object {
            $status = if ($_.Enabled -eq $true -or $_.Enabled -eq "True") { "" } else { " <span class=`"badge b-red`">OFF</span>" }
            "$($_.PoolName)$status"
          }) -join "<br>"
        } else { "<span style=`"color:#999`">Unassigned</span>" }

        $htmlReport += @"
                    <tr>
                        <td><span title="$planName">$shortPlan</span></td>
                        <td>$($sch.ScheduleName)</td>
                        <td style="font-family:monospace">$rampUpTime</td>
                        <td style="font-family:monospace">$peakTime</td>
                        <td style="font-family:monospace">$rampDnTime</td>
                        <td style="font-family:monospace">$offPeakTime</td>
                        <td>$rampUpCap%</td>
                        <td><span class="badge $rampDnBadge">$rampDnCap%</span></td>
                        <td>$forceLogoff</td>
                        <td style="font-size:11px">$pools</td>
                    </tr>
"@
      }

      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@

      # Host Pool Scaling Coverage Table
      $htmlReport += @"
        <div class="table-wrap" style="margin-bottom:24px">
            <div class="table-title">Host Pool Scaling Coverage</div>
            <div class="table-scroll">
                <table id="scaling-coverage-table">
                    <thead><tr>
                        <th onclick="sortTable('scaling-coverage-table',0)">Host Pool</th>
                        <th onclick="sortTable('scaling-coverage-table',1)"><span title="Personal / Pooled">Type</span></th>
                        <th onclick="sortTable('scaling-coverage-table',2)"><span title="Desktop or RemoteApp">Workload</span></th>
                        <th onclick="sortTable('scaling-coverage-table',3)"><span title="BreadthFirst / DepthFirst / Persistent">Load Balancer</span></th>
                        <th onclick="sortTable('scaling-coverage-table',4)">Total VMs</th>
                        <th onclick="sortTable('scaling-coverage-table',5)">Running</th>
                        <th onclick="sortTable('scaling-coverage-table',6)">Deallocated</th>
                        <th onclick="sortTable('scaling-coverage-table',7)"><span title="% of VMs currently running">Running %</span></th>
                        <th onclick="sortTable('scaling-coverage-table',8)">Scaling Plan</th>
                        <th onclick="sortTable('scaling-coverage-table',9)"><span title="Start VM on Connect enabled?">StartVM</span></th>
                        <th onclick="sortTable('scaling-coverage-table',10)">Status</th>
                    </tr></thead>
                    <tbody>
"@

      foreach ($poolName in ($allPoolNames | Sort-Object)) {
        $total = if ($poolVmCounts.ContainsKey($poolName)) { $poolVmCounts[$poolName] } else { 0 }
        $running = if ($poolRunning.ContainsKey($poolName)) { $poolRunning[$poolName] } else { 0 }
        $dealloc = $total - $running
        $runPct = if ($total -gt 0) { [math]::Round($running / $total * 100, 0) } else { 0 }

        # Pool metadata
        $covPoolType = $hpTypeLookup[$poolName]
        $covAppGroup = $hpAppGroupLookup[$poolName]
        $covLB = $hpLoadBalancerLookup[$poolName]
        $workloadDisplay = if ($covAppGroup -eq "RailApplications") { "RemoteApp" } elseif ($covAppGroup -eq "Desktop") { "Desktop" } else { $covAppGroup }
        $lbDisplay = if ($covLB -eq "BreadthFirst") { "Breadth" } elseif ($covLB -eq "DepthFirst") { "Depth" } elseif ($covLB -eq "Persistent") { "Persistent" } else { $covLB }

        # Find assigned plan
        $assignedPlan = $scalingPlanAssignments | Where-Object {
          ($_.HostPoolName -eq $poolName) -or (($_.HostPoolArmId -split '/')[-1] -eq $poolName)
        } | Select-Object -First 1
        $planDisplay = if ($assignedPlan) { $assignedPlan.ScalingPlanName } else { "—" }
        $shortPlanDisplay = if ($planDisplay.Length -gt 30) { $planDisplay.Substring(0,27) + "..." } else { $planDisplay }

        $statusBadge = if (-not $assignedPlan) {
          "<span class=`"badge b-red`">No Plan</span>"
        } elseif ($assignedPlan.IsEnabled -eq $false -or $assignedPlan.IsEnabled -eq "False") {
          "<span class=`"badge b-red`">Disabled</span>"
        } else {
          "<span class=`"badge b-green`">Active</span>"
        }

        $runPctBadge = if ($runPct -eq 100 -and $total -gt 1) { "b-yellow" } elseif ($runPct -eq 0 -and $total -gt 0) { "b-gray" } else { "" }

        # StartVMOnConnect lookup
        $poolHpObj = $hostPools | Where-Object { $_.HostPoolName -eq $poolName } | Select-Object -First 1
        $startVmEnabled = if ($poolHpObj -and ($poolHpObj.StartVMOnConnect -eq $true -or $poolHpObj.StartVMOnConnect -eq "True")) { $true } else { $false }
        $startVmBadge = if ($startVmEnabled) { "<span class=`"badge b-green`">On</span>" } else { "<span class=`"badge b-red`">Off</span>" }

        $htmlReport += @"
                    <tr>
                        <td>$(Scrub-HostPoolName $poolName)</td>
                        <td>$covPoolType</td>
                        <td>$workloadDisplay</td>
                        <td>$lbDisplay</td>
                        <td>$total</td>
                        <td>$running</td>
                        <td>$dealloc</td>
                        <td><span class="badge $(if ($runPctBadge) { $runPctBadge })">$runPct%</span></td>
                        <td style="font-size:11px" title="$planDisplay">$shortPlanDisplay</td>
                        <td>$startVmBadge</td>
                        <td>$statusBadge</td>
                    </tr>
"@
      }

      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@

      # Guidance note
      $htmlReport += @"
        <div class="alert alert-info" style="margin-bottom:24px">
            <strong>Tuning Tips:</strong><br>
            • <strong>Ramp-down capacity >70%</strong> means most hosts stay running well past business hours. For low-usage pools, try 20-40%.<br>
            • <strong>Weekend capacity should match weekend usage.</strong> If your pools are business-only, weekend ramp-up/down capacity of 10-20% avoids paying for idle VMs all weekend.<br>
            • <strong>Identical schedules across all pools</strong> is a sign the plans were copy-pasted. Pools with different workload profiles should have different scaling thresholds.<br>
            • <strong>RemoteApp pools can scale more aggressively</strong> than Desktop pools — users reconnect seamlessly to any host, and sessions are lighter weight. Consider lower ramp-up capacity and faster ramp-down.<br>
            • <strong>BreadthFirst vs DepthFirst:</strong> BreadthFirst spreads users evenly (better UX). DepthFirst fills hosts before starting new ones (better cost). Match your LB strategy to your scaling thresholds.<br>
            • <strong>100% Running</strong> on a pool with a scaling plan may indicate the plan is disabled, the capacity thresholds are too high, or the pool has persistent sessions that prevent deallocation.
        </div>
"@

    } # end scaling plan config

    # ── Autoscale Effectiveness (from Log Analytics) ──
    $autoscaleDetailedData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_WVDAutoscaleDetailed" -and $_.QueryName -eq "AVD" -and $_.PoolName -and $_.PoolName -ne "" -and $_.PoolName -ne "NoTable" })
    $autoscaleSummaryData = @($laResults | Where-Object { $_.Label -eq "CurrentWindow_WVDAutoscaleActivity" -and $_.QueryName -eq "AVD" -and $_.Result -and $_.Result -ne "NoTable" })

    if ($autoscaleDetailedData.Count -gt 0 -or $autoscaleSummaryData.Count -gt 0) {

      # Compute totals
      $totalEvals = 0; $totalSucceeded = 0; $totalFailed = 0
      if ($autoscaleDetailedData.Count -gt 0) {
        foreach ($ad in $autoscaleDetailedData) {
          $totalEvals += [int]$ad.Evaluations
          $totalSucceeded += [int]$ad.Succeeded
          $totalFailed += [int]$ad.Failed
        }
      } elseif ($autoscaleSummaryData.Count -gt 0) {
        foreach ($as in $autoscaleSummaryData) {
          $eCount = [int]$as.EvaluationCount
          if ($as.Result -match "Succeeded") { $totalSucceeded += $eCount }
          elseif ($as.Result -match "Failed") { $totalFailed += $eCount }
          $totalEvals += $eCount
        }
      }
      $successRate = if ($totalEvals -gt 0) { [math]::Round(($totalSucceeded / $totalEvals) * 100, 0) } else { 0 }
      $successColor = if ($successRate -ge 95) { "#4caf50" } elseif ($successRate -ge 80) { "#ff9800" } else { "#f44336" }

      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">⚡ Autoscale Effectiveness</h3>
        <div class="card-grid" style="margin-bottom:16px">
            <div class="card"><div class="card-value">$totalEvals</div><div class="card-sub">Total Evaluations</div></div>
            <div class="card"><div class="card-value" style="color:#4caf50">$totalSucceeded</div><div class="card-sub">Succeeded</div></div>
            <div class="card"><div class="card-value" style="color:$(if ($totalFailed -gt 0) { '#f44336' } else { '#4caf50' })">$totalFailed</div><div class="card-sub">Failed</div></div>
            <div class="card"><div class="card-value" style="color:$successColor">$successRate%</div><div class="card-sub">Success Rate</div></div>
        </div>
"@

      if ($totalFailed -gt 0) {
        $htmlReport += @"
        <div class="alert alert-warning" style="margin-bottom:16px">
            <strong>$totalFailed autoscale evaluation(s) failed.</strong> Failed evaluations mean the autoscaler attempted to start or deallocate VMs but was unable to — this can be caused by
            allocation failures (no capacity for the requested SKU), insufficient permissions, or Azure platform throttling. Check the <code>WVDAutoscaleEvaluationPooled</code> table in Log Analytics for error details.
        </div>
"@
      }

      # Per-pool breakdown (if detailed data available)
      if ($autoscaleDetailedData.Count -gt 0) {
        $htmlReport += @"
        <div class="table-wrap" style="margin-bottom:24px">
            <div class="table-title">Autoscale Evaluations by Host Pool</div>
            <div class="table-scroll">
                <table id="autoscale-table">
                    <thead><tr>
                        <th onclick="sortTable('autoscale-table',0)">Host Pool</th>
                        <th onclick="sortTable('autoscale-table',1)">Evaluations</th>
                        <th onclick="sortTable('autoscale-table',2)">Succeeded</th>
                        <th onclick="sortTable('autoscale-table',3)">Failed</th>
                        <th onclick="sortTable('autoscale-table',4)">Success Rate</th>
                        <th onclick="sortTable('autoscale-table',5)"><span title="Average active session hosts during evaluations">Avg Active Hosts</span></th>
                        <th onclick="sortTable('autoscale-table',6)"><span title="Maximum active session hosts observed">Max Active Hosts</span></th>
                        <th onclick="sortTable('autoscale-table',7)"><span title="Average sessions during evaluations">Avg Sessions</span></th>
                    </tr></thead>
                    <tbody>
"@

        foreach ($ad in ($autoscaleDetailedData | Sort-Object { [int]$_.Failed } -Descending)) {
          $poolEvals = [int]$ad.Evaluations
          $poolSucceeded = [int]$ad.Succeeded
          $poolFailed = [int]$ad.Failed
          $poolRate = if ($poolEvals -gt 0) { [math]::Round(($poolSucceeded / $poolEvals) * 100, 0) } else { 0 }
          $rateBadge = if ($poolRate -ge 95) { "b-green" } elseif ($poolRate -ge 80) { "b-yellow" } else { "b-red" }
          $failBadge = if ($poolFailed -gt 0) { "<span class=`"badge b-red`">$poolFailed</span>" } else { "<span style=`"color:#999`">0</span>" }

          $htmlReport += @"
                    <tr>
                        <td><strong>$($ad.PoolName)</strong></td>
                        <td>$poolEvals</td>
                        <td>$poolSucceeded</td>
                        <td>$failBadge</td>
                        <td><span class="badge $rateBadge">$poolRate%</span></td>
                        <td>$($ad.AvgActiveHosts)</td>
                        <td>$($ad.MaxActiveHosts)</td>
                        <td>$($ad.AvgSessions)</td>
                    </tr>
"@
        }

        $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
      }
    } # end autoscale effectiveness

    # ── Concurrency Bar Chart (from Log Analytics) ──
    if ($concurrencyData.Count -gt 0) {
      $maxConcurrency = ($concurrencyData | ForEach-Object { [int]$_.PeakConcurrency } | Measure-Object -Maximum).Maximum
      if ($maxConcurrency -eq 0) { $maxConcurrency = 1 }

      $htmlReport += @"
        <h3 style="margin:0 0 16px 0;font-size:16px;color:#333">Weekday Session Concurrency by Hour</h3>
        <div class="alert alert-info">
            Compare this pattern with your scaling plan ramp-up/ramp-down times. If users arrive before ramp-up starts, sessions hit cold hosts.
        </div>
        <div class="table-wrap">
            <div style="padding:24px;">
"@

      foreach ($hr in ($concurrencyData | Sort-Object { [int]$_.HourOfDay })) {
        $hour = [int]$hr.HourOfDay
        $hourLabel = "{0:00}:00" -f $hour
        $avgW = [math]::Round([int]$hr.AvgConcurrency / $maxConcurrency * 100, 0)
        $peakW = [math]::Round([int]$hr.PeakConcurrency / $maxConcurrency * 100, 0)
        $p95W = [math]::Round([int]$hr.P95Concurrency / $maxConcurrency * 100, 0)
        $barColor = if ($hour -ge 7 -and $hour -le 18) { "#0078d4" } else { "#90caf9" }

        $htmlReport += @"
                <div style="display:flex;align-items:center;gap:12px;margin-bottom:6px;">
                    <span style="width:45px;font-size:12px;font-weight:500;text-align:right;color:#555;font-family:monospace">$hourLabel</span>
                    <div style="flex:1;position:relative;height:24px;">
                        <div style="position:absolute;width:${peakW}%;height:100%;background:#e3f2fd;border-radius:4px" title="Peak: $($hr.PeakConcurrency)"></div>
                        <div style="position:absolute;width:${p95W}%;height:100%;background:#90caf9;border-radius:4px" title="P95: $($hr.P95Concurrency)"></div>
                        <div style="position:absolute;width:${avgW}%;height:100%;background:$barColor;border-radius:4px" title="Avg: $($hr.AvgConcurrency)"></div>
                    </div>
                    <span style="width:120px;font-size:11px;color:#555">Avg:$($hr.AvgConcurrency) P95:$($hr.P95Concurrency) Peak:$($hr.PeakConcurrency)</span>
                </div>
"@
      }

      $htmlReport += @"
            </div>
            <div style="padding:0 24px 16px;font-size:12px;color:#888;display:flex;gap:16px">
                <span>■ <span style="color:#0078d4">Avg</span></span>
                <span>■ <span style="color:#90caf9">P95</span></span>
                <span>■ <span style="color:#e3f2fd">Peak</span></span>
                <span style="margin-left:auto">Business hours (07-18) highlighted</span>
            </div>
        </div>
"@
    } # end concurrency chart

    $htmlReport += @"
    </div>
"@
  }

  # Advisor section
  if ($IncludeAzureAdvisor -and (SafeCount $advisorRecommendations) -gt 0) {
    $htmlReport += @"
    
    <!-- ========== ADVISOR ========== -->
    <div class="section" id="sec-advisor">
        <div class="table-wrap">
            <div class="table-title">Azure Advisor Recommendations (Top 30)</div>
            <div class="table-scroll">
                <table id="advisor-table">
                    <thead><tr>
                        <th onclick="sortTable('advisor-table',0)">Category</th>
                        <th onclick="sortTable('advisor-table',1)">Impact</th>
                        <th onclick="sortTable('advisor-table',2)">VM</th>
                        <th onclick="sortTable('advisor-table',3)">Description</th>
                        <th onclick="sortTable('advisor-table',4)">Solution</th>
                    </tr></thead>
                    <tbody>
"@
    $topAdvisor = $advisorRecommendations | Select-Object -First 30
    foreach ($adv in $topAdvisor) {
      $impBadge = switch ($adv.Impact) { "High" { "b-red" }; "Medium" { "b-yellow" }; default { "b-green" } }
      $vmDisplay = if ($adv.VMName) { $adv.VMName } else { "Subscription" }
      
      $htmlReport += @"
                    <tr>
                        <td><strong>$($adv.Category)</strong></td>
                        <td><span class="badge $impBadge">$($adv.Impact)</span></td>
                        <td>$vmDisplay</td>
                        <td style="font-size:12px">$($adv.ShortDescription)</td>
                        <td style="font-size:12px">$($adv.Solution)</td>
                    </tr>
"@
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
    </div>
"@
  }

  # W365 Cloud PC Readiness section
  if ($w365Analysis.Count -gt 0) {
    $w365Strong = @($w365Analysis | Where-Object { $_.Recommendation -eq "Strong W365 Candidate" })
    $w365Consider = @($w365Analysis | Where-Object { $_.Recommendation -match "Consider" })
    $w365Keep = @($w365Analysis | Where-Object { $_.Recommendation -eq "Keep AVD" })
    $w365TotalSavings = ($w365Analysis | Where-Object { $_.CostDelta -and $_.CostDelta -lt 0 -and $_.Recommendation -match "W365|Consider" } | ForEach-Object { [math]::Abs($_.CostDelta) } | Measure-Object -Sum).Sum
    $w365TcoSavings = ($w365Analysis | Where-Object { $_.TCOCostDelta -and $_.TCOCostDelta -lt 0 -and $_.Recommendation -match "W365|Consider" } | ForEach-Object { [math]::Abs($_.TCOCostDelta) } | Measure-Object -Sum).Sum
    $w365UsageSavings = ($w365Analysis | Where-Object { $_.UsageSavings -and $_.UsageSavings -gt 0 } | ForEach-Object { $_.UsageSavings } | Measure-Object -Sum).Sum
    
    # Find best pilot pool
    $bestPilot = $w365Analysis | Where-Object { $_.PilotScore -gt 0 -and $_.Recommendation -match "W365|Consider" } | Sort-Object PilotScore -Descending | Select-Object -First 1
    
    $htmlReport += @"
    
    <!-- ========== W365 READINESS ========== -->
    <div class="section" id="sec-w365">
        <div class="alert alert-info" style="line-height:1.6">
            <strong>💻 W365 Cloud PC Readiness Assessment</strong><br>
            This analysis evaluates each host pool's fit for Windows 365 Cloud PCs based on pool type, session density,
            VM sizing, actual usage metrics, and full TCO comparison including storage and profile costs.
            W365 pricing shown is list pricing (East US). <strong>This is directional guidance — validate with Microsoft licensing.</strong>
        </div>
        
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Strong W365 Candidates</div>
                <div class="card-value $(if ($w365Strong.Count -gt 0) { 'green' } else { 'blue' })">$($w365Strong.Count)</div>
            </div>
            <div class="card">
                <div class="card-label">Consider W365 / Hybrid</div>
                <div class="card-value $(if ($w365Consider.Count -gt 0) { 'orange' } else { 'blue' })">$($w365Consider.Count)</div>
            </div>
            <div class="card">
                <div class="card-label">Keep AVD</div>
                <div class="card-value blue">$($w365Keep.Count)</div>
            </div>
            <div class="card">
                <div class="card-label">Compute Savings (W365)</div>
                <div class="card-value $(if ($w365TotalSavings -gt 0) { 'green' } else { 'blue' })">`$$w365TotalSavings/mo</div>
            </div>
            <div class="card">
                <div class="card-label">Full TCO Savings</div>
                <div class="card-value $(if ($w365TcoSavings -gt 0) { 'green' } else { 'blue' })">`$$w365TcoSavings/mo</div>
                <div class="card-sub">incl. storage + profiles</div>
            </div>
            <div class="card">
                <div class="card-label">Annual Projection</div>
                <div class="card-value $(if ($w365TcoSavings -gt 0) { 'green' } else { 'blue' })">`$$([math]::Round($w365TcoSavings * 12, 0))/yr</div>
                <div class="card-sub">candidates only</div>
            </div>
        </div>
$(if ($w365UsageSavings -gt 0) {
@"
        <div class="alert" style="background:#e8f5e9;border-left:4px solid #4caf50;padding:12px 16px;margin-bottom:16px">
            <strong>💡 Usage-Based SKU Optimization:</strong> Based on actual CPU and memory usage, right-sizing W365 plans to match workload (not VM spec) would save an additional <strong>`$$w365UsageSavings/mo</strong> across candidate pools. See detail cards below.
        </div>
"@
})
$(if ($bestPilot) {
  $pilotShort = Scrub-HostPoolName $bestPilot.HostPoolName
@"
        <div style="background:linear-gradient(135deg,#e3f2fd,#f3e5f5);border-left:4px solid #7c4dff;border-radius:4px;padding:16px 20px;margin-bottom:16px">
            <h4 style="margin:0 0 8px 0;font-size:14px;color:#4a148c">🚀 Recommended Pilot Pool: $(Scrub-HostPoolName $bestPilot.HostPoolName)</h4>
            <div style="font-size:13px;line-height:1.7;color:#333">
                <div style="margin-bottom:8px">Pilot suitability: <strong>$($bestPilot.PilotScore)/100</strong> — $($bestPilot.PilotReasons -replace '; ', ' · ')</div>
                <div style="display:flex;gap:24px;flex-wrap:wrap;margin-bottom:8px">
                    <span>📊 Fit Score: <strong>$($bestPilot.FitScore)/100</strong></span>
                    <span>🖥️ VMs: <strong>$($bestPilot.VMCount)</strong></span>
                    <span>🔄 Migration: <strong>$($bestPilot.MigrationComplexity)</strong></span>
                    <span>$(if ($bestPilot.IntuneReady) { '✅' } else { '⚠️' }) Intune: <strong>$(if ($bestPilot.IntuneReady) { 'Ready' } else { 'Needs setup' })</strong></span>
                </div>
                <div style="background:white;border-radius:4px;padding:12px;font-size:12px">
                    <strong>3-Step Pilot Plan:</strong><br>
                    1️⃣ Provision $([math]::Min(5, $bestPilot.VMCount)) W365 Cloud PCs ($($bestPilot.BestW365Plan)) via Intune for volunteer users<br>
                    2️⃣ Run parallel for 2 weeks — compare login times, app performance, and user satisfaction<br>
                    3️⃣ If successful, migrate remaining $($bestPilot.VMCount) users in $([math]::Ceiling($bestPilot.VMCount / 10)) wave(s) of 10, decommissioning AVD hosts as users move
                </div>
            </div>
        </div>
"@
})
        
        <div class="table-wrap">
            <div class="table-title">Host Pool W365 Readiness</div>
            <div class="table-scroll">
                <table id="w365-table">
                    <thead><tr>
                        <th onclick="sortTable('w365-table',0)">Host Pool</th>
                        <th onclick="sortTable('w365-table',1)">Type</th>
                        <th onclick="sortTable('w365-table',2)"><span title="Desktop or RemoteApp">Workload</span></th>
                        <th onclick="sortTable('w365-table',3)">VMs</th>
                        <th onclick="sortTable('w365-table',4)"><span title="Unique users (from Log Analytics or assigned users)">Users</span></th>
                        <th onclick="sortTable('w365-table',5)"><span title="W365 licenses needed (= users for Enterprise, concurrent for Frontline)">W365 Licenses</span></th>
                        <th onclick="sortTable('w365-table',6)">Current SKU</th>
                        <th onclick="sortTable('w365-table',7)">Fit Score</th>
                        <th onclick="sortTable('w365-table',8)">Recommendation</th>
                        <th onclick="sortTable('w365-table',9)">Best W365 Plan</th>
                        <th onclick="sortTable('w365-table',10)">W365/mo</th>
                        <th onclick="sortTable('w365-table',11)">AVD/mo</th>
                        <th onclick="sortTable('w365-table',12)">Delta</th>
                        <th onclick="sortTable('w365-table',13)"><span title="Estimated migration complexity">Migration</span></th>
                    </tr></thead>
                    <tbody>
"@
    foreach ($wa in ($w365Analysis | Sort-Object FitScore -Descending)) {
      $recBadge = switch ($wa.Recommendation) {
        "Strong W365 Candidate" { "<span class='badge b-green'>$($wa.Recommendation)</span>" }
        "Consider W365 / Hybrid" { "<span class='badge b-orange'>$($wa.Recommendation)</span>" }
        default { "<span class='badge b-gray'>$($wa.Recommendation)</span>" }
      }
      $scoreColor = if ($wa.FitScore -ge 70) { "color:#2e7d32;font-weight:600" } elseif ($wa.FitScore -ge 45) { "color:#e65100;font-weight:600" } else { "" }
      $deltaStr = if ($null -ne $wa.CostDelta) {
        if ($wa.CostDelta -lt 0) { "<span style='color:#2e7d32;font-weight:600'>-`$$([math]::Abs($wa.CostDelta))</span>" }
        elseif ($wa.CostDelta -gt 0) { "<span style='color:#c62828'>+`$$($wa.CostDelta)</span>" }
        else { "`$0" }
      } else { "—" }
      $userVmRatio = if ($wa.UniqueUsers -gt $wa.VMCount) { "<br><span style='font-size:10px;color:#e65100'>$($wa.UniqueUsers):$($wa.VMCount) user:VM</span>" } else { "" }
      $w365CostBreakdown = if ($wa.W365MonthlyTotal -and $wa.W365LicensesNeeded -gt 0) { "`$$($wa.W365MonthlyTotal)<br><span style='font-size:10px;color:#888'>$($wa.W365LicensesNeeded) × `$$($wa.W365MonthlyPerUser)/user</span>" } elseif ($wa.W365MonthlyTotal) { "`$$($wa.W365MonthlyTotal)" } else { "—" }
      
      # Show usage-based plan if it differs from spec-matched
      $planDisplay = "<span style='font-size:12px'>$($wa.BestW365Plan)</span>"
      if ($wa.UsageBasedPlan -and $wa.UsageBasedMonthly -and $wa.UsageBasedMonthly -lt $wa.W365MonthlyPerUser) {
        $planDisplay += "<br><span style='font-size:10px;color:#2e7d32'>💡 Usage-based: $($wa.UsageBasedPlan) @ `$$($wa.UsageBasedMonthly)/user</span>"
      }
      
      $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $wa.HostPoolName)</strong></td>
                        <td>$($wa.HostPoolType)</td>
                        <td>$(if ($wa.AppGroupType -eq 'RailApplications') { "<span class='badge' style='background:#e3f2fd;color:#1565c0;font-size:10px'>RemoteApp</span>" } elseif ($wa.AppGroupType -eq 'Desktop') { 'Desktop' } else { $wa.AppGroupType })</td>
                        <td>$($wa.VMCount)</td>
                        <td>$($wa.UniqueUsers)$userVmRatio</td>
                        <td style="$(if ($wa.W365LicensesNeeded -gt $wa.VMCount * 2) { 'color:#c62828;font-weight:600' } else { '' })">$($wa.W365LicensesNeeded)</td>
                        <td>$($wa.DominantSku)<br><span style='font-size:11px;color:#888'>$($wa.vCPU) vCPU / $($wa.RamGB) GB</span></td>
                        <td style="$scoreColor">$($wa.FitScore)/100</td>
                        <td>$recBadge</td>
                        <td>$planDisplay</td>
                        <td>$w365CostBreakdown</td>
                        <td>$(if ($wa.AVDEffectiveMonthly) { "`$$($wa.AVDEffectiveMonthly)" } else { "—" })<br><span style='font-size:10px;color:#888'>$(if ($wa.CostSource -eq 'Actual') { '📊 Actual billing' } elseif ($wa.HasScalingPlan) { '~60% utilization est.' } else { 'PAYG estimate' })</span></td>
                        <td>$deltaStr</td>
                        <td>$(
                          $mcBadge = switch ($wa.MigrationComplexity) {
                            "Low" { "<span class='badge b-green'>Low</span>" }
                            "Medium" { "<span class='badge b-orange'>Medium</span>" }
                            "High" { "<span class='badge b-red'>High</span>" }
                            default { "<span class='badge b-gray'>—</span>" }
                          }
                          $mcBadge
                        )</td>
                    </tr>
"@
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@

    # Detail cards for candidates
    $w365CandidateDetails = @($w365Analysis | Where-Object { $_.Recommendation -match "W365|Consider" })
    if ($w365CandidateDetails.Count -gt 0) {
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">Detailed Assessment — W365 Candidates</h3>
"@
      foreach ($wd in $w365CandidateDetails) {
        $borderColor = if ($wd.Recommendation -eq "Strong W365 Candidate") { "#4caf50" } else { "#ff9800" }
        $htmlReport += @"
        <div style="border-left:4px solid $borderColor;background:#fafafa;padding:16px 20px;margin:12px 0;border-radius:4px">
            <strong style="font-size:14px">$(Scrub-HostPoolName $wd.HostPoolName)</strong>
            <span class="badge $(if ($wd.Recommendation -match 'Strong') { 'b-green' } else { 'b-orange' })" style="margin-left:8px">$($wd.Recommendation)</span>
            <span style="color:#888;margin-left:8px">Fit Score: $($wd.FitScore)/100</span>
            $(if ($wd.MigrationComplexity) {
              $mcColor = switch ($wd.MigrationComplexity) { "Low" { "b-green" } "Medium" { "b-orange" } "High" { "b-red" } default { "b-gray" } }
              "<span class='badge $mcColor' style='margin-left:8px'>Migration: $($wd.MigrationComplexity)</span>"
            })
            <div style="margin-top:10px;font-size:13px">
                <div style="margin-bottom:6px"><strong style="color:#2e7d32">✅ W365 Advantages:</strong> $(if ($wd.Advantages) { $wd.Advantages } else { "—" })</div>
                <div style="margin-bottom:6px"><strong style="color:#e65100">⚠️ Considerations:</strong> $(if ($wd.Considerations) { $wd.Considerations } else { "None" })</div>
                $(if ($wd.Blockers) { "<div style='margin-bottom:6px'><strong style='color:#c62828'>🚫 Blockers:</strong> $($wd.Blockers)</div>" })
                $(if ($wd.MigrationFactors) { "<div style='margin-bottom:6px'><strong style='color:#5c6bc0'>🔄 Migration Notes:</strong> $($wd.MigrationFactors)</div>" })
                <div style="margin-top:8px;padding-top:8px;border-top:1px solid #eee">
                    <strong>Cost Comparison:</strong>
                    Current AVD: $(if ($wd.AVDEffectiveMonthly) { "`$$($wd.AVDEffectiveMonthly)/mo" } else { "unknown" })$(if ($wd.CostSource -eq 'Actual') { " (actual billing)" } elseif ($wd.HasScalingPlan) { " (with scaling est.)" } else { " (PAYG est.)" })
                     →  W365: $(if ($wd.W365MonthlyTotal) { "`$$($wd.W365MonthlyTotal)/mo" } else { "unknown" })
                    ($($wd.BestW365Plan))
                    <strong style="margin-left:8px">$($wd.CostVerdict)</strong>
                    $(if ($wd.TCOTotal -gt 0) { "<br><span style='font-size:12px;color:#555'>📦 Full TCO: AVD `$$($wd.AVDFullTCO)/mo (compute + `$$($wd.TCOStorageCost) disks + `$$($wd.TCOProfileCost) profiles + `$$($wd.TCONetworkCost) network) vs W365 `$$($wd.W365MonthlyTotal)/mo</span>" })
                    $(if ($wd.UsageBasedPlan -and $wd.UsageSavings -gt 0) { "<br><span style='font-size:12px;color:#1565c0'>📊 Usage-sized: Actual workload fits <strong>$($wd.UsageBasedPlan)</strong> (avg $($wd.PoolAvgCPU)% CPU, $($wd.PoolPeakMem) GB peak RAM) — saves `$$($wd.UsageSavings)/mo vs spec-matched plan</span>" })
                    $(if ($wd.BreakevenUtilization) { "<br><span style='font-size:12px;color:#666'>⚖️ Breakeven: W365 is cheaper if current AVD utilization exceeds $($wd.BreakevenUtilization)% (i.e., VMs run >$($wd.BreakevenUtilization)% of the month)</span>" })
                </div>
            </div>
        </div>
"@
      }
    }
    
    # Notes from the field

    # === TCO Comparison Table ===
    $w365CandidatesTCO = @($w365Analysis | Where-Object { $_.Recommendation -match "W365|Consider" })
    if ($w365CandidatesTCO.Count -gt 0) {
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">💰 Full TCO Comparison — W365 Candidates</h3>
        <div class="alert" style="background:#fff8e1;border-left:4px solid #ffa726;padding:10px 16px;margin-bottom:12px;font-size:12px">
            AVD costs include compute, OS disks, estimated profile storage (Azure Files), and VNet overhead. W365 bundles all of these into a single per-user price.
        </div>
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="tco-table">
                    <thead><tr>
                        <th>Host Pool</th>
                        <th>Users/VMs</th>
                        <th style="border-right:2px solid #ddd">AVD Compute</th>
                        <th>+ Disks</th>
                        <th>+ Profiles</th>
                        <th>+ Network</th>
                        <th style="font-weight:700;border-right:2px solid #ddd">AVD Full TCO</th>
                        <th style="font-weight:700">W365 Total</th>
                        <th>Delta</th>
$(if ($w365UsageSavings -gt 0) { "                        <th><span title='W365 plan matched to actual workload instead of VM spec'>Usage-Sized</span></th>" })
                    </tr></thead>
                    <tbody>
"@
      foreach ($tc in $w365CandidatesTCO) {
        $tcoDelta = if ($tc.TCOCostDelta) {
          if ($tc.TCOCostDelta -lt 0) { "<span style='color:#2e7d32;font-weight:600'>-`$$([math]::Abs($tc.TCOCostDelta))</span>" }
          elseif ($tc.TCOCostDelta -gt 0) { "<span style='color:#c62828'>+`$$($tc.TCOCostDelta)</span>" }
          else { "`$0" }
        } else { "—" }
        $usageCell = if ($tc.UsageBasedPlan -and $tc.UsageSavings -gt 0) {
          "<span style='color:#2e7d32;font-size:11px'>$($tc.UsageBasedPlan)<br>saves `$$($tc.UsageSavings)/mo</span>"
        } elseif ($tc.UsageBasedPlan) {
          "<span style='font-size:11px;color:#888'>$($tc.UsageBasedPlan)</span>"
        } else { "—" }
        $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $tc.HostPoolName)</strong></td>
                        <td>$($tc.VMCount)</td>
                        <td style="border-right:2px solid #f0f0f0">$(if ($tc.AVDEffectiveMonthly) { "`$$($tc.AVDEffectiveMonthly)" } else { "—" })</td>
                        <td>`$$($tc.TCOStorageCost)</td>
                        <td>$(if ($tc.TCOProfileCost -gt 0) { "`$$($tc.TCOProfileCost)" } else { "—" })</td>
                        <td>`$$($tc.TCONetworkCost)</td>
                        <td style="font-weight:600;border-right:2px solid #f0f0f0">$(if ($tc.AVDFullTCO) { "`$$($tc.AVDFullTCO)" } else { "—" })</td>
                        <td style="font-weight:600">$(if ($tc.W365MonthlyTotal) { "`$$($tc.W365MonthlyTotal)" } else { "—" })</td>
                        <td>$tcoDelta</td>
$(if ($w365UsageSavings -gt 0) { "                        <td>$usageCell</td>" })
                    </tr>
"@
      }
      
      # Totals row
      $totalAVDCompute = ($w365CandidatesTCO | Where-Object { $_.AVDEffectiveMonthly } | ForEach-Object { $_.AVDEffectiveMonthly } | Measure-Object -Sum).Sum
      $totalDisks = ($w365CandidatesTCO | ForEach-Object { $_.TCOStorageCost } | Measure-Object -Sum).Sum
      $totalProfiles = ($w365CandidatesTCO | ForEach-Object { $_.TCOProfileCost } | Measure-Object -Sum).Sum
      $totalNetwork = ($w365CandidatesTCO | ForEach-Object { $_.TCONetworkCost } | Measure-Object -Sum).Sum
      $totalAVDFull = ($w365CandidatesTCO | Where-Object { $_.AVDFullTCO } | ForEach-Object { $_.AVDFullTCO } | Measure-Object -Sum).Sum
      $totalW365 = ($w365CandidatesTCO | Where-Object { $_.W365MonthlyTotal } | ForEach-Object { $_.W365MonthlyTotal } | Measure-Object -Sum).Sum
      $totalDelta = if ($totalAVDFull -and $totalW365) { $totalW365 - $totalAVDFull } else { $null }
      $totalDeltaStr = if ($totalDelta) {
        if ($totalDelta -lt 0) { "<span style='color:#2e7d32;font-weight:700'>-`$$([math]::Abs([math]::Round($totalDelta, 0)))</span>" }
        elseif ($totalDelta -gt 0) { "<span style='color:#c62828;font-weight:700'>+`$$([math]::Round($totalDelta, 0))</span>" }
        else { "`$0" }
      } else { "—" }
      
      $htmlReport += @"
                    <tr style="background:#f5f5f5;font-weight:600;border-top:2px solid #ccc">
                        <td>TOTAL (Candidates)</td>
                        <td>$($w365CandidatesTCO | ForEach-Object { $_.VMCount } | Measure-Object -Sum | Select-Object -ExpandProperty Sum)</td>
                        <td style="border-right:2px solid #f0f0f0">`$$([math]::Round($totalAVDCompute, 0))</td>
                        <td>`$$([math]::Round($totalDisks, 0))</td>
                        <td>$(if ($totalProfiles -gt 0) { "`$$([math]::Round($totalProfiles, 0))" } else { "—" })</td>
                        <td>`$$([math]::Round($totalNetwork, 0))</td>
                        <td style="border-right:2px solid #f0f0f0">`$$([math]::Round($totalAVDFull, 0))</td>
                        <td>`$$([math]::Round($totalW365, 0))</td>
                        <td>$totalDeltaStr</td>
$(if ($w365UsageSavings -gt 0) { "                        <td></td>" })
                    </tr>
                    </tbody>
                </table>
            </div>
        </div>
"@

      # === 12-Month Cost Projection ===
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">📅 12-Month Cost Projection</h3>
        <div class="two-col">
            <div class="table-wrap">
                <div style="padding:20px">
                    <table style="width:100%;border-collapse:collapse;font-size:13px">
                        <tr style="border-bottom:1px solid #eee">
                            <td style="padding:8px 0;color:#666">Current AVD Full TCO (monthly)</td>
                            <td style="padding:8px 0;text-align:right;font-weight:600">`$$([math]::Round($totalAVDFull, 0))</td>
                        </tr>
                        <tr style="border-bottom:1px solid #eee">
                            <td style="padding:8px 0;color:#666">Current AVD Full TCO (annual)</td>
                            <td style="padding:8px 0;text-align:right;font-weight:600">`$$([math]::Round($totalAVDFull * 12, 0))</td>
                        </tr>
                        <tr style="border-bottom:2px solid #0078d4">
                            <td style="padding:8px 0;color:#666">W365 Total (annual)</td>
                            <td style="padding:8px 0;text-align:right;font-weight:600">`$$([math]::Round($totalW365 * 12, 0))</td>
                        </tr>
                        <tr>
                            <td style="padding:12px 0;font-weight:700;font-size:14px">Projected Annual Savings</td>
                            <td style="padding:12px 0;text-align:right;font-weight:700;font-size:18px;color:$(if ($totalDelta -and $totalDelta -lt 0) { '#2e7d32' } else { '#c62828' })">$(if ($totalDelta) { if ($totalDelta -lt 0) { "`$$([math]::Round([math]::Abs($totalDelta) * 12, 0))" } else { "-`$$([math]::Round($totalDelta * 12, 0))" } } else { "—" })</td>
                        </tr>
$(if ($w365UsageSavings -gt 0) {
  "                        <tr style='border-top:1px dashed #ccc'><td style='padding:8px 0;color:#1565c0;font-size:12px'>+ Usage-sized plan savings (annual)</td><td style='padding:8px 0;text-align:right;color:#1565c0;font-weight:600'>`$$([math]::Round($w365UsageSavings * 12, 0))</td></tr>"
})
                    </table>
                </div>
            </div>
            <div class="table-wrap">
                <div style="padding:20px;font-size:13px;line-height:1.7">
                    <div style="font-weight:600;margin-bottom:8px">Cost Assumptions</div>
                    <div style="color:#555;margin-bottom:4px">• AVD compute uses $(if ($w365CandidatesTCO[0].CostSource -eq 'Actual') { 'actual billing data' } else { 'PAYG estimates with scaling discount' })</div>
                    <div style="color:#555;margin-bottom:4px">• OS disk costs: P10 Premium `$19.71, E10 Standard SSD `$7.68, S10 HDD `$5.89/mo</div>
                    <div style="color:#555;margin-bottom:4px">• Profile storage: Azure Files Premium at ~`$0.06/GB × 30 GB avg per user</div>
                    <div style="color:#555;margin-bottom:4px">• W365 pricing: East US list price (Enterprise or Frontline as appropriate)</div>
                    <div style="color:#555;margin-bottom:4px">• M365 licensing cost NOT included — assumed users already have E3/E5/F1/F3</div>
                    <div style="color:#e65100;margin-top:8px;font-weight:600">⚠️ Reserved Instances (RI) could reduce AVD costs 30-50%. If RIs are in place, W365 savings will be lower than shown.</div>
                </div>
            </div>
        </div>
"@
    }

    # === Migration Readiness Checklist ===
    $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">✅ Migration Readiness Checklist</h3>
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="w365-checklist">
                    <thead><tr>
                        <th>Host Pool</th>
                        <th>Fit</th>
                        <th><span title="Entra ID / Azure AD join status">Identity</span></th>
                        <th><span title="Ready for Intune management">Intune</span></th>
                        <th><span title="Custom or marketplace image">Image</span></th>
                        <th><span title="FSLogix profile container dependency">Profiles</span></th>
                        <th><span title="Within W365 size limits">SKU Fit</span></th>
                        <th>Complexity</th>
                        <th>Pilot Score</th>
                    </tr></thead>
                    <tbody>
"@
    foreach ($cl in ($w365Analysis | Sort-Object FitScore -Descending)) {
      $checkIdentity = if ($cl.IntuneReady) { "<span style='color:#2e7d32'>✅ $($cl.JoinType)</span>" } else { "<span style='color:#e65100'>⚠️ $($cl.JoinType)</span>" }
      $checkIntune = if ($cl.IntuneReady) { "<span style='color:#2e7d32'>✅ Likely ready</span>" } else { "<span style='color:#c62828'>❌ Needs setup</span>" }
      $checkImage = if ($cl.MigrationFactors -match 'Custom|gallery') { "<span style='color:#e65100'>⚠️ Custom image</span>" } else { "<span style='color:#2e7d32'>✅ Marketplace</span>" }
      $checkProfile = if ($cl.MigrationFactors -match 'FSLogix') { 
        if ($cl.HostPoolType -match 'Personal') { "<span style='color:#2e7d32'>✅ One-time migrate</span>" }
        else { "<span style='color:#2e7d32'>✅ Eliminated</span>" }
      } else { "<span style='color:#888'>— No FSLogix detected</span>" }
      $checkSku = if ($cl.HasGPU) { "<span style='color:#c62828'>❌ GPU required</span>" }
                 elseif ($cl.Blockers -match 'exceed.*vCPU') { "<span style='color:#c62828'>❌ Exceeds 16 vCPU</span>" }
                 elseif ($cl.Blockers -match 'Multi-session') { "<span style='color:#e65100'>⚠️ Multi-session</span>" }
                 else { "<span style='color:#2e7d32'>✅ Within limits</span>" }
      $mcBadge = switch ($cl.MigrationComplexity) { "Low" { "<span class='badge b-green'>Low</span>" } "Medium" { "<span class='badge b-orange'>Medium</span>" } "High" { "<span class='badge b-red'>High</span>" } default { "—" } }
      $pilotBadge = if ($cl.PilotScore -ge 70) { "<span class='badge b-green'>$($cl.PilotScore)</span>" } elseif ($cl.PilotScore -ge 40) { "<span class='badge b-orange'>$($cl.PilotScore)</span>" } else { "<span style='color:#888'>$($cl.PilotScore)</span>" }
      $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $cl.HostPoolName)</strong></td>
                        <td>$($cl.FitScore)/100</td>
                        <td style="font-size:11px">$checkIdentity</td>
                        <td style="font-size:11px">$checkIntune</td>
                        <td style="font-size:11px">$checkImage</td>
                        <td style="font-size:11px">$checkProfile</td>
                        <td style="font-size:11px">$checkSku</td>
                        <td>$mcBadge</td>
                        <td>$pilotBadge</td>
                    </tr>
"@
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@

    # === Per-User Cost (v4.1) ===
    if ($perUserCost.Count -gt 0) {
      $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">👤 Cost Per User — AVD vs W365</h3>
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="peruser-table">
                    <thead><tr>
                        <th>Host Pool</th>
                        <th>VMs</th>
                        <th>Users</th>
                        <th>Drain</th>
                        <th>AVD Monthly</th>
                        <th>AVD/User</th>
                        <th>W365/User</th>
                        <th>Verdict</th>
                    </tr></thead>
                    <tbody>
"@
      foreach ($pu in ($perUserCost | Sort-Object CostPerUser -Descending)) {
        $w365Match = $w365Analysis | Where-Object { $_.HostPoolName -eq $pu.HostPoolName } | Select-Object -First 1
        $w365PerUser = if ($w365Match -and $w365Match.W365MonthlyPerUser) { "`$$($w365Match.W365MonthlyPerUser)" } else { "—" }
        $verdict = if ($w365Match -and $w365Match.W365MonthlyPerUser -and $pu.CostPerUser -gt $w365Match.W365MonthlyPerUser) { "<span style='color:#2e7d32;font-weight:600'>W365 cheaper</span>" }
                   elseif ($w365Match -and $w365Match.W365MonthlyPerUser -and $pu.CostPerUser -lt $w365Match.W365MonthlyPerUser) { "<span style='color:#0078d4'>AVD cheaper</span>" }
                   else { "—" }
        $drainBadge = if ($pu.DrainModeHosts -gt 0) { "<span class='badge b-orange'>$($pu.DrainModeHosts)</span>" } else { "0" }
        $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $pu.HostPoolName)</strong></td>
                        <td>$($pu.VMCount)</td>
                        <td>$($pu.UserCount)<br><span style='font-size:10px;color:#888'>$($pu.UserCountSource)</span></td>
                        <td>$drainBadge</td>
                        <td>`$$($pu.TotalMonthly)</td>
                        <td style="font-weight:600">`$$($pu.CostPerUser)</td>
                        <td>$w365PerUser</td>
                        <td>$verdict</td>
                    </tr>
"@
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # === W365 Feature Gap Analysis (v4.1) ===
    # Detect which gaps are relevant to this environment
    $relevantGaps = [System.Collections.Generic.List[object]]::new()
    $hasMultiSession = @($w365Analysis | Where-Object { $_.Blockers -match 'Multi-session' }).Count -gt 0
    $hasRemoteAppPools = @($w365Analysis | Where-Object { $_.AppGroupType -eq 'RailApplications' }).Count -gt 0
    $hasGpuPools = @($w365Analysis | Where-Object { $_.HasGPU }).Count -gt 0
    $hasScalingPlans = @($w365Analysis | Where-Object { $_.HasScalingPlan }).Count -gt 0
    $hasPrivateEndpoints = @($networkFindings | Where-Object { $_.Detail -match 'Private' }).Count -gt 0

    if ($hasMultiSession) { $relevantGaps.Add($w365FeatureGaps["MultiSession"]) }
    if ($hasRemoteAppPools) { $relevantGaps.Add($w365FeatureGaps["RemoteApp"]) }
    if ($hasGpuPools) { $relevantGaps.Add($w365FeatureGaps["GPU"]) }
    if ($hasScalingPlans) { $relevantGaps.Add($w365FeatureGaps["Autoscale"]) }
    $relevantGaps.Add($w365FeatureGaps["AppAttach"])
    if ($hasPrivateEndpoints) { $relevantGaps.Add($w365FeatureGaps["PrivateLink"]) }
    # Always add informational items
    $relevantGaps.Add($w365FeatureGaps["TeamsOptimization"])
    $relevantGaps.Add($w365FeatureGaps["RDPShortpath"])
    $relevantGaps.Add($w365FeatureGaps["MMR"])
    $relevantGaps.Add($w365FeatureGaps["MultiMonitor"])
    $relevantGaps.Add($w365FeatureGaps["CustomNetworking"])

    $htmlReport += @"
        <h3 style="margin:24px 0 16px 0;font-size:16px;color:#333">🔍 W365 Feature Gap Analysis</h3>
        <div class="alert" style="background:#f3e5f5;border-left:4px solid #9c27b0;padding:10px 16px;margin-bottom:12px;font-size:12px">
            Comparison of capabilities between your current AVD configuration and Windows 365. Items marked with ⚠️ are detected in your environment and may impact migration.
        </div>
        <div class="table-wrap">
            <div class="table-scroll">
                <table id="feature-gap-table">
                    <thead><tr>
                        <th style="width:15%">Feature</th>
                        <th style="width:30%">Azure Virtual Desktop</th>
                        <th style="width:30%">Windows 365</th>
                        <th style="width:10%">Impact</th>
                        <th style="width:15%">Detected</th>
                    </tr></thead>
                    <tbody>
"@
    foreach ($fg in $relevantGaps) {
      $impactBadge = switch ($fg.Impact) { "High" { "<span class='badge b-red'>High</span>" } "Medium" { "<span class='badge b-orange'>Medium</span>" } "Low" { "<span class='badge b-green'>Low</span>" } default { "—" } }
      $detectedIcon = if ($fg.Detection -match 'Informational') { "<span style='color:#888'>ℹ️ Info</span>" } else { "<span style='color:#e65100'>⚠️ Detected</span>" }
      $htmlReport += @"
                    <tr>
                        <td><strong>$($fg.Feature)</strong></td>
                        <td style="font-size:12px">$($fg.AVD)</td>
                        <td style="font-size:12px">$($fg.W365)</td>
                        <td>$impactBadge</td>
                        <td style="font-size:11px">$detectedIcon</td>
                    </tr>
"@
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@

    $htmlReport += @"
        <div style="background:#e8f5e9;border-radius:8px;padding:20px;margin-top:24px">
            <h3 style="margin:0 0 12px 0;font-size:16px;color:#333">📝 W365 Migration Considerations</h3>
            <div style="font-size:13px;line-height:1.8">
                <div>• <strong>Licensing prerequisite:</strong> W365 requires Microsoft 365 E3/E5, F1/F3, or Business Premium. Frontline plans require F1/F3. Factor existing M365 licensing into TCO — if users already have E3/E5, the W365 add-on cost is the only incremental spend.</div>
                <div>• <strong>Hybrid approach:</strong> Many organizations run both — W365 for simple knowledge workers and personal desktops, AVD for multi-session pools, RemoteApp, and specialized workloads. This assessment identifies pools where the economics favor each.</div>
                <div>• <strong>FSLogix elimination:</strong> W365 Cloud PCs are persistent — no FSLogix profile containers needed. This eliminates Azure Files storage costs, profile load latency, and the #1 source of session host troubleshooting tickets.</div>
                <div>• <strong>Networking:</strong> W365 Enterprise supports Azure Network Connection (ANC) for hybrid connectivity. Microsoft Hosted Network is simpler but limits network control. If your AVD hosts currently route through custom VNets with private endpoints, plan the ANC configuration early.</div>
                <div>• <strong>Management shift:</strong> W365 is managed via Intune, not Azure portal. Security baselines, app deployment, and compliance policies move from GPO/Azure to Intune. If your team is Azure-centric, budget time for the operational transition.</div>
                <div>• <strong>Scaling model:</strong> W365 is fixed-cost per user per month — great for predictable workloads, but expensive for seasonal/burst use. AVD autoscale wins when utilization fluctuates. Check the breakeven % on each candidate above.</div>
                <div>• <strong>GPU and high-compute:</strong> W365 caps at 16 vCPU / 64 GB RAM with no GPU option. GPU-accelerated and high-memory workloads must stay on AVD.</div>
$(if ($w365Strong.Count -gt 0 -or $w365Consider.Count -gt 0) {
  "                <div style='margin-top:12px;padding-top:12px;border-top:1px solid #c8e6c9'><strong>🎯 Recommended next steps:</strong> (1) Validate M365 licensing coverage for candidate pools. (2) Run a 2-week pilot with the strongest candidate pool. (3) Compare actual W365 user experience metrics (login time, app launch) against current AVD baseline. (4) Plan phased migration starting with Low-complexity pools.</div>"
})
            </div>
        </div>
    </div>
"@
  }

  # ========== SECURITY POSTURE ==========
  if ($securityPosture.Count -gt 0) {
    $htmlReport += @"
    
    <!-- ========== SECURITY POSTURE ========== -->
    <div class="section" id="sec-security">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Avg Security Score</div>
                <div class="card-value $(if ($summary.SecurityPosture.AvgSecurityScore -ge 75) { 'green' } elseif ($summary.SecurityPosture.AvgSecurityScore -ge 60) { 'yellow' } else { 'red' })">$($summary.SecurityPosture.AvgSecurityScore)/100</div>
            </div>
            <div class="card">
                <div class="card-label">Pools Analyzed</div>
                <div class="card-value blue">$($summary.SecurityPosture.HostPoolsAnalyzed)</div>
            </div>
            <div class="card">
                <div class="card-label">Below Grade C</div>
                <div class="card-value $(if ($summary.SecurityPosture.PoolsBelowGradeC -gt 0) { 'red' } else { 'green' })">$($summary.SecurityPosture.PoolsBelowGradeC)</div>
            </div>
        </div>

        <div class="table-wrap">
            <div class="table-title">🔒 Security Posture by Host Pool</div>
            <div class="table-scroll">
                <table id="security-table">
                    <thead><tr>
                        <th onclick="sortTable('security-table',0)">Host Pool</th>
                        <th onclick="sortTable('security-table',1)">VMs</th>
                        <th onclick="sortTable('security-table',2)">Score</th>
                        <th onclick="sortTable('security-table',3)">Grade</th>
                        <th>Trusted Launch</th>
                        <th>Secure Boot</th>
                        <th>vTPM</th>
                        <th>Host Encryption</th>
                        <th>AccelNet</th>
                        <th>Ephemeral Disk</th>
                    </tr></thead>
                    <tbody>
"@
    foreach ($sp in ($securityPosture | Sort-Object SecurityScore)) {
      $gradeClass = switch ($sp.SecurityGrade) { "A" { "b-green" }; "B" { "b-green" }; "C" { "b-yellow" }; "D" { "b-red" }; default { "b-red" } }
      $pctColor = { param($p) if ($p -eq 100) { "#2e7d32" } elseif ($p -ge 50) { "#f57c00" } else { "#c62828" } }
      $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $sp.HostPoolName)</strong></td>
                        <td>$($sp.VMCount)</td>
                        <td><strong>$($sp.SecurityScore)</strong></td>
                        <td><span class="badge $gradeClass">$($sp.SecurityGrade)</span></td>
                        <td style="color:$(& $pctColor $sp.TrustedLaunchPct)">$($sp.TrustedLaunchPct)%</td>
                        <td style="color:$(& $pctColor $sp.SecureBootPct)">$($sp.SecureBootPct)%</td>
                        <td style="color:$(& $pctColor $sp.VTpmPct)">$($sp.VTpmPct)%</td>
                        <td style="color:$(& $pctColor $sp.HostEncryptionPct)">$($sp.HostEncryptionPct)%</td>
                        <td style="color:$(& $pctColor $sp.AccelNetPct)">$($sp.AccelNetPct)%</td>
                        <td style="color:$(& $pctColor $sp.EphemeralDiskPct)">$($sp.EphemeralDiskPct)%</td>
                    </tr>
"@
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@

    # Findings detail
    $spWithFindings = @($securityPosture | Where-Object { $_.Findings })
    if ($spWithFindings.Count -gt 0) {
      $htmlReport += @"
        <div class="table-wrap">
            <div class="table-title">🔍 Security Findings</div>
            <div class="table-scroll">
                <table id="secfindings-table">
                    <thead><tr><th>Host Pool</th><th>Grade</th><th>Findings</th></tr></thead>
                    <tbody>
"@
      foreach ($sf in $spWithFindings) {
        $sfGradeClass = switch ($sf.SecurityGrade) { "A" { "b-green" }; "B" { "b-green" }; "C" { "b-yellow" }; default { "b-red" } }
        $htmlReport += "                    <tr><td><strong>$(Scrub-HostPoolName $sf.HostPoolName)</strong></td><td><span class='badge $sfGradeClass'>$($sf.SecurityGrade)</span></td><td style='font-size:12px'>$($sf.Findings -replace '; ', '<br>')</td></tr>`n"
      }
      $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    $htmlReport += @"
        <div style="background:#e8f5e9;border-radius:8px;padding:16px;margin-top:16px">
            <h3 style="margin:0 0 12px 0;font-size:16px;color:#333">📝 Security Hardening Notes</h3>
            <div style="padding:0;font-size:13px;line-height:1.7;color:#444">
                <p><strong>Trusted Launch</strong> — Requires Gen2 VM images. Provides Secure Boot + vTPM + boot integrity monitoring. New deployments should always use Trusted Launch. Existing VMs require redeployment.</p>
                <p><strong>Host-Based Encryption</strong> — Encrypts temp disks and OS/data disk caches at the host level. Zero performance impact. Enable via <code>Set-AzVMOperatingSystem</code> or ARM template <code>securityProfile.encryptionAtHost: true</code>.</p>
                <p><strong>Ephemeral OS Disks</strong> — For pooled session hosts, ephemeral disks eliminate persistent OS disk costs, speed up reimage operations, and reduce attack surface since nothing persists between reimages.</p>
            </div>
        </div>
    </div>
"@
  }

  # ========== UX SCORES ==========
  if ($uxScores.Count -gt 0) {
    $htmlReport += @"
    
    <!-- ========== UX SCORES ========== -->
    <div class="section" id="sec-ux">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Avg UX Score</div>
                <div class="card-value $(if ($summary.UserExperience.AvgUXScore -ge 75) { 'green' } elseif ($summary.UserExperience.AvgUXScore -ge 60) { 'yellow' } else { 'red' })">$($summary.UserExperience.AvgUXScore)/100</div>
            </div>
            <div class="card">
                <div class="card-label">Pools Scored</div>
                <div class="card-value blue">$($summary.UserExperience.HostPoolsScored)</div>
            </div>
            <div class="card">
                <div class="card-label">Below Grade C</div>
                <div class="card-value $(if ($summary.UserExperience.PoolsBelowGradeC -gt 0) { 'red' } else { 'green' })">$($summary.UserExperience.PoolsBelowGradeC)</div>
            </div>
        </div>

        <div class="table-wrap">
            <div class="table-title">🎯 User Experience Score by Host Pool</div>
            <div class="table-scroll">
                <table id="ux-table">
                    <thead><tr>
                        <th onclick="sortTable('ux-table',0)">Host Pool</th>
                        <th onclick="sortTable('ux-table',1)">UX Score</th>
                        <th onclick="sortTable('ux-table',2)">Grade</th>
                        <th onclick="sortTable('ux-table',3)">Profile Load</th>
                        <th>P95 Load (s)</th>
                        <th onclick="sortTable('ux-table',5)">Connection</th>
                        <th>Avg RTT (ms)</th>
                        <th onclick="sortTable('ux-table',7)">Disconnects</th>
                        <th onclick="sortTable('ux-table',8)">Errors</th>
                        <th>Components</th>
                    </tr></thead>
                    <tbody>
"@
    foreach ($ux in ($uxScores | Sort-Object UXScore)) {
      $uxGradeClass = switch ($ux.UXGrade) { "A" { "b-green" }; "B" { "b-green" }; "C" { "b-yellow" }; "D" { "b-red" }; default { "b-red" } }
      $subScoreColor = { param($s) if ($null -eq $s) { "#999" } elseif ($s -ge 20) { "#2e7d32" } elseif ($s -ge 12) { "#f57c00" } else { "#c62828" } }
      $htmlReport += @"
                    <tr>
                        <td><strong>$(Scrub-HostPoolName $ux.HostPoolName)</strong></td>
                        <td><strong>$($ux.UXScore)</strong></td>
                        <td><span class="badge $uxGradeClass">$($ux.UXGrade)</span></td>
                        <td style="color:$(& $subScoreColor $ux.ProfileLoadScore)">$(if ($null -ne $ux.ProfileLoadScore) { "$($ux.ProfileLoadScore)/25" } else { "—" })</td>
                        <td>$(if ($null -ne $ux.ProfileP95Sec) { "$($ux.ProfileP95Sec)s" } else { "—" })</td>
                        <td style="color:$(& $subScoreColor $ux.ConnectionScore)">$(if ($null -ne $ux.ConnectionScore) { "$($ux.ConnectionScore)/25" } else { "—" })</td>
                        <td>$(if ($null -ne $ux.AvgRTTms) { "$($ux.AvgRTTms)ms" } else { "—" })</td>
                        <td style="color:$(& $subScoreColor $ux.DisconnectScore)">$(if ($null -ne $ux.DisconnectScore) { "$($ux.DisconnectScore)/25" } else { "—" })</td>
                        <td style="color:$(& $subScoreColor $ux.ErrorScore)">$(if ($null -ne $ux.ErrorScore) { "$($ux.ErrorScore)/25" } else { "—" })</td>
                        <td style="color:#888">$($ux.ComponentsAvailable)/4</td>
                    </tr>
"@
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
        <div class="alert alert-info">
            <strong>ℹ️ Scoring:</strong> Each component scores 0-25 points. The UX Score scales to 100 based on available data.
            Components without Log Analytics data are skipped (not penalized). Profile Load = P95 logon time, Connection = avg RTT,
            Disconnects = abnormal disconnect rate, Errors = connection error count. Grade: A (90+), B (75+), C (60+), D (40+), F (<40).
        </div>
    </div>
"@
  }

  # ========== ORPHANED RESOURCES ==========
  if ($orphanedResources.Count -gt 0) {
    $orphanedDisks = @($orphanedResources | Where-Object { $_.ResourceType -eq "Disk" })
    $orphanedNics = @($orphanedResources | Where-Object { $_.ResourceType -eq "NIC" })
    $orphanedPips = @($orphanedResources | Where-Object { $_.ResourceType -eq "PublicIP" })
    $htmlReport += @"
    
    <!-- ========== ORPHANED RESOURCES ========== -->
    <div class="section" id="sec-orphans">
        <div class="card-grid">
            <div class="card">
                <div class="card-label">Total Orphaned</div>
                <div class="card-value orange">$($orphanedResources.Count)</div>
            </div>
            <div class="card">
                <div class="card-label">Est. Monthly Waste</div>
                <div class="card-value red">~`$$orphanedWaste</div>
            </div>
            <div class="card">
                <div class="card-label">Disks</div>
                <div class="card-value $(if ($orphanedDisks.Count -gt 0) { 'orange' } else { 'green' })">$($orphanedDisks.Count)</div>
            </div>
            <div class="card">
                <div class="card-label">NICs / Public IPs</div>
                <div class="card-value $(if (($orphanedNics.Count + $orphanedPips.Count) -gt 0) { 'orange' } else { 'green' })">$($orphanedNics.Count) / $($orphanedPips.Count)</div>
            </div>
        </div>

        <div class="table-wrap">
            <div class="table-title">🗑️ Orphaned Resources</div>
            <div class="table-scroll">
                <table id="orphan-table">
                    <thead><tr>
                        <th onclick="sortTable('orphan-table',0)">Resource Name</th>
                        <th onclick="sortTable('orphan-table',1)">Type</th>
                        <th onclick="sortTable('orphan-table',2)">Resource Group</th>
                        <th onclick="sortTable('orphan-table',3)">Details</th>
                        <th onclick="sortTable('orphan-table',4)">Est. Cost/mo</th>
                    </tr></thead>
                    <tbody>
"@
    foreach ($orph in ($orphanedResources | Sort-Object EstMonthlyCost -Descending)) {
      $typeBadge = switch ($orph.ResourceType) { "Disk" { "b-blue" }; "NIC" { "b-yellow" }; "PublicIP" { "b-red" }; default { "b-gray" } }
      $htmlReport += @"
                    <tr>
                        <td><strong>$($orph.ResourceName)</strong></td>
                        <td><span class="badge $typeBadge">$($orph.ResourceType)</span></td>
                        <td style="font-size:12px">$($orph.ResourceGroup)</td>
                        <td style="font-size:12px">$($orph.Details)</td>
                        <td>$(if ($orph.EstMonthlyCost -gt 0) { "`$$($orph.EstMonthlyCost)" } else { "—" })</td>
                    </tr>
"@
    }
    $htmlReport += @"
                    </tbody>
                </table>
            </div>
        </div>
        <div class="alert alert-warning">
            <strong>💡 Quick Win:</strong> Deleting orphaned resources is low-risk, low-effort, and saves money immediately. Verify each resource is truly unneeded, then delete via Portal or <code>Remove-AzDisk</code> / <code>Remove-AzNetworkInterface</code> / <code>Remove-AzPublicIpAddress</code>.
        </div>
    </div>
"@
  }

  # Incident section
  if ($summary.IncidentWindowAnalysis.Enabled) {
    $htmlReport += @"
    
    <!-- ========== INCIDENT ========== -->
    <div class="section" id="sec-incident">
        <div class="alert alert-info"><strong>Incident Window:</strong> $($summary.IncidentWindowAnalysis.TimeWindow)</div>
        <div class="card-grid">
            <div class="card">
                <div class="card-label">VMs Analyzed</div>
                <div class="card-value blue">$($summary.IncidentWindowAnalysis.VMsAnalyzed)</div>
            </div>
            <div class="card">
                <div class="card-label">Critical</div>
                <div class="card-value red">$($summary.IncidentWindowAnalysis.CriticalIssues)</div>
            </div>
            <div class="card">
                <div class="card-label">High Impact</div>
                <div class="card-value orange">$($summary.IncidentWindowAnalysis.HighIssues)</div>
            </div>
            <div class="card">
                <div class="card-label">Avg CPU Increase</div>
                <div class="card-value yellow">$($summary.IncidentWindowAnalysis.AvgCPUIncrease)%</div>
            </div>
        </div>
        <div class="alert alert-info">See <code>ENHANCED-Incident-Comparative-Analysis.csv</code> for per-VM incident impact details.</div>
    </div>
"@
  }

  # Footer and JavaScript
  $htmlReport += @"
    
    <div class="footer">
        <p>Generated by Enhanced AVD Evidence Pack v3.0.0</p>
        <p>Report files are in the output folder. Review CSVs for full detail.</p>
    </div>
</div>

<script>
function showSection(name) {
    document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('sec-' + name).classList.add('active');
    event.target.classList.add('active');
}

function toggleDetail(row) {
    const detail = row.nextElementSibling;
    if (detail && detail.classList.contains('detail-row')) {
        detail.style.display = detail.style.display === 'none' ? '' : 'none';
    }
}

function filterTable(tableId, query) {
    const rows = document.getElementById(tableId).querySelectorAll('tbody tr');
    query = query.toLowerCase();
    rows.forEach(r => {
        if (r.classList.contains('detail-row')) { r.style.display = 'none'; return; }
        r.style.display = r.textContent.toLowerCase().includes(query) ? '' : 'none';
    });
}

function filterRs(btn, action) {
    document.querySelectorAll('.toolbar .filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const rows = document.getElementById('rs-table').querySelectorAll('tbody tr');
    rows.forEach(r => {
        if (r.classList.contains('detail-row')) { r.style.display = 'none'; return; }
        if (action === 'all') { r.style.display = ''; }
        else { r.style.display = r.dataset.action === action ? '' : 'none'; }
    });
}

function sortTable(tableId, colIdx) {
    const table = document.getElementById(tableId);
    const tbody = table.querySelector('tbody');
    const allRows = Array.from(tbody.querySelectorAll('tr'));
    const th = table.querySelectorAll('th')[colIdx];
    const isAsc = th.classList.contains('sorted-asc');
    
    table.querySelectorAll('th').forEach(h => { h.classList.remove('sorted-asc','sorted-desc'); });
    th.classList.add(isAsc ? 'sorted-desc' : 'sorted-asc');
    
    // Group data rows with their detail rows
    const groups = [];
    allRows.forEach(r => {
        if (r.classList.contains('detail-row')) {
            if (groups.length > 0) groups[groups.length - 1].push(r);
        } else {
            groups.push([r]);
        }
    });
    
    groups.sort((a, b) => {
        let aVal = a[0].cells[colIdx]?.textContent.trim() || '';
        let bVal = b[0].cells[colIdx]?.textContent.trim() || '';
        let aNum = parseFloat(aVal.replace(/[^0-9.\-]/g, ''));
        let bNum = parseFloat(bVal.replace(/[^0-9.\-]/g, ''));
        if (!isNaN(aNum) && !isNaN(bNum)) { return isAsc ? bNum - aNum : aNum - bNum; }
        return isAsc ? bVal.localeCompare(aVal) : aVal.localeCompare(bVal);
    });
    groups.forEach(g => g.forEach(r => tbody.appendChild(r)));
}
</script>
</body>
</html>
"@

  $htmlReport | Out-File (Join-Path $outFolder "ENHANCED-Analysis-Report.html") -Encoding utf8
  
  # Export enriched cross-region analysis (built during HTML generation)
  if ($crossRegionAnalysis.Count -gt 0) {
    $crossRegionAnalysis | Export-Csv (Join-Path $outFolder "ENHANCED-CrossRegion-Analysis.csv") -NoTypeInformation
  }
  
  # Export disconnect reason analysis
  if ($disconnectReasonData.Count -gt 0) {
    $disconnectReasonData | Export-Csv (Join-Path $outFolder "ENHANCED-Disconnect-Reasons.csv") -NoTypeInformation
  }
  if ($disconnectsByHostData.Count -gt 0) {
    $disconnectsByHostData | Export-Csv (Join-Path $outFolder "ENHANCED-Disconnects-ByHost.csv") -NoTypeInformation
  }
  
  Write-ProgressSection -Section "Step 8: Generating HTML Report" -Status Complete -Message "Report saved: ENHANCED-Analysis-Report.html"
}
else {
  Write-ProgressSection -Section "Step 8: Generating HTML Report" -Status Skip -Message "-GenerateHtmlReport not specified (CSV files still generated)"
}

# =========================================================
# Final Summary Output
# =========================================================
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Green
Write-Host "                  ✓ ANALYSIS COMPLETE" -ForegroundColor Green
Write-Host "========================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Output Location:" -ForegroundColor Cyan
Write-Host "  $outFolder" -ForegroundColor White
if ($CreateZip) {
  Write-Host "  $zipPath" -ForegroundColor White
}
Write-Host ""

Write-Host "Environment Summary:" -ForegroundColor Cyan
Write-Host "  • Total VMs Analyzed: $($summary.VMs)" -ForegroundColor White
Write-Host "  • Host Pools: $($summary.HostPools)" -ForegroundColor White
Write-Host "  • Session Hosts: $($summary.SessionHosts)" -ForegroundColor White
if ($summary.ScaleSets -gt 0) {
  Write-Host "  • Scale Sets: $($summary.ScaleSets) ($($summary.ScaleSetInstances) instances)" -ForegroundColor White
}
Write-Host ""

Write-Host "Key Findings:" -ForegroundColor Yellow
Write-Host "  • Downsize Candidates: $($summary.RightSizingAnalysis.DownsizeCandidates) VMs" -ForegroundColor $(if ($summary.RightSizingAnalysis.DownsizeCandidates -gt 0) { "Yellow" } else { "Gray" })
Write-Host "  • Upsize Candidates: $($summary.RightSizingAnalysis.UpsizeCandidates) VMs" -ForegroundColor $(if ($summary.RightSizingAnalysis.UpsizeCandidates -gt 0) { "Red" } else { "Gray" })
Write-Host "  • Appropriately Sized: $($summary.RightSizingAnalysis.AppropriatelySized) VMs" -ForegroundColor Green
Write-Host ""

Write-Host "Cost Optimization:" -ForegroundColor Yellow
Write-Host "  • Estimated Monthly Savings: ~`$$($summary.RightSizingAnalysis.PotentialMonthlySavings)" -ForegroundColor $(if ([int]$summary.RightSizingAnalysis.PotentialMonthlySavings -gt 1000) { "Green" } else { "Gray" })
Write-Host "  • Estimated Annual Savings: ~`$$($summary.RightSizingAnalysis.PotentialAnnualSavings)" -ForegroundColor $(if ([int]$summary.RightSizingAnalysis.PotentialAnnualSavings -gt 12000) { "Green" } else { "Gray" })
Write-Host "  • Savings Percentage: $($summary.RightSizingAnalysis.SavingsPercentage)" -ForegroundColor White
Write-Host "  (Remember: These are estimates - validate against actual billing)" -ForegroundColor Gray
Write-Host ""

Write-Host "Zone Resiliency:" -ForegroundColor Yellow
$resiliencyScore = $summary.ZoneResiliencyAnalysis.AverageResiliencyScore
$resiliencyColor = if ($resiliencyScore -ge 75) { "Green" } elseif ($resiliencyScore -ge 40) { "Yellow" } else { "Red" }
Write-Host "  • Average Score: $resiliencyScore/100" -ForegroundColor $resiliencyColor
Write-Host "  • High Resiliency (75-100): $($summary.ZoneResiliencyAnalysis.HighResiliency) host pools" -ForegroundColor Green
Write-Host "  • Medium Resiliency (40-74): $($summary.ZoneResiliencyAnalysis.MediumResiliency) host pools" -ForegroundColor Yellow
Write-Host "  • Low Resiliency (0-39): $($summary.ZoneResiliencyAnalysis.LowResiliency) host pools" -ForegroundColor Red
Write-Host ""

if ($IncludeAzureAdvisor -and (SafeCount $advisorRecommendations) -gt 0) {
  Write-Host "Azure Advisor:" -ForegroundColor Yellow
  Write-Host "  • Total Recommendations: $(SafeCount $advisorRecommendations)" -ForegroundColor White
  $highImpact = SafeCount ($advisorRecommendations | Where-Object { $_.Impact -eq "High" })
  $mediumImpact = SafeCount ($advisorRecommendations | Where-Object { $_.Impact -eq "Medium" })
  if ($highImpact -gt 0) {
    Write-Host "  • High Impact: $highImpact" -ForegroundColor Red
  }
  if ($mediumImpact -gt 0) {
    Write-Host "  • Medium Impact: $mediumImpact" -ForegroundColor Yellow
  }
  Write-Host ""
}

# New findings (v3.0.0)
Write-Host "Session Host Health:" -ForegroundColor Yellow
Write-Host "  • Healthy: $($summary.SessionHostHealth.Healthy) / $($summary.SessionHostHealth.TotalHosts)" -ForegroundColor $(if ($summary.SessionHostHealth.Healthy -eq $summary.SessionHostHealth.TotalHosts) { "Green" } else { "Yellow" })
if ($stuckHosts -gt 0) {
  Write-Host "  • Stuck in Drain: $stuckHosts" -ForegroundColor Red
}
if ($unavailableHosts -gt 0) {
  Write-Host "  • Issues Detected: $unavailableHosts" -ForegroundColor Red
}
Write-Host ""

if ($premiumOnPooled -gt 0 -or $eligibleNotEnabled -gt 0 -or $multiVersionImages -gt 0) {
  Write-Host "Infrastructure Findings:" -ForegroundColor Yellow
  if ($premiumOnPooled -gt 0) {
    Write-Host "  • Premium SSD on Pooled: $premiumOnPooled VMs (cost savings available)" -ForegroundColor Yellow
  }
  if ($nonEphemeral -gt 0) {
    Write-Host "  • Non-Ephemeral OS Disk: $nonEphemeral pooled VMs" -ForegroundColor Yellow
  }
  if ($eligibleNotEnabled -gt 0) {
    Write-Host "  • AccelNet Not Enabled: $eligibleNotEnabled eligible VMs" -ForegroundColor Yellow
  }
  if ($multiVersionImages -gt 0) {
    Write-Host "  • Image Version Drift: $multiVersionImages image groups" -ForegroundColor Yellow
  }
  Write-Host ""
}

Write-Host "Next Steps:" -ForegroundColor Cyan
Write-Host "  1. Review ENHANCED-Executive-Summary.txt for detailed overview" -ForegroundColor White
Write-Host "  2. Check ENHANCED-VM-RightSizing-Recommendations.csv for specific actions" -ForegroundColor White
Write-Host "  3. Address any VMs with 'Upsize' recommendations immediately" -ForegroundColor White
Write-Host "  4. Plan implementation of high-confidence downsize recommendations" -ForegroundColor White
if ($summary.ZoneResiliencyAnalysis.LowResiliency -gt 0) {
  Write-Host "  5. Improve zone resiliency for low-scoring host pools" -ForegroundColor White
}
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Green
Write-Host ""
