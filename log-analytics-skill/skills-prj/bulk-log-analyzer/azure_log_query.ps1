<#
.SYNOPSIS
    Azure Monitor Log Query Script - China Cloud Environment
    Query Log Analytics using official Az.OperationalInsights module
.DESCRIPTION
    Authenticate via Az module, execute KQL queries using official Az.OperationalInsights library
    Supports interactive browser login and device code flow
.EXAMPLE
    .\azure_log_query.ps1
    .\azure_log_query.ps1 -Query "AzureActivity | top 20 by TimeGenerated desc" -Hours 48
    .\azure_log_query.ps1 -UseDeviceCode
    .\azure_log_query.ps1 -Query "AuditGeneralDCR_CL | take 100" -ExportCsv ".\results.csv"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$Query = "AuditGeneralDCR_CL",

    [Parameter(Mandatory = $false)]
    [string]$TableName = "",

    [Parameter(Mandatory = $false)]
    [string]$StartTime = "",

    [Parameter(Mandatory = $false)]
    [string]$EndTime = "",

    [Parameter(Mandatory = $false)]
    [int]$Hours = 1,

    [Parameter(Mandatory = $false)]
    [string]$WorkspaceId = "703a5771-97fc-4bf3-a585-f607d18c4479",

    [Parameter(Mandatory = $false)]
    [string]$TenantId = "420c4dab-8603-402f-afe0-75bc28c51c13",

    [Parameter(Mandatory = $false)]
    [switch]$UseDeviceCode,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeStats,

    [Parameter(Mandatory = $false)]
    [switch]$ForceLogin,

    [Parameter(Mandatory = $false)]
    [string]$ExportCsv
)

# ============================================================
# Configuration
# ============================================================
$AzureEnvironment = "AzureChinaCloud"
$ModuleName = "Az.OperationalInsights"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $ScriptDir 'log-analyzer-core.ps1')

# ============================================================
# Functions
# ============================================================

function Initialize-Modules {
    <#
    .SYNOPSIS
        Check and install required Az modules
    #>
    Write-Host "`n=== Checking Az Modules ===" -ForegroundColor Cyan

    $requiredModules = @("Az.Accounts", "Az.OperationalInsights")
    foreach ($mod in $requiredModules) {
        $module = Get-Module -ListAvailable -Name $mod
        if (-not $module) {
            Write-Host "Installing $mod..." -ForegroundColor Yellow
            Install-Module -Name $mod -Scope CurrentUser -Force -Confirm:$false -AllowClobber
            Write-Host "$mod installed successfully!" -ForegroundColor Green
        }
        Import-Module $mod -Force -WarningAction SilentlyContinue
        Write-Host "$mod Version: $((Get-Module $mod).Version)" -ForegroundColor Green
    }
}

function Initialize-AzAuth {
    <#
    .SYNOPSIS
        Initialize Azure authentication
    #>
    Write-Host "`n=== Azure Authentication ===" -ForegroundColor Cyan

    # Check if already logged in
    if (-not $ForceLogin) {
        try {
            $context = Get-AzContext -ErrorAction SilentlyContinue
            if ($context -and $context.Account) {
                # Check if tenant matches
                $currentTenant = $context.Tenant.Id
                if ($currentTenant -and $currentTenant -eq $TenantId) {
                    # Check if environment is Azure China Cloud
                    if ($context.Environment.Name -like "*China*") {
                        Write-Host "Current account: $($context.Account)"
                        Write-Host "Tenant ID: $currentTenant"
                        Write-Host "Environment: $($context.Environment.Name)"
                        Write-Host "Using existing session`n" -ForegroundColor Green
                        return
                    }
                }
            }
        }
        catch {
            Write-Verbose "No existing session found"
        }
    }

    # Ensure Azure China Cloud environment is registered
    $env = Get-AzEnvironment -Name $AzureEnvironment -ErrorAction SilentlyContinue
    if (-not $env) {
        Write-Host "Adding Azure China Cloud environment..." -ForegroundColor Yellow
        Add-AzEnvironment -Name $AzureEnvironment `
            -ActiveDirectoryEndpoint "https://login.chinacloudapi.cn/" `
            -ResourceManagerEndpoint "https://management.chinacloudapi.cn/" `
            -ServiceManagementUrl "https://management.core.chinacloudapi.cn/" `
            -GalleryEndpoint "https://gallery.chinacloudapi.cn/" `
            -ManagementPortalUrl "https://portal.azure.cn/" `
            | Out-Null
    }

    # Interactive login
    Write-Host "A browser window will open for login..." -ForegroundColor Yellow
    $connectParams = @{
        Environment = $AzureEnvironment
        Tenant      = $TenantId
    }
    if ($UseDeviceCode) {
        Write-Host "Using device code mode..." -ForegroundColor Yellow
        $connectParams['UseDeviceAuthentication'] = $true
    }

    Connect-AzAccount @connectParams | Out-Null

    $ctx = Get-AzContext
    Write-Host "Login successful! Account: $($ctx.Account)" -ForegroundColor Green
}

function Invoke-LogQuery {
    <#
    .SYNOPSIS
        Execute query using Az.OperationalInsights
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Query,

        [Parameter(Mandatory = $true)]
        [string]$WorkspaceId,

        [int]$Hours = 1,

        [datetime]$QueryStartTime,

        [datetime]$QueryEndTime,

        [switch]$IncludeStats
    )

    Write-Host "`n=== Executing Log Query ===" -ForegroundColor Cyan
    Write-Host "Workspace: $WorkspaceId"
    Write-Host "Query: $($Query.Substring(0, [Math]::Min(80, $Query.Length)))..."
    if ($QueryStartTime -and $QueryEndTime) {
        $startTime = $QueryStartTime
        $endTime = $QueryEndTime
        Write-Host "Time range: $($startTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    } else {
        $startTime = (Get-Date).AddHours(-$Hours)
        $endTime = (Get-Date)
        Write-Host "Time range: Past $Hours hours"
    }

    $queryMode = Get-LogQueryExecutionMode -QueryStartTime $QueryStartTime -QueryEndTime $QueryEndTime -Hours $Hours

    try {
        $queryParams = @{
            WorkspaceId = $WorkspaceId
            Query = $Query
            ErrorAction = 'Stop'
        }
        if ($queryMode.UseTimespan) {
            $queryParams['Timespan'] = $queryMode.Timespan
        }

        $response = Invoke-AzOperationalInsightsQuery @queryParams
    }
    catch {
        Write-Host "`nQuery failed!" -ForegroundColor Red

        if ($_.Exception.Message -match "401") {
            Write-Host "Status code: 401 (Authentication failed)" -ForegroundColor Red
            Write-Host "`n=== 401 Troubleshooting ===" -ForegroundColor Yellow
            Write-Host "1. Run: .\azure_log_query.ps1 -ForceLogin (re-login)"
            Write-Host "2. Confirm current user has read access to Workspace"
            Write-Host "   Role: Log Analytics Reader or higher"
        }
        elseif ($_.Exception.Message -match "403") {
            Write-Host "Status code: 403 (Insufficient permissions)" -ForegroundColor Red
            Write-Host "`n=== 403 Troubleshooting ===" -ForegroundColor Yellow
            Write-Host "1. Confirm user has permissions on:"
            Write-Host "   - Subscription: Reader at minimum"
            Write-Host "   - Resource group: Reader at minimum"
            Write-Host "   - Workspace: Log Analytics Reader"
        }
        else {
            Write-Host $_.Exception.Message -ForegroundColor Red
            if ($_.Exception.InnerException) {
                Write-Host "Internal error: $($_.Exception.InnerException.Message)" -ForegroundColor Red
            }
        }
        return
    }

    # Parse response
    if ($response.Results) {
        if ($ExportCsv) {
            Write-Host "`nExporting to CSV: $ExportCsv" -ForegroundColor Cyan
            $response.Results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
            Write-Host "Export successful! Total: $($response.Results.Count) rows" -ForegroundColor Green
        } else {
            $response.Results
        }
    }
    else {
        Write-Host "`nQuery returned empty results" -ForegroundColor Yellow
        if ($ExportCsv) {
            'TimeGenerated' | Out-File -FilePath $ExportCsv -Encoding UTF8 -Force
            Write-Host "Empty CSV created: $ExportCsv" -ForegroundColor Yellow
        }
    }
}

# ============================================================
# Main
# ============================================================
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "  Azure Monitor Log Query - China Cloud" -ForegroundColor Magenta
Write-Host "  Using Az.OperationalInsights Official Library" -ForegroundColor Magenta
Write-Host "============================================" -ForegroundColor Magenta

# Initialize modules
Initialize-Modules

# Authenticate
Initialize-AzAuth

# Build table query when table and explicit range are provided
$queryStartTime = $null
$queryEndTime = $null
if ($TableName) {
    if (-not $StartTime -or -not $EndTime) {
        throw '-TableName requires -StartTime and -EndTime.'
    }
    $queryStartTime = [DateTime]::Parse($StartTime)
    $queryEndTime = [DateTime]::Parse($EndTime)
    $Query = New-LogTableQuery -TableName $TableName -StartTime $queryStartTime -EndTime $queryEndTime
}

# Execute query
Invoke-LogQuery -Query $Query -WorkspaceId $WorkspaceId -Hours $Hours -QueryStartTime $queryStartTime -QueryEndTime $queryEndTime -IncludeStats:$IncludeStats


<#
# ============================================================
# Additional usage examples
# ============================================================

# Re-login (clear cache)
# .\azure_log_query.ps1 -ForceLogin

# Device code mode
# .\azure_log_query.ps1 -UseDeviceCode

# Custom query and time range
# .\azure_log_query.ps1 -Query "AzureActivity | top 20 by TimeGenerated desc" -Hours 48

# Specify Workspace and Tenant
# .\azure_log_query.ps1 -TenantId "your-tenant-id" -WorkspaceId "your-workspace-id"
#>
