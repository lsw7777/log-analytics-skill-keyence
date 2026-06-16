<#
.SYNOPSIS
    Azure Monitor Log Query Script - China Cloud Environment
    Query Log Analytics using official Az.OperationalInsights module
.DESCRIPTION
    Authenticate via Az module, execute KQL queries using official Az.OperationalInsights library
    Supports interactive browser login and device code flow
.EXAMPLE
    .\query-log-analytics.ps1
    .\query-log-analytics.ps1 -Query "AzureActivity | top 20 by TimeGenerated desc" -Hours 48
    .\query-log-analytics.ps1 -UseDeviceCode
    .\query-log-analytics.ps1 -Query "AuditGeneralDCR_CL | take 100" -ExportCsv ".\results.csv"
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
    [string]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [switch]$RepairAzModules,

    [Parameter(Mandatory = $false)]
    [switch]$RiskOnly,

    [Parameter(Mandatory = $false)]
    [switch]$RawCount,

    [Parameter(Mandatory = $false)]
    [switch]$NoProfile
)

# ============================================================
# Configuration
# ============================================================
$AzureEnvironment = "AzureChinaCloud"
$ModuleName = "Az.OperationalInsights"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $ScriptDir 'log-analyzer-shared.ps1')

# ============================================================
# Functions
# ============================================================

function Initialize-Modules {
    <#
    .SYNOPSIS
        Check and install required Az modules
    #>
    param([switch]$Silent)
    
    if (-not $Silent) {
        Write-Host "`n=== Checking Az Modules ===" -ForegroundColor Cyan
    }

    if ($RepairAzModules) {
        Repair-AzModules
    }

    Get-Module AzureRM* -ErrorAction SilentlyContinue | Remove-Module -Force -ErrorAction SilentlyContinue

    $requiredModules = @("Az.Accounts", "Az.OperationalInsights")
    foreach ($mod in $requiredModules) {
        $module = Get-Module -ListAvailable -Name $mod
        if (-not $module) {
            if (-not $Silent) {
                Write-Host "Installing $mod..." -ForegroundColor Yellow
            }
            Install-Module -Name $mod -Scope CurrentUser -Force -Confirm:$false -AllowClobber
            if (-not $Silent) {
                Write-Host "$mod installed successfully!" -ForegroundColor Green
            }
        }
        Import-Module $mod -Force -WarningAction SilentlyContinue
        if (-not $Silent) {
            Write-Host "$mod Version: $((Get-Module $mod).Version)" -ForegroundColor Green
        }
    }
}

function Repair-AzModules {
    Write-Host "`n=== Repairing Az Modules ===" -ForegroundColor Cyan
    Write-Host "Installing/updating Az.Accounts and Az.OperationalInsights for CurrentUser..." -ForegroundColor Yellow
    Install-Module -Name Az.Accounts -Scope CurrentUser -Force -AllowClobber -Confirm:$false
    Install-Module -Name Az.OperationalInsights -Scope CurrentUser -Force -AllowClobber -Confirm:$false
    Write-Host "Az module repair completed. Restart PowerShell before running the report again." -ForegroundColor Green
}

function Write-AzModuleRepairHelp {
    Write-Host "`n=== Az Module Repair Required ===" -ForegroundColor Yellow
    Write-Host "Connect-AzAccount failed before authentication completed. This usually means incompatible Az module assemblies are installed or already loaded in the current PowerShell session." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Recommended fix:" -ForegroundColor Cyan
    Write-Host "1. Close all PowerShell windows."
    Write-Host "2. Open a new PowerShell window."
    Write-Host "3. Run:"
    Write-Host "   Set-PSRepository -Name PSGallery -InstallationPolicy Trusted"
    Write-Host "   Install-Module Az.Accounts -Scope CurrentUser -Force -AllowClobber"
    Write-Host "   Install-Module Az.OperationalInsights -Scope CurrentUser -Force -AllowClobber"
    Write-Host "4. Re-open PowerShell again, then run .\scripts\main.ps1 from the skill root"
    Write-Host ""
    Write-Host "Project shortcut:"
    Write-Host "   .\scripts\query-log-analytics.ps1 -RepairAzModules"
    Write-Host ""
}

function Initialize-AzAuth {
    <#
    .SYNOPSIS
        Initialize Azure authentication
    #>
    param([switch]$Silent)
    
    if (-not $Silent) {
        Write-Host "`n=== Azure Authentication ===" -ForegroundColor Cyan
    }

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
                        if (-not $Silent) {
                            Write-Host "Current account: $($context.Account)"
                            Write-Host "Tenant ID: $currentTenant"
                            Write-Host "Environment: $($context.Environment.Name)"
                            Write-Host "Using existing session`n" -ForegroundColor Green
                        }
                        return
                    }
                }
            }
        }
        catch {
            Write-Verbose "No existing session found"
        }
    }

    # Connect using built-in AzureChinaCloud environment name
    if (-not $Silent) {
        Write-Host "A browser window will open for login..." -ForegroundColor Yellow
    }
    $connectParams = @{
        Environment = $AzureEnvironment
        Tenant      = $TenantId
    }
    if ($UseDeviceCode) {
        if (-not $Silent) {
            Write-Host "Using device code mode..." -ForegroundColor Yellow
        }
        $connectParams['UseDeviceAuthentication'] = $true
    }

    try {
        Connect-AzAccount @connectParams | Out-Null
    }
    catch {
        if ($_.Exception -is [System.TypeLoadException] -or $_.Exception.Message -match 'SerializationSettings|TypeLoadException|ResourceManagementClient') {
            Write-AzModuleRepairHelp
            exit 20
        }
        throw
    }

    $ctx = Get-AzContext
    if (-not $Silent) {
        Write-Host "Login successful! Account: $($ctx.Account)" -ForegroundColor Green
    }
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

        [switch]$IncludeStats,

        [switch]$RawCount
    )

    if (-not $RawCount) {
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
    }

    $queryMode = Get-LogQueryExecutionMode -QueryStartTime $QueryStartTime -QueryEndTime $QueryEndTime -Hours $Hours

    try {
        $queryParams = @{
            WorkspaceId = $WorkspaceId
            Query = $Query
            ErrorAction = 'Stop'
        }

        # For RawCount queries, the KQL query already contains time filtering in the where clause.
        # Do NOT pass Timespan parameter, as it may override or conflict with the query's time filter.
        # For non-RawCount queries, use Timespan parameter as usual.
        if (-not $RawCount) {
            if ($queryMode.UseTimespan) {
                $queryParams['Timespan'] = $queryMode.Timespan
            } else {
                # Calculate timespan from start and end times
                $queryParams['Timespan'] = New-TimeSpan -Start $queryMode.StartTime -End $queryMode.EndTime
            }
        }

        $response = Invoke-AzOperationalInsightsQuery @queryParams
    }
    catch {
        if (-not $RawCount) {
            Write-Host "`nQuery failed!" -ForegroundColor Red

            # 显示查询摘要信息（不打印完整查询以避免暴露IP地址）
            Write-Host "`n=== Query Debug Info ===" -ForegroundColor Yellow
            Write-Host "Table: $TableName" -ForegroundColor Yellow
            Write-Host "Query length: $($Query.Length) characters" -ForegroundColor Yellow
            # Do NOT print the full query to avoid exposing IP addresses

            if ($_.Exception.Message -match "401") {
                Write-Host "Status code: 401 (Authentication failed)" -ForegroundColor Red
                Write-Host "`n=== 401 Troubleshooting ===" -ForegroundColor Yellow
                Write-Host "1. Run: .\scripts\query-log-analytics.ps1 -ForceLogin (re-login)"
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
            elseif ($_.Exception.Message -match "BadRequest|400") {
                Write-Host "Status code: 400 (Bad Request - Query syntax error)" -ForegroundColor Red
                Write-Host "`n=== 400 Troubleshooting ===" -ForegroundColor Yellow
                Write-Host "1. Check if the table exists in the workspace"
                Write-Host "2. Verify column names in the query match the table schema"
                Write-Host "3. Try running a simple query first: $TableName | take 10"
                Write-Host "4. Check for KQL syntax errors in the query above"
            }
            else {
                Write-Host $_.Exception.Message -ForegroundColor Red
                if ($_.Exception.InnerException) {
                    Write-Host "Internal error: $($_.Exception.InnerException.Message)" -ForegroundColor Red
                }
            }
        }
        throw
    }

    # Parse response
    # Handle both DataTable and array/null cases for Results
    $hasResults = $false
    if ($null -ne $response.Results) {
        if ($response.Results -is [System.Data.DataTable]) {
            $hasResults = $response.Results.Rows.Count -gt 0
        } elseif ($response.Results -is [array]) {
            $hasResults = $response.Results.Count -gt 0
        } else {
            # Single object or other type
            $hasResults = $true
        }
    }
    
    if ($hasResults) {
        if ($RawCount) {
            # For count queries, output just the number
            # KQL count operator returns field as 'count_' (with underscore)
            $countValue = $null
            
            # Check if Results is a DataTable (Invoke-AzOperationalInsightsQuery returns DataTable)
            if ($response.Results -is [System.Data.DataTable]) {
                $dataTable = $response.Results
                Write-Verbose "DataTable columns: $($dataTable.Columns | ForEach-Object { $_.ColumnName })"
                Write-Verbose "DataTable rows count: $($dataTable.Rows.Count)"
                
                if ($dataTable.Rows.Count -gt 0) {
                    $row = $dataTable.Rows[0]
                    # Try multiple possible column names (KQL count returns 'count_')
                    foreach ($colName in @('count_', 'count', 'Count', 'COUNT')) {
                        if ($dataTable.Columns.Contains($colName)) {
                            $countValue = $row[$colName]
                            Write-Verbose "Found count in column '$colName': $countValue"
                            break
                        }
                    }
                    # If still null, try the first column
                    if ($null -eq $countValue -and $dataTable.Columns.Count -gt 0) {
                        $countValue = $row[0]
                        Write-Verbose "Using first column value: $countValue"
                    }
                }
            } else {
                # Results is an array or single object (fallback)
                $resultArray = if ($response.Results -is [array]) { $response.Results } else { @($response.Results) }
                $result = $resultArray[0]
                $countValue = $null
                
                # Debug: Output all property names to help identify the correct field
                $debugFieldNames = @($result.PSObject.Properties.Name) -join ', '
                Write-Verbose "Available fields: $debugFieldNames"
                
                # Try multiple possible field names
                foreach ($fieldName in @('count_', 'count', 'Count', 'COUNT')) {
                    if ($result.PSObject.Properties.Name -contains $fieldName) {
                        $countValue = $result.$fieldName
                        Write-Verbose "Found count in field '$fieldName': $countValue"
                        break
                    }
                }
                
                # If still null, try to get the first property value
                if ($null -eq $countValue) {
                    $firstProp = $result.PSObject.Properties | Select-Object -First 1
                    if ($firstProp) {
                        $countValue = $firstProp.Value
                        Write-Verbose "Using first property value: $countValue"
                    }
                }
            }
            
            if ($null -eq $countValue) { $countValue = 0 }
            # Output as string to ensure consistent parsing
            Write-Output "$countValue"
        } elseif ($ExportCsv) {
            Write-Host "`nExporting to CSV: $ExportCsv" -ForegroundColor Cyan
            $response.Results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
            Write-Host "Export successful! Total: $($response.Results.Count) rows" -ForegroundColor Green
        } else {
            $response.Results
        }
    }
    else {
        if ($RawCount) {
            # Output as string to ensure consistent parsing
            Write-Output "0"
        } else {
            Write-Host "`nQuery returned empty results" -ForegroundColor Yellow
            if ($ExportCsv) {
                'TimeGenerated' | Out-File -FilePath $ExportCsv -Encoding UTF8 -Force
                Write-Host "Empty CSV created: $ExportCsv" -ForegroundColor Yellow
            }
        }
    }
}

# ============================================================
# Main
# ============================================================
if (-not $RawCount) {
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host "  Azure Monitor Log Query - China Cloud" -ForegroundColor Magenta
    Write-Host "  Using Az.OperationalInsights Official Library" -ForegroundColor Magenta
    Write-Host "============================================" -ForegroundColor Magenta
}

# Initialize modules
Initialize-Modules -Silent:$RawCount

if ($RepairAzModules) {
    exit 0
}

# Authenticate
Initialize-AzAuth -Silent:$RawCount

# Build table query when table and explicit range are provided
$queryStartTime = $null
$queryEndTime = $null
if ($TableName) {
    if (-not $StartTime -or -not $EndTime) {
        throw '-TableName requires -StartTime and -EndTime.'
    }
    $queryStartTime = [DateTime]::Parse($StartTime)
    $queryEndTime = [DateTime]::Parse($EndTime)
    $Query = New-LogTableQuery -TableName $TableName -StartTime $queryStartTime -EndTime $queryEndTime -RiskOnly:$RiskOnly
} elseif ($StartTime -and $EndTime) {
    # When no TableName but StartTime/EndTime are provided (e.g., for count queries)
    $queryStartTime = [DateTime]::Parse($StartTime)
    $queryEndTime = [DateTime]::Parse($EndTime)
}

# Execute query
Invoke-LogQuery -Query $Query -WorkspaceId $WorkspaceId -Hours $Hours -QueryStartTime $queryStartTime -QueryEndTime $queryEndTime -IncludeStats:$IncludeStats -RawCount:$RawCount


<#
# ============================================================
# Additional usage examples
# ============================================================

# Re-login (clear cache)
# .\query-log-analytics.ps1 -ForceLogin

# Device code mode
# .\query-log-analytics.ps1 -UseDeviceCode

# Custom query and time range
# .\query-log-analytics.ps1 -Query "AzureActivity | top 20 by TimeGenerated desc" -Hours 48

# Specify Workspace and Tenant
# .\query-log-analytics.ps1 -TenantId "your-tenant-id" -WorkspaceId "your-workspace-id"
#>