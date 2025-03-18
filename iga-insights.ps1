<#
.SYNOPSIS
    Analyzes Azure/Entra ID SSO application usage and generates detailed HTML reports.

.DESCRIPTION
    This script identifies inactive users of specified Entra ID SSO applications by analyzing sign-in logs in Log Analytics Workspace.
    It generates comprehensive HTML and CSV reports showing user activity, group memberships, and job titles.

.PARAMETER AppName
    The display name of the SSO application (Service Principle) to analyze.

.PARAMETER ThresholdDays
    Number of days to consider for determining user inactivity (e.g., 90).

.EXAMPLE
    .\iga-insights.ps1 -AppName "Salesforce" -ThresholdDays 90

.NOTES
    Version:        
    Author:         Marcel Nguyen
    Creation Date:  
    Requirements:   
        -PowerShell 7.0 or higher
        -Az PowerShell module (Az.Accounts, Az.OperationalInsights)
        -Microsoft.Graph PowerShell module
        - A Log Analytics Workspace with at least 365 days of retention period (recommended)
        - SignInLogs data collection enabled in the workspace
        Azure/Entra ID account with:
        - Log Analytics Workspace access
        - Log Analytics Reader permissions
        - Azure/Entra ID Reader permissions

.LINK
    https://github.com/marcel-ngn/iga-insights
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the SSO application display name")]
    [string]$AppName,

    [Parameter(Mandatory = $true, HelpMessage = "Enter number of days for inactivity threshold")]
    [int]$ThresholdDays
)

Write-Host -ForegroundColor blue "
.___  ________    _____     .___ _______    _________.___  ________  ___ _______________________
|   |/  _____/   /  _  \    |   |\      \  /   _____/|   |/  _____/ /   |   \__    ___/   _____/
|   /   \  ___  /  /_\  \   |   |/   |   \ \_____  \ |   /   \  ___/    ~    \|    |  \_____  \ 
|   \    \_\  \/    |    \  |   /    |    \/        \|   \    \_\  \    Y    /|    |  /        \
|___|\______  /\____|__  /  |___\____|__  /_______  /|___|\______  /\___|_  / |____| /_______  /
            \/         \/               \/        \/             \/       \/                 \/
                by Marcel Nguyen
                       
For usage information see the documentation here: https://github.com/marcel-ngn/iga-insights"

#Region Functions
function Test-ModuleAvailability {
    param (
        [string[]]$RequiredModules
    )
    
    foreach ($module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Error "Required module '$module' is not installed. Please install it using: Install-Module $module -Scope CurrentUser"
            return $false
        }
    }
    return $true
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        'Info' { 'White' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
    }
    
    Write-Host "[$timestamp] $Level - $Message" -ForegroundColor $color
}

function Initialize-UserConnection {
    try {
        Write-Log "Connecting to Azure and Microsoft Graph services..." -Level Info
        
        # Clear existing context and connect to Azure
        Clear-AzContext -Force
        Connect-AzAccount -ErrorAction Stop | Out-Null

        Write-Log "Connected to Azure, continuing with Microsoft Graph"
        
        # Connect to Microsoft Graph
        Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -ErrorAction Stop | Out-Null

        Write-Log "✅ Successfully connected to both services" -Level Info
        return $true
    }
    catch {
        Write-Log "Connection failed: $_" -Level Error
        return $false
    }
}



function Initialize-AppRegistrationConnection {

    $TenantId = ""
    $ClientId = ""
    $CertThumbprint = ""

    try {
        Write-Log "Connecting to Azure and Microsoft Graph services..." -Level Info
        
        # Connect to Azure Account using certificate
        Write-Log "Checking Azure connection..." -Level Info
        $azContext = Get-AzContext -ErrorAction Stop
        if (-not $azContext) {
            Write-Log "Azure connection not found. Connecting..." -Level Info
            Connect-AzAccount -ServicePrincipal -Tenant $TenantId -ApplicationId $ClientId -CertificateThumbprint $CertThumbprint -ErrorAction Stop | Out-Null
            $azContext = Get-AzContext -ErrorAction Stop
            if (-not $azContext) {
                throw "Failed to establish Azure connection"
            }
        }
        Write-Log "Connected to Azure as service principal $ClientId" -Level Info

        # Connect to Microsoft Graph using certificate
        Write-Log "Checking Microsoft Graph connection..." -Level Info
        try {
            $graphContext = Get-MgContext -ErrorAction Stop
            if ($null -eq $graphContext) {
                Write-Log "Microsoft Graph connection not found. Connecting..." -Level Info
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertThumbprint -ErrorAction Stop | Out-Null
                
                $graphContext = Get-MgContext -ErrorAction Stop
                if ($null -eq $graphContext) {
                    throw "Failed to establish Microsoft Graph connection"
                }
            }
            Write-Log "Connected to Microsoft Graph as application $ClientId" -Level Info
        }
        catch {
            throw "Microsoft Graph connection failed: $_"
        }

        Write-Log "Successfully connected to both services" -Level Info
        return $true
    }
    catch {
        Write-Log "Connection failed: $_" -Level Error
        return $false
    }
}



function Get-WorkspaceDetails {
    try {
        $workspaceName = Read-Host "Enter the Log Analytics Workspace Name"
        $workspaceRG = Read-Host "Enter the Workspace Resource Group"
        
        $workspace = Get-AzOperationalInsightsWorkspace -Name $workspaceName -ResourceGroupName $workspaceRG -ErrorAction Stop
        return $workspace.CustomerId.Guid
    }
    catch {
        Write-Log "Error getting workspace details: $_" -Level Error
        throw
    }
}

function Get-EnterpriseAppGroups {
    param (
        [string]$ApplicationName
    )
    
    try {
        Write-Log "⚙️ Fetching Enterprise Application details for '$ApplicationName'..." -Level Info
        
        # Get the service principal for the application
        $servicePrincipal = Get-MgServicePrincipal -Filter "DisplayName eq '$ApplicationName'"
        
        if (-not $servicePrincipal) {
            Write-Log "Enterprise Application '$ApplicationName' not found." -Level Error
            return $null
        }

        if ($servicePrincipal.Count -gt 1) {
            Write-Log "Multiple applications found with name '$ApplicationName'. Using the first one." -Level Warning
            $servicePrincipal = $servicePrincipal[0]
        }

        Write-Log "🔎 Found Enterprise Application with ID: $($servicePrincipal.Id)" -Level Info

        # Get app role assignments (group assignments)
        $assignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id -All
        
        # Filter for group assignments
        $groupAssignments = $assignments | Where-Object { $_.PrincipalType -eq "Group" }
        
        if (-not $groupAssignments) {
            Write-Log "No groups found assigned to the application." -Level Warning
            return $null
        }

        # Get group details for each assignment
        $groups = @()
        foreach ($assignment in $groupAssignments) {
            try {
                $group = Get-MgGroup -GroupId $assignment.PrincipalId
                $groups += $group
                Write-Log "🔎 Found assigned group: $($group.DisplayName)" -Level Info
            }
            catch {
                Write-Log "Error retrieving group details for ID $($assignment.PrincipalId): $_" -Level Warning
                continue
            }
        }

        return $groups
    }
    catch {
        Write-Log "Error retrieving Enterprise Application groups: $_" -Level Error
        return $null
    }
}

function Disconnect-AllServices {
    Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Clear-AzContext -Force -ErrorAction SilentlyContinue | Out-Null
    
 }

#EndRegion Functions

#Region Script Initialization
# Verify required modules
$requiredModules = @('Az.Accounts', 'Az.OperationalInsights', 'Microsoft.Graph')
if (-not (Test-ModuleAvailability -RequiredModules $requiredModules)) {
    exit 1
}

# Prompt for authentication method
$authChoice = Read-Host "Select authentication method:
1. User Credentials
2. App Registration
Enter choice (1 or 2)"

switch ($authChoice) {
    "1" {
        if (-not (Initialize-UserConnection)) {
            exit 1
        }
    }
    "2" {
        if (-not (Initialize-AppRegistrationConnection -TenantId $TenantId -ClientId $ClientId -CertThumbprint $CertThumbprint)) {
            exit 1
        }
    }
    default {
        Write-Log "Invalid choice. Please select 1 or 2." -Level Error
        exit 1
    }
}

# Setup report paths
$today = Get-Date -Format "yyyy-MM-dd"
$defaultPath = [Environment]::GetFolderPath("Desktop")
$ReportPath = Read-Host "Enter the report destination path (press Enter for Desktop)"
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = $defaultPath
}

$csvFilePath = Join-Path $ReportPath "$($AppName)_users_${ThresholdDays}days_$today.csv"
$htmlFilePath = Join-Path $ReportPath "$($AppName)_Users_Report_$today.html"

# Get Workspace ID
try {
    $WorkSpaceId = Get-WorkspaceDetails
}
catch {
    Write-Log "Failed to get workspace details. Exiting script." -Level Error
    exit 1
}
#EndRegion Script Initialization

#Region Data Collection
# Define KQL queries
$activeUsersQuery = @"
SigninLogs
| where TimeGenerated > ago(${ThresholdDays}d)
| where AppDisplayName =~ '$($AppName)'
| summarize LatestSignIn = arg_max(TimeGenerated, *) by UserPrincipalName
| project UserPrincipalName, LatestSignIn = format_datetime(LatestSignIn, 'yyyy-MM-dd HH:mm:ss')
"@

$historicalQuery = @"
SigninLogs
| where TimeGenerated > ago(365d)
| where AppDisplayName =~ '$($AppName)'
| summarize LatestSignIn = arg_max(TimeGenerated, *) by UserPrincipalName
| project UserPrincipalName, LatestSignIn = format_datetime(LatestSignIn, 'yyyy-MM-dd HH:mm:ss')
"@

try {
    Write-Log "⚙️ Fetching Log Analytics Workspace Data..."
    $activeUsersResults = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkSpaceId -Query $activeUsersQuery
    $historicalResults = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkSpaceId -Query $historicalQuery
    
    # Convert results to hashtables for faster lookup
    $activeUsers = @{}
    $SignInDates365d = @{}
    $AppDisplayNames = @{}
    
    foreach ($result in $activeUsersResults.Results) {
        $activeUsers[$result.UserPrincipalName] = $result.LatestSignIn
    }
    
    foreach ($result in $historicalResults.Results) {
        $SignInDates365d[$result.UserPrincipalName] = $result.LatestSignIn
        $AppDisplayNames[$result.UserPrincipalName] = $result.AppDisplayName
    }
}
catch {
    Write-Log "Error executing KQL queries: $_" -Level Error
    exit 1
}

# Get relevant groups and users
try {
    $app_groups = Get-EnterpriseAppGroups -ApplicationName $AppName
    if (-not $app_groups) {
        Write-Log "❌ No assigned groups found for application '$AppName'" -Level Error
        exit 1
    }
    
    $app_users = @()
    $userGroupMemberships = @{}
    $userJobTitles = @{}
    $userDepartments = @{}
    
    foreach ($group in $app_groups) {
        Write-Log "⚙️ Processing group: $($group.DisplayName)" -Level Info
        try {
            $group_members = Get-MgGroupMember -GroupId $group.Id -All
            
            foreach ($member in $group_members) {
                try {
                    $userDetails = Get-MgUser -UserId $member.Id -Property @('UserPrincipalName','JobTitle','Department') -ErrorAction SilentlyContinue | 
                        Select-Object UserPrincipalName, JobTitle, Department                    
                    if ($userDetails) {
                        $userPrincipalName = $userDetails.UserPrincipalName
                        
                        if (-not $userGroupMemberships.ContainsKey($userPrincipalName)) {
                            $userGroupMemberships[$userPrincipalName] = @()
                        }
                        $userGroupMemberships[$userPrincipalName] += $group.DisplayName
                        
                        $userJobTitles[$userPrincipalName] = if ([string]::IsNullOrWhiteSpace($userDetails.JobTitle)) { 
                            "n/a" 
                        } else { 
                            $userDetails.JobTitle 
                        }
                        
                        $userDepartments[$userPrincipalName] = if ([string]::IsNullOrWhiteSpace($userDetails.Department)) { 
                            "n/a" 
                        } else { 
                            $userDetails.Department 
                        }
                        
                        if (-not ($app_users | Where-Object { $_.UserPrincipalName -eq $userPrincipalName })) {
                            $app_users += [PSCustomObject]@{
                                UserPrincipalName = $userPrincipalName
                            }
                        }
                    }
                    else {
                        Write-Log "Skipping member $($member.Id) in group $($group.DisplayName) - Unable to get user details" -Level Warning
                    }
                }
                catch {
                    Write-Log "❌ Error processing member $($member.Id) in group $($group.DisplayName): $_" -Level Warning
                    continue
                }
            }
        }
        catch {
            Write-Log "❌ Error getting members for group $($group.DisplayName): $_" -Level Warning
            continue
        }
    }
    
    Write-Log "🔎 Found $($app_users.Count) users across $($app_groups.Count) groups" -Level Info
}
catch {
    Write-Log "Critical error processing groups and users: $_" -Level Error
    exit 1
}

# Process user activity
$unique_app_users = $app_users | Sort-Object -Unique -Property UserPrincipalName
$activeAppUsers = @()
$inactiveAppUsers = @()

foreach ($user in $unique_app_users) {
    $userGroups = $userGroupMemberships[$user.UserPrincipalName] -join "; "
    $jobTitle = $userJobTitles[$user.UserPrincipalName]
    $Department = $userDepartments[$user.UserPrincipalName]
    
    if ($activeUsers.ContainsKey($user.UserPrincipalName)) {
        $lastSignIn = $activeUsers[$user.UserPrincipalName]
        $AppDisplayName = $AppDisplayNames[$user.UserPrincipalName]
        
        $activeAppUsers += [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            LatestSignIn = $lastSignIn
            AppDisplayName = $AppDisplayName
            Groups = $userGroups
            JobTitle = $jobTitle
            Department = $Department
        }
    }
    else {
        $lastSignIn = if ($SignInDates365d.ContainsKey($user.UserPrincipalName)) {
            $SignInDates365d[$user.UserPrincipalName]
        }
        else {
            "No sign-in record in the last 365 days"
        }
        
        $AppDisplayName = $AppDisplayNames[$user.UserPrincipalName]
        
        $inactiveAppUsers += [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            LatestSignIn = $lastSignIn
            AppDisplayName = $AppDisplayName
            Groups = $userGroups
            JobTitle = $jobTitle = $userJobTitles[$user.UserPrincipalName]
            Department = $Department = $userDepartments[$user.UserPrincipalName]
        }
    }
}

#EndRegion Data Collection

#Region Report Generation
# Export to CSV
try {
    # Combine active and inactive users
    $allUsers = $activeAppUsers + $inactiveAppUsers | Sort-Object UserPrincipalName
    $allUsers | Export-Csv -Path $csvFilePath -NoTypeInformation
    Write-Log "✅ CSV report exported to: $csvFilePath"
}
catch {
    Write-Log "Error exporting CSV report: $_" -Level Error
}

# Generate HTML Report
$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Identity Governance Insights: $($AppName)</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
    <style>
        :root {
            --primary-color: #2563eb;
            --secondary-color: #1e40af;
            --danger-color: #dc2626;
            --success-color: #16a34a;
            --background-color: #f8fafc;
            --card-background: #ffffff;
            --text-primary: #1e293b;
            --text-secondary: #64748b;
            --border-color: #e2e8f0;
            --hover-color: #f1f5f9;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: system-ui, -apple-system, sans-serif;
        }

        body {
            background-color: var(--background-color);
            color: var(--text-primary);
            line-height: 1.5;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 2rem;
        }

        .header {
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
            color: white;
            padding: 2rem;
            border-radius: 1rem;
            margin-bottom: 2rem;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }

        .header h1 {
            font-size: 1.875rem;
            font-weight: 700;
            margin-bottom: 0.5rem;
        }

        .header p {
            opacity: 0.9;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2rem;
        }

        .stat-card {
            background: var(--card-background);
            padding: 1.5rem;
            border-radius: 1rem;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
            transition: transform 0.2s;
        }

        .stat-card:hover {
            transform: translateY(-2px);
        }

        .stat-card h3 {
            color: var(--text-secondary);
            font-size: 0.875rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 0.5rem;
        }

        .stat-card p {
            font-size: 2rem;
            font-weight: 700;
            margin-bottom: 0.5rem;
        }

        .stat-card .percentage {
            font-size: 0.875rem;
            color: var(--text-secondary);
        }

        .charts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2rem;
        }

        .chart-card {
            background: var(--card-background);
            padding: 1.5rem;
            border-radius: 1rem;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        }

        .chart-title {
            font-size: 1.25rem;
            font-weight: 600;
            margin-bottom: 1.5rem;
            color: var(--text-primary);
        }

        .filters-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1rem;
            margin-bottom: 1.5rem;
        }

        .filter-group {
            display: flex;
            flex-direction: column;
        }

        .filter-label {
            font-size: 0.875rem;
            color: var(--text-secondary);
            margin-bottom: 0.5rem;
        }

        .filter-select {
            padding: 0.5rem;
            border: 1px solid var(--border-color);
            border-radius: 0.5rem;
            background-color: var(--card-background);
            color: var(--text-primary);
            font-size: 0.875rem;
            transition: border-color 0.2s;
        }

        .filter-select:hover {
            border-color: var(--primary-color);
        }

        .table-container {
            background: var(--card-background);
            border-radius: 1rem;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
            overflow: hidden;
            margin-bottom: 2rem;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        th {
            background-color: var(--background-color);
            padding: 1rem;
            text-align: left;
            font-weight: 600;
            color: var(--text-primary);
            font-size: 0.875rem;
            border-bottom: 2px solid var(--border-color);
        }

        td {
            padding: 1rem;
            border-bottom: 1px solid var(--border-color);
            color: var(--text-primary);
            font-size: 0.875rem;
        }

        tr:hover {
            background-color: var(--hover-color);
        }

        .status-badge {
            display: inline-block;
            padding: 0.25rem 0.75rem;
            border-radius: 9999px;
            font-size: 0.75rem;
            font-weight: 500;
        }

        .status-active {
            background-color: #dcfce7;
            color: var(--success-color);
        }

        .status-inactive {
            background-color: #fee2e2;
            color: var(--danger-color);
        }

        .footer {
            text-align: center;
            padding: 2rem;
            color: var(--text-secondary);
            font-size: 0.875rem;
            border-top: 1px solid var(--border-color);
        }

        .pagination-controls {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1rem;
            padding: 1rem;
            background-color: var(--card-background);
            border-radius: 0.5rem;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        }

        .entries-per-page {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .pagination-info {
            color: var(--text-secondary);
            font-size: 0.875rem;
        }

        .pagination-buttons {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 0.5rem;
            margin: 1.5rem 0;
        }

        .pagination-button {
            padding: 0.5rem 1rem;
            border: 1px solid var(--border-color);
            border-radius: 0.5rem;
            background-color: var(--card-background);
            color: var(--text-primary);
            cursor: pointer;
            transition: all 0.2s;
        }

        .pagination-button:hover:not(:disabled) {
            background-color: var(--primary-color);
            color: white;
            border-color: var(--primary-color);
        }

        .pagination-button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .page-numbers {
            display: flex;
            gap: 0.25rem;
        }

        .page-number {
            padding: 0.5rem 1rem;
            border: 1px solid var(--border-color);
            border-radius: 0.5rem;
            background-color: var(--card-background);
            color: var(--text-primary);
            cursor: pointer;
            transition: all 0.2s;
        }

        .page-number:hover {
            background-color: var(--hover-color);
        }

        .page-number.active {
            background-color: var(--primary-color);
            color: white;
            border-color: var(--primary-color);
        }
        .footer {
            text-align: center;
            padding: 2rem;
            color: var(--text-secondary);
            font-size: 0.875rem;
            border-top: 1px solid var(--border-color);
        }

        #floating-github {
            position: fixed;
            right: 20px;
            bottom: 20px;
            background: #ffffff;
            border: 1px solid #e5e7eb;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            border-radius: 6px;
            padding: 6px 12px;
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 12px;
            transition: all 0.2s ease;
            text-decoration: none;
            color: #000000;
            z-index: 9999;
            opacity: 0.7;
        }

        #floating-github:hover {
            opacity: 1;
            transform: translateY(-2px);
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }

        #floating-github svg {
            width: 14px;
            height: 14px;
        }


        @media (max-width: 768px) {
            .container {
                padding: 1rem;
            }

            .header {
                padding: 1.5rem;
            }

            .stat-card p {
                font-size: 1.5rem;
            }
        }
    
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Identity Governance Insights: $($AppName)</h1>
            <p>Analysis Period: Past $ThresholdDays Days</p>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card">
                <h3>Total Users in Groups</h3>
                <p>$($unique_app_users.Count)</p>
            </div>
            <div class="stat-card">
                <h3>Active Users</h3>
                <p style="color: var(--success-color)">$($activeAppUsers.Count)</p>
                <div class="percentage">
                    $(if ($unique_app_users.Count -gt 0) { 
                        [math]::Round(($activeAppUsers.Count / $unique_app_users.Count) * 100, 2)
                    } else { 0 })%
                </div>
            </div>
            <div class="stat-card">
                <h3>Inactive Users</h3>
                <p style="color: var(--danger-color)">$($inactiveAppUsers.Count)</p>
                <div class="percentage">
                    $(if ($unique_app_users.Count -gt 0) { 
                        [math]::Round(($inactiveAppUsers.Count / $unique_app_users.Count) * 100, 2)
                    } else { 0 })%
                </div>
            </div>
        </div>

        <div class="charts-grid">
            <div class="chart-card">
                <h2 class="chart-title">Activity Distribution</h2>
                <canvas id="userActivityChart"></canvas>
            </div>
            <div class="chart-card">
                <h2 class="chart-title">Job Title Distribution</h2>
                <canvas id="jobTitlesChart"></canvas>
            </div>
            <div class="chart-card">
                <h2 class="chart-title">Department Distribution</h2>
                <canvas id="departmentsChart"></canvas>
            </div>
        </div>

        <h2 class="chart-title">User Details</h2>
        
        <div class="filters-grid">
            <div class="filter-group">
                <label class="filter-label">User Principal Name</label>
                <select id="filter-upn" class="filter-select"></select>
            </div>
            <div class="filter-group">
                <label class="filter-label">Job Title</label>
                <select id="filter-job-title" class="filter-select"></select>
            </div>
            <div class="filter-group">
                <label class="filter-label">Department</label>
                <select id="filter-department" class="filter-select"></select>
            </div>
            <div class="filter-group">
                <label class="filter-label">Last Sign-In</label>
                <select id="filter-last-signin" class="filter-select"></select>
            </div>
            <div class="filter-group">
                <label class="filter-label">Status</label>
                <select id="filter-status" class="filter-select"></select>
            </div>
            <div class="filter-group">
                <label class="filter-label">Group Memberships</label>
                <select id="filter-groups" class="filter-select"></select>
            </div>
        </div>
        
        <div class="table-container">
            <table id="users-table">
                <thead>
                    <tr>
                        <th>UserPrincipalName</th>
                        <th>Job Title</th>
                        <th>Department</th>
                        <th>Last Sign-In</th>
                        <th>Status</th>
                        <th>Group Memberships</th>
                    </tr>
                </thead>
                <tbody>
                $(
                    $allUsers = $activeAppUsers + $inactiveAppUsers | Sort-Object UserPrincipalName
                    foreach ($user in $allUsers) {
                        $status = if ($activeAppUsers.UserPrincipalName -contains $user.UserPrincipalName) { 
                            '<span class="status-badge status-active">Active</span>' 
                        } else { 
                            '<span class="status-badge status-inactive">Inactive</span>' 
                        }
                        "<tr>
                            <td>$($user.UserPrincipalName)</td>
                            <td>$($user.JobTitle)</td>
                            <td>$($user.Department)</td>
                            <td>$($user.LatestSignIn)</td>
                            <td>$status</td>
                            <td>$($user.Groups)</td>
                        </tr>"
                    }
                )
                </tbody>
            </table>
        </div>

        <div class="pagination-controls">
            <div class="entries-per-page">
                <label class="filter-label">Show entries:</label>
                <select id="entries-per-page" class="filter-select">
                    <option value="10">10</option>
                    <option value="20" selected>20</option>
                    <option value="50">50</option>
                    <option value="100">100</option>
                </select>
            </div>
            <div class="pagination-info">
                Showing <span id="showing-start">0</span> to <span id="showing-end">0</span> of <span id="total-entries">0</span> entries
            </div>
        </div>

        <div class="pagination-buttons">
            <button id="prev-page" class="pagination-button" disabled>Previous</button>
            <div id="page-numbers" class="page-numbers"></div>
            <button id="next-page" class="pagination-button" disabled>Next</button>
        </div>

        <div class="footer">
            <p>Report generated on $today</p>
            <p>Analysis threshold: $ThresholdDays days</p>
            <p>CSV Export: $csvFilePath</p>
        </div>
    </div>

    <script>
        (function() {
            try {
                // Wait for Chart.js to load
                if (typeof Chart === 'undefined') {
                    throw new Error('Chart.js not loaded');
                }
                
                const chartOptions = {
                    responsive: true,
                    maintainAspectRatio: true,
                    aspectRatio: 1,
                    plugins: {
                        legend: {
                            position: 'top',
                            labels: {
                                padding: 20,
                                font: {
                                    size: 12
                                }
                            }
                        }
                    }
                };

                // Activity Chart
                var activityCtx = document.getElementById('userActivityChart').getContext('2d');
                new Chart(activityCtx, {
                    type: 'doughnut',
                    data: {
                        labels: ['Active Users', 'Inactive Users'],
                        datasets: [{
                            data: [$($activeAppUsers.Count), $($inactiveAppUsers.Count)],
                            backgroundColor: ['#16a34a', '#dc2626'],
                            borderWidth: 2
                        }]
                    },
                    options: {
                        ...chartOptions,
                        cutout: '65%'
                    }
                });

                // Job Titles Chart
                var jobTitlesCtx = document.getElementById('jobTitlesChart').getContext('2d');
                const jobTitleCounts = {};
                const table = document.getElementById('users-table');
                const rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');

                for (let row of rows) {
                    const jobTitle = row.cells[1].textContent.trim();
                    jobTitleCounts[jobTitle] = (jobTitleCounts[jobTitle] || 0) + 1;
                }

                const sortedJobTitles = Object.entries(jobTitleCounts)
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 5);

                new Chart(jobTitlesCtx, {
                    type: 'bar',
                    data: {
                        labels: sortedJobTitles.map(([title]) => title),
                        datasets: [{
                            label: 'Number of Users',
                            data: sortedJobTitles.map(([, count]) => count),
                            backgroundColor: '#2563eb',
                            borderRadius: 6
                        }]
                    },
                    options: {
                        ...chartOptions,
                        plugins: {
                            legend: {
                                display: false
                            }
                        },
                        scales: {
                            y: {
                                beginAtZero: true,
                                ticks: {
                                    precision: 0,
                                    font: {
                                        size: 12
                                    }
                                }
                            },
                            x: {
                                ticks: {
                                    font: {
                                        size: 12
                                    }
                                }
                            }
                        }
                    }
                });

                // Departments Chart
                var departmentsCtx = document.getElementById('departmentsChart').getContext('2d');
                
                const departmentCounts = {};
                for (let row of rows) {
                    const department = row.cells[2].textContent.trim();
                    if (department) {
                        departmentCounts[department] = (departmentCounts[department] || 0) + 1;
                    }
                }

                const sortedDepartments = Object.entries(departmentCounts)
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 5);

                new Chart(departmentsCtx, {
                    type: 'bar',
                    data: {
                        labels: sortedDepartments.map(([dept]) => dept),
                        datasets: [{
                            label: 'Number of Users',
                            data: sortedDepartments.map(([, count]) => count),
                            backgroundColor: '#2563eb',
                            borderRadius: 6
                        }]
                    },
                    options: {
                        ...chartOptions,
                        plugins: {
                            legend: {
                                display: false
                            }
                        },
                        scales: {
                            y: {
                                beginAtZero: true,
                                ticks: {
                                    precision: 0,
                                    font: {
                                        size: 12
                                    }
                                }
                            },
                            x: {
                                ticks: {
                                    font: {
                                        size: 12
                                    }
                                }
                            }
                        }
                    }
                });

            } catch (error) {
                console.error('Error creating charts:', error);
                document.querySelector('.charts-grid').innerHTML = 
                    '<div class="chart-card"><p style="color: var(--danger-color)">Error loading charts. Please check your internet connection.</p></div>';
            }

            // Table filtering functionality
            const filterConfig = [
                { dropdownId: 'filter-upn', columnIndex: 0 },
                { dropdownId: 'filter-job-title', columnIndex: 1 },
                { dropdownId: 'filter-department', columnIndex: 2 },
                { dropdownId: 'filter-last-signin', columnIndex: 3 },
                { dropdownId: 'filter-status', columnIndex: 4 },
                { dropdownId: 'filter-groups', columnIndex: 5 }
            ];

            function populateDropdowns() {
                const table = document.getElementById('users-table');
                const rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');

                filterConfig.forEach(config => {
                    const dropdown = document.getElementById(config.dropdownId);
                    const uniqueValues = new Set();

                    for (let row of rows) {
                        const cellValue = row.cells[config.columnIndex].textContent.trim();
                        if (cellValue) uniqueValues.add(cellValue);
                    }

                    dropdown.innerHTML = '<option value="">All</option>';

                    Array.from(uniqueValues)
                        .sort((a, b) => a.localeCompare(b))
                        .forEach(value => {
                            const option = document.createElement('option');
                            option.value = value;
                            option.textContent = value;
                            dropdown.appendChild(option);
                        });
                });
            }

            function filterTable() {
                const table = document.getElementById('users-table');
                const rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');

                for (let row of rows) {
                    let showRow = true;

                    filterConfig.forEach(config => {
                        const dropdown = document.getElementById(config.dropdownId);
                        const dropdownValue = dropdown.value;
                        const cellValue = row.cells[config.columnIndex].textContent.trim();

                        if (dropdownValue && cellValue !== dropdownValue) {
                            showRow = false;
                        }
                    });

                    row.style.display = showRow ? '' : 'none';
                }
            }

            // Pagination functionality
            let currentPage = 1;
            let entriesPerPage = 20;
            let filteredRows = [];
            const table = document.getElementById('users-table');

            function hideAllRows() {
                const tbody = table.getElementsByTagName('tbody')[0];
                const rows = tbody.getElementsByTagName('tr');
                for (let row of rows) {
                    row.style.display = 'none';
                }
            }

            function goToPage(pageNumber) {
                currentPage = pageNumber;
                updateTable();
            }

            function updatePagination() {
                const totalPages = Math.ceil(filteredRows.length / entriesPerPage);
                const pageNumbers = document.getElementById('page-numbers');
                pageNumbers.innerHTML = '';

                // Show maximum 5 page numbers
                let startPage = Math.max(1, currentPage - 2);
                let endPage = Math.min(totalPages, startPage + 4);
                
                if (endPage - startPage < 4) {
                    startPage = Math.max(1, endPage - 4);
                }

                // First page
                if (startPage > 1) {
                    const firstPageBtn = document.createElement('button');
                    firstPageBtn.textContent = '1';
                    firstPageBtn.onclick = () => goToPage(1);
                    pageNumbers.appendChild(firstPageBtn);

                    if (startPage > 2) {
                        const ellipsis = document.createElement('span');
                        ellipsis.textContent = '...';
                        ellipsis.className = 'page-ellipsis';
                        pageNumbers.appendChild(ellipsis);
                    }
                }

                // Page numbers
                for (let i = startPage; i <= endPage; i++) {
                    const pageButton = document.createElement('button');
                    pageButton.textContent = i;
                    pageButton.className = i === currentPage ? 'active' : '';
                    pageButton.onclick = () => goToPage(i);
                    pageNumbers.appendChild(pageButton);
                }

                // Last page
                if (endPage < totalPages) {
                    if (endPage < totalPages - 1) {
                        const ellipsis = document.createElement('span');
                        ellipsis.textContent = '...';
                        ellipsis.className = 'page-ellipsis';
                        pageNumbers.appendChild(ellipsis);
                    }

                    const lastPageBtn = document.createElement('button');
                    lastPageBtn.textContent = totalPages;
                    lastPageBtn.onclick = () => goToPage(totalPages);
                    pageNumbers.appendChild(lastPageBtn);
                }

                // Update pagination info
                document.getElementById('showing-start').textContent = 
                    filteredRows.length === 0 ? 0 : (currentPage - 1) * entriesPerPage + 1;
                document.getElementById('showing-end').textContent = 
                    Math.min(currentPage * entriesPerPage, filteredRows.length);
                document.getElementById('total-entries').textContent = filteredRows.length;

                // Update button states
                document.getElementById('prev-page').disabled = currentPage === 1;
                document.getElementById('next-page').disabled = currentPage === totalPages || totalPages === 0;
            }

            function updateTable() {
                hideAllRows();
                const startIndex = (currentPage - 1) * entriesPerPage;
                const endIndex = Math.min(startIndex + entriesPerPage, filteredRows.length);

                // Show only the rows for the current page
                for (let i = startIndex; i < endIndex; i++) {
                    if (filteredRows[i]) {
                        filteredRows[i].style.display = '';
                    }
                }

                updatePagination();
            }

            function filterTable() {
                const tbody = table.getElementsByTagName('tbody')[0];
                const rows = Array.from(tbody.getElementsByTagName('tr'));
                
                filteredRows = rows.filter(row => {
                    let showRow = true;

                    filterConfig.forEach(config => {
                        const dropdown = document.getElementById(config.dropdownId);
                        const dropdownValue = dropdown.value;
                        const cellValue = row.cells[config.columnIndex].textContent.trim();

                        if (dropdownValue && !cellValue.includes(dropdownValue)) {
                            showRow = false;
                        }
                    });

                    return showRow;
                });

                currentPage = 1;
                updateTable();
            }

            // Event listeners for pagination controls
            document.getElementById('prev-page').onclick = () => {
                if (currentPage > 1) {
                    goToPage(currentPage - 1);
                }
            };

            document.getElementById('next-page').onclick = () => {
                const totalPages = Math.ceil(filteredRows.length / entriesPerPage);
                if (currentPage < totalPages) {
                    goToPage(currentPage + 1);
                }
            };

            document.getElementById('entries-per-page').onchange = (e) => {
                entriesPerPage = parseInt(e.target.value);
                currentPage = 1;
                updateTable();
            };

            // Initialize filtering and pagination
            populateDropdowns();
            filterConfig.forEach(config => {
                document.getElementById(config.dropdownId).addEventListener('change', filterTable);
            });

            // Initial table setup
            filteredRows = Array.from(table.getElementsByTagName('tbody')[0].getElementsByTagName('tr'));
            updateTable();
        })();
    </script>

    <a href="https://github.com/marcel-ngn/" target="_blank" rel="noopener noreferrer" id="floating-github">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M9 19c-5 1.5-5-2.5-7-3m14 6v-3.87a3.37 3.37 0 0 0-.94-2.61c3.14-.35 6.44-1.54 6.44-7A5.44 5.44 0 0 0 20 4.77 5.07 5.07 0 0 0 19.91 1S18.73.65 16 2.48a13.38 13.38 0 0 0-7 0C6.27.65 5.09 1 5.09 1A5.07 5.07 0 0 0 5 4.77a5.44 5.44 0 0 0-1.5 3.78c0 5.42 3.3 6.61 6.44 7A3.37 3.37 0 0 0 9 18.13V22"></path>
        </svg>
        GitHub
    </a>
</body>
</html>
"@
try {
    $htmlContent | Out-File -FilePath $htmlFilePath -Encoding UTF8
    Write-Log "✅ HTML report generated at: $htmlFilePath"
}
catch {
    Write-Log "Error generating HTML report: $_" -Level Error
}

Invoke-Item $htmlFilePath

#EndRegion Report Generation

#Disconnecting all services
Disconnect-AllServices

Write-Log "✅ Disconnected all connections successfully!"
Write-Log "✅ Script execution completed successfully!" -Level Info
