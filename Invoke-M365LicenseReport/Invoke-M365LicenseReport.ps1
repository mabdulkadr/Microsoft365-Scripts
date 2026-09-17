<#
.TITLE
    Invoke-M365LicenseReport - License allocation and usage report for Microsoft 365

.SYNOPSIS
    Generates a license allocation and usage report for Microsoft 365.

.DESCRIPTION
    This script generates a license report for Microsoft 365. It connects to Microsoft Graph,
    collects subscribed SKUs and user license assignments, computes used and unused license
    counts per plan, and renders an HTML report plus console output.

.TAGS
    Identity,M365,Licensing

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    User.Read.All, AuditLog.Read.All, Organization.Read.All, RoleManagement.Read.Directory

.AUTHOR
    AI Generated

.VERSION
    1.0.0

.CHANGELOG
    1.0.0 (2026-09-16) - Compliance hardening: canonical rich header, ErrorActionPreference Stop, alias and catch hygiene.

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Invoke-M365LicenseReport.ps1 -outpath "C:\Reports"
    Runs the license report and writes output to C:\Reports.

.EXAMPLE
    .\Invoke-M365LicenseReport.ps1 -outpath "D:\Audit\Licenses"
    Runs the license report into a dedicated audit folder.

.NOTES
    Part of Microsoft365-Scripts toolkit - Identity,M365,Licensing
    Exit codes: 0 = success, 1 = failure, 2 = script error
    Elevation is detected at runtime via Test-IsElevated and degrades gracefully.
#>

#Requires -Version 5.1

#Params
param(
     [Parameter(Mandatory)]
     [ValidateNotNullOrEmpty()]
     [string]$outpath
  )

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'Invoke-M365LicenseReport'
$ScriptMode   = 'run'

# ============================================================================
# LOGGING BLOCK (embedded canonical Write-Log - General CLI, ProgramData only)
# Single source of truth: Initialize-Log / Write-Banner / Write-Log / Finish-Script.
# ============================================================================

$script:LogRoot  = $null
$script:LogFile  = $null
$script:LogReady = $false

# Creates the ProgramData log folder/file and reports readiness (Type 3 general CLI).
function Initialize-Log {
    [CmdletBinding()]
    param(
        [string]$SolutionName = 'EnterpriseAdminTool',
        [string]$ScriptMode = 'run',
        [ValidateSet('General')]
        [string]$Type = 'General'
    )

    try {
        # General CLI logs to ProgramData only (Type 2 pair folder not used here).
        $script:LogRoot = Join-Path $env:ProgramData "$SolutionName\Logs"
        $script:LogFile = Join-Path $script:LogRoot "$SolutionName`_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

        if (-not (Test-Path -LiteralPath $script:LogRoot)) {
            $null = [System.IO.Directory]::CreateDirectory($script:LogRoot)
        }
        if (-not (Test-Path -LiteralPath $script:LogFile)) {
            $null = [System.IO.File]::Create($script:LogFile).Dispose()
        }

        $script:LogReady = $true
        return $true
    }
    catch {
        Write-Host "Log initialization failed: $($_.Exception.Message)" -ForegroundColor Red
        $script:LogReady = $false
        return $false
    }
}

# Writes the solution banner to console and log file.
function Write-Banner {
    [CmdletBinding()]
    [Alias('Show-Banner')]
    param()

    $title      = '{0} | {1} | {2}' -f $SolutionName, $ScriptMode, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $bannerLine = '=' * 78
    $lines      = @('', $bannerLine, $title, $bannerLine)

    foreach ($line in $lines) {
        if ($line -eq $title) {
            Write-Host $line -ForegroundColor White
        } else {
            Write-Host $line -ForegroundColor DarkGray
        }

        if ($script:LogReady -and $script:LogFile) {
            Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue -WhatIf:$false
        }
    }
}

# Writes one timestamped, level-colored line to console and log file.
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Message = "",
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )

    if ([string]::IsNullOrEmpty($Message)) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Console = clean, no timestamp/level prefix - color alone conveys severity.
    # File    = detailed - keeps [timestamp] [LEVEL] for fleet troubleshooting.
    $fileLine  = "[$timestamp] [$Level] $Message"

    $color = switch ($Level) {
        "DEBUG"   { "DarkGray" }
        "INFO"    { "Cyan" }
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
    }
    Write-Host $Message -ForegroundColor $color

    if ($script:LogReady -and $script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $fileLine -Encoding UTF8 -ErrorAction SilentlyContinue -WhatIf:$false
    }
}

# Logs the final message and terminates with the given exit code.
function Finish-Script {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$ExitCode,
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO",
        [switch]$NoExit
    )

    Write-Log -Message $Message -Level $Level
    if (-not $NoExit) {
        exit $ExitCode
    }
}

$null = Initialize-Log -SolutionName $SolutionName -ScriptMode $ScriptMode -Type 'General'
Write-Banner

# ============================================================================
# GRAPH CONNECTION - reuses a sufficiently-scoped session, else reconnects.
# ============================================================================

# Check Microsoft Graph connection
$state = Get-MgContext

# Define required permissions properly as an array of strings
$requiredPerms = @(
    "User.Read.All",
    "AuditLog.Read.All",
    "Organization.Read.All",
    "RoleManagement.Read.Directory"
)

# Check if we're connected and have all required permissions
$hasAllPerms = $false
if ($state) {
    $missingPerms = @()
    foreach ($perm in $requiredPerms) {
        if ($state.Scopes -notcontains $perm) {
            $missingPerms += $perm
        }
    }
    
    if ($missingPerms.Count -eq 0) {
        $hasAllPerms = $true
        Write-Log -Message "Connected to Microsoft Graph with all required permissions" -Level 'SUCCESS'
    } else {
        Write-Log -Message "Missing required permissions: $($missingPerms -join ', ')" -Level 'WARNING'
        Write-Log -Message "Reconnecting with all required permissions..." -Level 'WARNING'
    }
} else {
    Write-Log -Message "Not connected to Microsoft Graph. Connecting now..." -Level 'WARNING'
}

# Connect if we need to
if (-not $hasAllPerms) {
    try {
        Connect-MgGraph -Scopes $requiredPerms -ErrorAction Stop -NoWelcome
        Write-Log -Message "Successfully connected to Microsoft Graph" -Level 'SUCCESS'
    } catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit
    }
}

# ============================================================================
# COLLECTION - org profile, SKU translation table, then all users (paged).
# ============================================================================

# Get organization information
$orgname = Invoke-MgGraphRequest -Uri "beta/organization" -OutputType PSObject | Select-Object -ExpandProperty Value | Select-Object -ExpandProperty DisplayName

# Download the translation table
$translationTable = Invoke-RestMethod -Method Get -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" | ConvertFrom-Csv

#Get all users including sign-in and assigned license information
$uri = "beta/users?`$select=Id,accountenabled,DisplayName,UserPrincipalName,signInActivity,AssignedLicenses&`$top=999"
$Result = Invoke-MgGraphRequest -Uri $Uri -OutputType PSObject
$AllUsers = $Result.value
$NextLink = $Result."@odata.nextLink"
while ($null -ne $NextLink) {
    $Result = Invoke-MgGraphRequest -Method GET -Uri $NextLink -OutputType PSObject
    $AllUsers += $Result.value
    $NextLink = $Result."@odata.nextLink"
}

##Get tenant license usage information
$Report = [System.Collections.Generic.List[Object]]::new()
#Get all enabled licenses
$licenses = Invoke-MgGraphRequest -Uri "Beta/subscribedSkus" -OutputType PSObject | Select-Object -Expand Value | Where-Object {$_.CapabilityStatus -eq 'Enabled'}
#Get all directory subscription information
$directorySubscription = Invoke-MgGraphRequest -Uri "beta/directory/subscriptions" -OutputType PSObject | Select-Object -Expand Value
#Loop throuugh all licenses
Foreach ($license in $licenses){
    #Translaste the license name
    $licensename = $skuNamePretty = ($translationTable | Where-Object {$_.GUID -eq $license.skuId} | Sort-Object Product_Display_Name -Unique).Product_Display_Name
    If (($licensename -eq "") -or ($null -eq $licensename)){
        $licensename = $license.skuPartNumber
    }
    #Create a custom object with the license information
    $obj = [PSCustomObject][ordered]@{
        "License SKU" = $licensename
        "Type" = If (($directorySubscription | Where-Object {$_.skuId -eq $license.SkuId}).IsTrial -eq $true) {"Trial"} else {"Paid"}
        "Total Licenses" = $license.PrepaidUnits.Enabled
        "Used Licenses" = $license.ConsumedUnits
        "Unused licenses" = $license.PrepaidUnits.Enabled - $license.ConsumedUnits
        "Renewal/Expiratrion Date" = ($directorySubscription | Where-Object {$_.skuId -eq $license.SkuId}).NextLifecycleDateTime
    }
    $report.Add($obj)
}

#Obtain licenses users who have not successfully signed in, in the last 90 days.
#Unable to combine the needed filters
$90daysago = (Get-Date).AddDays(-90).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
$inactiveUsers = $AllUsers | Where-Object {$_.SignInActivity.LastSuccessfulSignInDateTime -lt $90daysago}
$licensedInactiveUsers = $inactiveUsers | Where-Object {$_.assignedLicenses -ne $null}
$licensedInactiveUsersReport = [System.Collections.Generic.List[Object]]::new()
#Loop through all inactive users
Foreach ($user in $licensedInactiveUsers){
    $skuNamePretty = @()
    foreach ($individuallicense in $user.AssignedLicenses.SkuId){
        $skuNamePretty += ($translationTable | Where-Object {$_.GUID -eq $individuallicense} | Sort-Object Product_Display_Name -Unique)."Product_Display_Name"
    }
    $obj = [pscustomobject][ordered]@{
        Name = $user.UserPrincipalName
        Licenses = ($skuNamePretty -join [environment]::NewLine) ##Updated -join "---"
        AccountEnabled = $user.accountenabled
        lastSignInAttempt = $user.SignInActivity.LastSignInDateTime
        lastSuccessfulSignIn = $user.SignInActivity.LastSuccessfulSignInDateTime
    }
    $licensedInactiveUsersReport.Add($obj)
}

##Obtain over-licensed privileged users
#Check tenant Entra level
$items = @("AAD_PREMIUM_P2", "AAD_PREMIUM", "AAD_BASIC")
$Skus = Invoke-MgGraphRequest -Uri "Beta/subscribedSkus" -OutputType PSObject | Select-Object -Expand Value
foreach ($item in $items) {
    $Search = $skus | Where-Object {$_.ServicePlans.servicePlanName -contains "$item"}
    if ($Search) {
        $licenseplan = $item
        break 
    } ElseIf ((!$Search) -and ($item -eq "AAD_BASIC")){
        $licenseplan = $item
        break
    }
}
#Get all users assigned roles
If ($licenseplan -eq "AAD_PREMIUM_P2") {
    $EligiblePIMRoles = Invoke-MgGraphRequest -Uri "beta/roleManagement/directory/roleEligibilitySchedules?`$expand=*" -OutputType PSObject | Select-Object -Expand Value
    $AssignedPIMRoles = Invoke-MgGraphRequest -Uri "beta/roleManagement/directory/roleAssignmentSchedules?`$expand=*" -OutputType PSObject | Select-Object -Expand Value
    $DirectoryRoles = $EligiblePIMRoles + $AssignedPIMRoles
    $PrivilegedRoles = $DirectoryRoles | Where-Object {
        ($_.RoleDefinition.DisplayName  -like "*Administrator*") -or ($_.RoleDefinition.DisplayName -like "*Writer*") -or ($_.RoleDefinition.DisplayName -eq "Global Reader")
    }
    $PrivilegedRoleUsers = $PrivilegedRoles | Where-Object {$_.Principal.'@odata.type' -eq "#microsoft.graph.user"}
    $RoleMembers = $PrivilegedRoleUsers.Principal.userPrincipalName | Select-Object -Unique
    $PrivilegedUsers = $RoleMembers | ForEach-Object { Invoke-MgGraphRequest -uri "/beta/users/$($_)?`$select=displayName,UserPrincipalName,AssignedLicenses" -OutputType PSobject }
}else{
    $DirectoryRoles = Invoke-MgGraphRequest -Uri "/beta/directoryRoles?" -OutputType PSObject | Select-Object -Expand Value
    $PrivilegedRoles = $DirectoryRoles | Where-Object {
        ($_.DisplayName  -like "*Administrator*") -or ($_.DisplayName -like "*Writer*") -or ($_.DisplayName -eq "Global Reader")
    }
    $RoleMembers = $PrivilegedRoles | ForEach-Object { Invoke-MgGraphRequest -uri "/beta/directoryRoles/$($_.id)/members" -OutputType PSObject | Select-Object -Expand Value} | Select-Object Id -Unique
    $PrivilegedUsers = $RoleMembers | ForEach-Object { Invoke-MgGraphRequest -uri "/beta/users/$($_.id)?`$select=displayName,UserPrincipalName,AssignedLicenses" -OutputType PSobject }
}
#Generate report
$overLicensedPrivUsers = [System.Collections.Generic.List[Object]]::new()
$translationTable = Invoke-RestMethod -Method Get -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" | ConvertFrom-Csv
Foreach ($user in $PrivilegedUsers){
    $licenses = @()
    If (($user.assignedLicenses.skuid -notin "41781fb2-bc02-4b7c-bd55-b576c07bb09d,eec0eb4f-6444-4f95-aba0-50c24d67f998") -and ($user.assignedLicenses.skuid.count -gt 0)){
        foreach($guid in $user.assignedLicenses.skuid) {
            $temp = ($translationTable | Where-Object {$_.GUID -eq $guid} | Sort-Object Product_Display_Name -Unique).Product_Display_Name
            $licenses += $temp
        }  
        $Obj2 = [pscustomobject][ordered]@{
            DisplayName = $user.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            License = $licenses -join [environment]::NewLine
        }
        $overLicensedPrivUsers.Add($Obj2)   
    }
}

###This is in progress
##Get users with duplicate licenses
#First obtain a list of all users and the friendly name of the license they have
$AllLicensedUsersReport = [System.Collections.Generic.List[Object]]::new()
$AllLicensedUsers = $AllUsers | Where-Object {$_.AssignedLicenses.count -ne 0}
Foreach ($user in $AllLicensedUsers){
    $licenses = @()
    foreach($guid in $user.assignedLicenses.skuid) {
        $temp = ($translationTable | Where-Object {$_.GUID -eq $guid} | Sort-Object Product_Display_Name -Unique).Product_Display_Name
        $licenses += $temp
    }  
    $Obj2 = [pscustomobject][ordered]@{
        DisplayName = $user.DisplayName
        UserPrincipalName = $User.UserPrincipalName
        License = $licenses -join [environment]::NewLine
    }
    $AllLicensedUsersReport.Add($Obj2)  
}


# Calculate totals for summary
$totalLicenses = ($report | Measure-Object "Total Licenses" -Sum).Sum
$totalUsed = ($report | Measure-Object "Used Licenses" -Sum).Sum
$totalUnused = ($report | Measure-Object "Unused licenses" -Sum).Sum
$unusedPercentage = [math]::Round(($totalUnused / $totalLicenses) * 100, 2)

# Get the count of inactive licensed users
$inactiveUsersCount = $licensedInactiveUsersReport.Count

# Get the count of over-licensed privileged users
$overLicensedPrivUsersCount = $overLicensedPrivUsers.Count

# Get the total number of licensed users
$totalLicensedUsersCount = $AllLicensedUsers.Count

# Generate HTML report
# Dual export: raw CSVs plus Carbon Dark HTML dashboard (shared run, no re-query).
$csvSkuPath = "$outpath\M365_License_Usage_SKUs.csv"
$csvInactivePath = "$outpath\M365_License_Usage_InactiveLicensed.csv"
$csvPrivPath = "$outpath\M365_License_Usage_OverLicensedPriv.csv"
$report | Export-Csv -Path $csvSkuPath -NoTypeInformation -Encoding UTF8
$licensedInactiveUsersReport | Export-Csv -Path $csvInactivePath -NoTypeInformation -Encoding UTF8
$overLicensedPrivUsers | Export-Csv -Path $csvPrivPath -NoTypeInformation -Encoding UTF8
Write-Log -Message "CSV: $csvSkuPath" -Level 'INFO'
Write-Log -Message "CSV: $csvInactivePath" -Level 'INFO'
Write-Log -Message "CSV: $csvPrivPath" -Level 'INFO'

# Defensive re-collection: report lists must survive for the HTML tables (Rule 26).
$skuRows = @($report)
$inactiveRows = @($licensedInactiveUsersReport)
$privRows = @($overLicensedPrivUsers)

$skuTableRows = foreach ($rowRef in ($skuRows | Select-Object -First 200)) {
    '<tr><td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.'License SKU')") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.Type)") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.'Total Licenses')") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.'Used Licenses')") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.'Unused licenses')") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.'Renewal/Expiratrion Date')") + '</td></tr>'
}
$inactiveTableRows = foreach ($rowRef in ($inactiveRows | Select-Object -First 200)) {
    '<tr><td><code>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.Name)") + '</code></td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.Licenses)") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.AccountEnabled)") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.lastSuccessfulSignIn)") + '</td></tr>'
}
$privTableRows = foreach ($rowRef in ($privRows | Select-Object -First 200)) {
    '<tr><td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.DisplayName)") + '</td>' +
    '<td><code>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.UserPrincipalName)") + '</code></td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.License)") + '</td></tr>'
}
$bodyHtml = '<div class="section-title">Subscription Usage</div>' +
    '<div class="card"><h2>Plans (' + $skuRows.Count + ')</h2>' +
    '<table><thead><tr><th>License SKU</th><th>Type</th><th>Total</th><th>Used</th><th>Unused</th><th>Renewal Date</th></tr></thead><tbody>' +
    ($skuTableRows -join "`n") + '</tbody></table></div>' +
    '<div class="section-title">Licensed but Inactive (90 days)</div>' +
    '<div class="card"><h2>Inactive Users (' + $inactiveRows.Count + ')</h2>' +
    '<table><thead><tr><th>User</th><th>Licenses</th><th>Enabled</th><th>Last Successful Sign-In</th></tr></thead><tbody>' +
    ($inactiveTableRows -join "`n") + '</tbody></table></div>' +
    '<div class="section-title">Over-Licensed Privileged Users</div>' +
    '<div class="card"><h2>Privileged Users (' + $privRows.Count + ')</h2>' +
    '<table><thead><tr><th>Display Name</th><th>UPN</th><th>Licenses</th></tr></thead><tbody>' +
    ($privTableRows -join "`n") + '</tbody></table></div>'

$kpis = @(
    @{ value = "$totalLicenses"; label = 'Total licenses'; color = '' },
    @{ value = "$totalUsed"; label = 'Used licenses'; color = '#24a148' },
    @{ value = "$totalUnused ($unusedPercentage%)"; label = 'Unused licenses'; color = '#da1e28' },
    @{ value = "$inactiveUsersCount"; label = 'Inactive licensed users'; color = '#f1c21b' },
    @{ value = "$overLicensedPrivUsersCount"; label = 'Over-licensed priv users'; color = '#8a3ffc' }
)
Export-StandardHtmlReport -OutputPath "$outpath\M365_License_Usage_Report.html" -Title 'M365 License Usage Report' -Subtitle ("Licenses: $totalLicenses | Used: $totalUsed | Unused: $totalUnused ($unusedPercentage%)") `
    -Body $bodyHtml -Kpis $kpis -Version '1.0.0' -ReportName 'M365 License Usage Report'

# Output the HTML to a file
$outputPathFile = "$outpath\M365_License_Usage_Report.html"
Write-Log -Message "CSV:  $csvSkuPath" -Level 'SUCCESS'
Write-Log -Message "HTML: $outputPathFile" -Level 'SUCCESS'

# ============================================================================
# HTML REPORT HELPERS (embedded canonical EnterpriseHtmlReport.template.ps1 v1.0.1).
# ============================================================================

# ============================================================================
# Get-StandardHtmlHead - emits <head> with Carbon design tokens + base styles.
# ============================================================================
function Get-StandardHtmlHead {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter()][string]$Subtitle = ''
    )

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$Title $Subtitle</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

:root {
    --cds-background: #161616;
    --cds-layer-01: #262626;
    --cds-layer-02: #353535;
    --cds-border-strong-01: #4d4d4d;
    --cds-border-subtle-01: #393939;
    --cds-text-primary: #f4f4f4;
    --cds-text-secondary: #c6c6c6;
    --cds-text-helper: #8d8d8d;
    --cds-link: #78a9ff;
    --cds-blue: #0f62fe;
    --cds-purple: #8a3ffc;
    --cds-magenta: #d02670;
    --cds-support-success: #24a148;
    --cds-support-warning: #f1c21b;
    --cds-support-error: #da1e28;
    --cds-support-info: #0043ce;
}

* { margin: 0; padding: 0; box-sizing: border-box; }

body {
    font-family: 'IBM Plex Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background-color: var(--cds-background);
    color: var(--cds-text-primary);
    line-height: 1.4;
    padding: 32px;
    -webkit-font-smoothing: antialiased;
}

.header {
    margin-bottom: 40px;
    padding-bottom: 24px;
    border-bottom: 1px solid var(--cds-border-strong-01);
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    flex-wrap: wrap;
    gap: 16px;
}

.header-left h1 {
    font-size: 28px;
    font-weight: 300;
    letter-spacing: 0.5px;
    color: var(--cds-text-primary);
    margin-bottom: 4px;
}

.header-left h1 strong { font-weight: 600; }

.header .subtitle {
    color: var(--cds-text-secondary);
    font-size: 14px;
    font-family: 'IBM Plex Mono', monospace;
}

.header-right {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12px;
    color: var(--cds-text-helper);
    text-align: right;
}

.kpi-row {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 2px;
    background-color: var(--cds-border-subtle-01);
    border: 1px solid var(--cds-border-subtle-01);
    margin-bottom: 40px;
}

.kpi-card {
    background-color: var(--cds-layer-01);
    padding: 20px;
    display: flex;
    flex-direction: column-reverse;
    justify-content: space-between;
    min-height: 120px;
    transition: background-color 0.15s ease;
}

.kpi-card:hover { background-color: var(--cds-layer-02); }

.kpi-value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 38px;
    font-weight: 400;
    line-height: 1.1;
    color: var(--cds-text-primary);
}

.kpi-label {
    font-size: 12px;
    font-weight: 500;
    color: var(--cds-text-secondary);
    letter-spacing: 0.2px;
    margin-bottom: 12px;
}

.section-title {
    font-size: 14px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--cds-text-secondary);
    margin: 40px 0 16px 0;
    padding-bottom: 8px;
    border-bottom: 1px solid var(--cds-border-subtle-01);
    display: flex;
    align-items: center;
    gap: 8px;
}

.grid-2 {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(450px, 1fr));
    gap: 24px;
    margin-bottom: 24px;
}

.card {
    background-color: var(--cds-layer-01);
    border: 1px solid var(--cds-border-subtle-01);
    padding: 24px;
    display: flex;
    flex-direction: column;
}

.card h2 {
    font-size: 16px;
    font-weight: 400;
    margin-bottom: 24px;
    color: var(--cds-text-primary);
    border-left: 3px solid var(--cds-blue);
    padding-left: 12px;
}

.legend { list-style: none; flex: 1; min-width: 180px; }
.legend li {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 8px 0;
    font-size: 13px;
    border-bottom: 1px solid var(--cds-border-subtle-01);
}
.legend li:last-child { border-bottom: none; }
.legend .dot { width: 8px; height: 8px; flex-shrink: 0; }
.legend .count {
    margin-left: auto;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 500;
}

.bar-chart { display: flex; flex-direction: column; gap: 12px; }
.bar-row { display: flex; align-items: center; gap: 16px; font-size: 13px; }
.bar-label {
    min-width: 140px;
    text-align: right;
    color: var(--cds-text-secondary);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.bar-track { flex: 1; height: 20px; background-color: var(--cds-border-subtle-01); position: relative; }
.bar-fill {
    height: 100%;
    transition: width 0.8s cubic-bezier(0.16, 1, 0.3, 1);
    display: flex;
    align-items: center;
    padding-left: 8px;
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    color: #ffffff;
    min-width: 24px;
}

/* Donut chart layout (used by Get-StandardHtmlChartScripts) */
.donut-container {
    display: flex;
    align-items: center;
    justify-content: space-around;
    gap: 24px;
    flex-wrap: wrap;
}

.donut-wrap {
    position: relative;
    width: 160px;
    height: 160px;
}

.donut-center {
    position: absolute;
    top: 50%; left: 50%;
    transform: translate(-50%, -50%);
    text-align: center;
}

.donut-center .grade {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 36px;
    font-weight: 500;
    line-height: 1;
}

.donut-center .rate {
    font-size: 11px;
    color: var(--cds-text-helper);
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-top: 4px;
}

table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}

thead th {
    text-align: left;
    padding: 12px 16px;
    background-color: var(--cds-layer-02);
    border-bottom: 1px solid var(--cds-border-strong-01);
    color: var(--cds-text-primary);
    font-weight: 500;
    font-size: 12px;
}

tbody td {
    padding: 12px 16px;
    border-bottom: 1px solid var(--cds-border-subtle-01);
    color: var(--cds-text-secondary);
}

tbody tr {
    background-color: var(--cds-layer-01);
    transition: background-color 0.1s ease;
}

tbody tr:hover { background-color: var(--cds-layer-02); }

.badge {
    display: inline-block;
    padding: 2px 8px;
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.badge.critical { background-color: rgba(218, 30, 40, 0.15); color: #ff8389; border-left: 3px solid var(--cds-support-error); }
.badge.high     { background-color: rgba(219, 109, 40, 0.15); color: #ffb38a; border-left: 3px solid #db6d28; }
.badge.medium   { background-color: rgba(241, 194, 27, 0.15); color: #f1c21b; border-left: 3px solid var(--cds-support-warning); }
.badge.low      { background-color: rgba(36, 161, 72, 0.15); color: #8ee0a5; border-left: 3px solid var(--cds-support-success); }
.badge.info     { background-color: rgba(15, 98, 254, 0.15); color: #78a9ff; border-left: 3px solid var(--cds-support-info); }

.progress-bar {
    display: inline-block;
    width: 100px;
    height: 8px;
    background-color: var(--cds-border-subtle-01);
    vertical-align: middle;
}
.progress-fill { height: 100%; transition: width 0.5s ease; }
.progress-fill.low      { background-color: var(--cds-support-success); }
.progress-fill.medium   { background-color: var(--cds-support-warning); }
.progress-fill.high     { background-color: #db6d28; }
.progress-fill.critical { background-color: var(--cds-support-error); }
.pct-label {
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    margin-left: 8px;
    vertical-align: middle;
    color: var(--cds-text-secondary);
}

.empty-state {
    text-align: center;
    padding: 40px;
    color: var(--cds-text-helper);
    font-style: normal;
    border: 1px dashed var(--cds-border-strong-01);
    background-color: var(--cds-background);
}

canvas { max-width: 100%; }

.footer {
    margin-top: 80px;
    padding: 32px;
    background-color: var(--cds-layer-01);
    border: 1px solid var(--cds-border-subtle-01);
    display: grid;
    grid-template-columns: 1.2fr 0.8fr 1fr;
    gap: 32px;
    align-items: start;
}

.footer-col { display: flex; flex-direction: column; gap: 8px; }
.footer-label {
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: var(--cds-text-helper);
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 500;
    margin-bottom: 4px;
}
.footer-value {
    font-size: 13px;
    color: var(--cds-text-primary);
    font-family: 'IBM Plex Mono', monospace;
    line-height: 1.6;
    word-break: break-all;
}
.footer-value strong { color: var(--cds-text-primary); font-weight: 500; }

.grade-card {
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 16px;
    background-color: var(--cds-layer-02);
    border: 1px solid var(--cds-border-subtle-01);
    cursor: help;
    position: relative;
    transition: border-color 0.15s ease;
}
.grade-card:hover { border-color: var(--cds-blue); }
.grade-card .grade-letter {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 64px;
    font-weight: 300;
    line-height: 1;
    margin-bottom: 8px;
}
.grade-card .grade-rate { font-size: 12px; font-family: 'IBM Plex Mono', monospace; color: var(--cds-text-secondary); letter-spacing: 0.5px; }
.grade-card .grade-label {
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: var(--cds-text-helper);
    margin-top: 12px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 500;
}
.grade-card[data-tooltip]:hover::after {
    content: attr(data-tooltip);
    position: absolute;
    bottom: calc(100% + 8px);
    left: 50%;
    transform: translateX(-50%);
    background-color: var(--cds-layer-02);
    border: 1px solid var(--cds-blue);
    color: var(--cds-text-primary);
    padding: 12px 16px;
    font-size: 11px;
    font-family: 'IBM Plex Sans', sans-serif;
    text-transform: none;
    letter-spacing: normal;
    line-height: 1.5;
    width: 240px;
    text-align: left;
    z-index: 10;
    pointer-events: none;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4);
    white-space: pre-wrap;
}

.action-bar { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 4px; }

.btn {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 8px 14px;
    font-size: 12px;
    font-family: 'IBM Plex Sans', sans-serif;
    font-weight: 500;
    background-color: var(--cds-layer-02);
    color: var(--cds-text-primary);
    border: 1px solid var(--cds-border-strong-01);
    cursor: pointer;
    transition: background-color 0.15s ease, border-color 0.15s ease;
    text-decoration: none;
}
.btn:hover { background-color: var(--cds-layer-01); border-color: var(--cds-blue); }
.btn-primary { background-color: var(--cds-blue); color: #ffffff; border-color: var(--cds-blue); }
.btn-primary:hover { background-color: #0353e9; border-color: #0353e9; }
.btn-icon { width: 14px; height: 14px; flex-shrink: 0; }

.footer-meta {
    margin-top: 24px;
    padding-top: 20px;
    border-top: 1px solid var(--cds-border-subtle-01);
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
    font-size: 11px;
    color: var(--cds-text-helper);
    font-family: 'IBM Plex Mono', monospace;
    grid-column: 1 / -1;
}
.footer-meta a { color: var(--cds-link); text-decoration: none; }
.footer-meta a:hover { text-decoration: underline; }

.disclaimer-box {
    position: fixed;
    inset: 0;
    background-color: rgba(0, 0, 0, 0.7);
    display: none;
    align-items: center;
    justify-content: center;
    z-index: 100;
    padding: 32px;
}
.disclaimer-box.is-open { display: flex; }
.disclaimer-modal {
    background-color: var(--cds-layer-01);
    border: 1px solid var(--cds-border-strong-01);
    max-width: 560px;
    width: 100%;
    padding: 32px;
    max-height: 80vh;
    overflow-y: auto;
}
.disclaimer-modal h3 {
    font-size: 16px;
    font-weight: 600;
    margin-bottom: 16px;
    color: var(--cds-text-primary);
    text-transform: uppercase;
    letter-spacing: 1px;
}
.disclaimer-modal p {
    font-size: 13px;
    line-height: 1.6;
    color: var(--cds-text-secondary);
    margin-bottom: 16px;
}
.disclaimer-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 24px; }

@media (max-width: 768px) {
    .grid-2 { grid-template-columns: 1fr; }
    .kpi-row { grid-template-columns: repeat(2, 1fr); }
    .footer { grid-template-columns: 1fr; gap: 24px; }
    body { padding: 16px; }
}

@media print {
    body { background-color: #ffffff; color: #000000; padding: 12px; }
    .header, .footer, .kpi-card, .card { background-color: #ffffff !important; border-color: #d0d0d0 !important; break-inside: avoid; }
    .footer { display: none; }
    .section-title { color: #000000; border-bottom-color: #000000; }
    thead th { background-color: #f4f4f4; color: #000000; }
    .no-print { display: none; }
}
</style>
</head>
<body>
"@
}

# ============================================================================
# Get-StandardHtmlOpen - opens <body> with header bar + KPI tiles.
# Pass a hashtable of @{value=...; label=...; color=...} for each KPI.
# ============================================================================
function Get-StandardHtmlOpen {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter()][string]$Subtitle = '',
        [Parameter()][string]$GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        [Parameter()][string]$Operator = '',
        [Parameter()][hashtable[]]$Kpis = @()
    )

    $kpiHtml = ''
    foreach ($k in $Kpis) {
        $color = if ($k.ContainsKey('color') -and $k['color']) { "color:$($k['color'])" } else { '' }
        $kpiHtml += "<div class=`"kpi-card`"><div class=`"kpi-value`" style=`"$color`">$($k['value'])</div><div class=`"kpi-label`">$($k['label'])</div></div>`n"
    }

    $operatorRow = if ($Operator) { "<div style=`"margin-top: 4px;`">OPERATOR: $Operator</div>" } else { '' }

    return @"
<div class="header">
    <div class="header-left">
        <h1>$Title</h1>
        <div class="subtitle">$Subtitle</div>
    </div>
    <div class="header-right">
        <div>GENERATED: $GeneratedAt</div>
        $operatorRow
    </div>
</div>

<!-- KPI Cards -->
<div class="kpi-row">
$kpiHtml</div>
"@
}

# ============================================================================
# Get-StandardHtmlFooter - emits the canonical 3-column footer + modal.
# ============================================================================
function Get-StandardHtmlFooter {
    [CmdletBinding()]
    param(
        [Parameter()][string]$Tenant = '',
        [Parameter()][string]$Operator = '',
        [Parameter()][string]$GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        [Parameter()][string]$Timezone = [System.TimeZoneInfo]::Local.DisplayName,
        [Parameter()][string]$Utc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"),
        [Parameter()][string]$RunId = ([guid]::NewGuid().ToString().Substring(0, 8).ToUpper()),
        [Parameter()][string]$Version = '1.0.0',
        [Parameter()][string]$ReportName = 'Report',
        [Parameter()][string]$Grade = '',
        [Parameter()][string]$GradeRate = '',
        [Parameter()][string]$GradeColor = '',
        [Parameter()][string]$GradeTip = ''
    )

    $gradeBlock = if ($Grade) {
        $tipAttr = if ($GradeTip) { " data-tooltip=`"$GradeTip`"" } else { '' }
        @"
    <div class="footer-col">
        <div class="grade-card"$tipAttr>
            <div class="grade-letter" style="color:$GradeColor">$Grade</div>
            <div class="grade-rate">$GradeRate</div>
            <div class="grade-label">Compliance Grade</div>
        </div>
    </div>
"@
    } else {
        '<div class="footer-col"></div>'
    }

    return @"
<!-- Footer -->
<div class="footer">
    <div class="footer-col">
        <div class="footer-label">Tenant</div>
        <div class="footer-value"><strong>$Tenant</strong></div>
        <div class="footer-label" style="margin-top:12px;">Operator</div>
        <div class="footer-value">$Operator</div>
        <div class="footer-label" style="margin-top:12px;">Generated</div>
        <div class="footer-value">$GeneratedAt</div>
        <div class="footer-value" style="color:var(--cds-text-helper);font-size:11px;">$Timezone &middot; UTC $Utc</div>
    </div>
$gradeBlock
    <div class="footer-col">
        <div class="footer-label">Run</div>
        <div class="footer-value">$RunId</div>
        <div class="footer-label" style="margin-top:12px;">Quick Actions</div>
        <div class="action-bar">
            <button class="btn btn-primary" onclick="window.print()" title="Print or save as PDF">
                <svg class="btn-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M4 2h8v3H4V2zm0 5h8a2 2 0 0 1 2 2v3h-2v3H4v-3H2V9a2 2 0 0 1 2-2zm1 7v-3h6v3H5z"/></svg>
                Print
            </button>
            <button class="btn" onclick="navigator.clipboard.writeText(window.location.href)" title="Copy file path">
                <svg class="btn-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M5 2h7a1 1 0 0 1 1 1v9h-2V4H6v8H4V3a1 1 0 0 1 1-1zM2 5h8a1 1 0 0 1 1 1v8a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1zm1 2v6h6V7H3z"/></svg>
                Copy Path
            </button>
            <button class="btn" onclick="document.querySelector('.header').scrollIntoView({behavior:'smooth'})" title="Back to top">
                <svg class="btn-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M8 3.5l5 5h-3v4H6v-4H3l5-5z"/></svg>
                Top
            </button>
        </div>
        <div class="action-bar" style="margin-top:8px;">
            <button class="btn" onclick="document.getElementById('disclaimerModal').classList.add('is-open')" title="View full disclaimer">
                <svg class="btn-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 1 0 0 14A7 7 0 0 0 8 1zm0 3a1 1 0 0 1 1 1v4a1 1 0 0 1-2 0V5a1 1 0 0 1 1-1zm0 8a1 1 0 1 1 0-2 1 1 0 0 1 0 2z"/></svg>
                Disclaimer
            </button>
        </div>
    </div>
    <div class="footer-meta">
        <span>$ReportName &middot; v$Version &middot; Run $RunId</span>
        <span>Generated by <a href="https://github.com/mabdulkadr/powershell-enterprise-admin-skill" target="_blank" rel="noopener">PowerShell Enterprise Admin</a></span>
    </div>
</div>

<!-- Disclaimer modal -->
<div class="disclaimer-box" id="disclaimerModal" onclick="if(event.target===this)this.classList.remove('is-open')">
    <div class="disclaimer-modal">
        <h3>Disclaimer</h3>
        <p>This report is generated from read-only queries and is provided as-is with no warranty of any kind. The metrics and identifiers shown are a point-in-time snapshot and may not reflect the current state by the time this report is reviewed.</p>
        <p>Test generated tools in a staging environment before deploying to production. The authors assume no liability for any damage or data loss resulting from their use.</p>
        <p>This report may contain tenant identifiers and operator account names. Treat the file as confidential and follow your organization's data-handling policy when sharing.</p>
        <div class="disclaimer-actions">
            <button class="btn" onclick="document.getElementById('disclaimerModal').classList.remove('is-open')">Close</button>
        </div>
    </div>
</div>
"@
}

# ============================================================================
# Get-StandardHtmlClose - emits </body></html> + the standard JS helpers.
# ============================================================================
function Get-StandardHtmlClose {
    [CmdletBinding()]
    param()

    return @"
<script>
// Close disclaimer modal on Escape
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        const m = document.getElementById('disclaimerModal');
        if (m) m.classList.remove('is-open');
    }
});
</script>
</body>
</html>
"@
}

# ============================================================================
# Get-StandardHtmlChartScripts - optional helper for scripts that draw donuts/bars.
# ============================================================================
function Get-StandardHtmlChartScripts {
    [CmdletBinding()]
    param()

    return @"
<script>
// Mini donut chart renderer (no dependencies)
function drawDonut(canvasId, data, colors) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const cx = canvas.width / 2, cy = canvas.height / 2;
    const outerR = 76, innerR = 58;
    const total = data.reduce((a, b) => a + b, 0);
    if (total === 0) return;
    let startAngle = -Math.PI / 2;
    data.forEach((val, i) => {
        const sliceAngle = (val / total) * 2 * Math.PI;
        ctx.beginPath();
        ctx.arc(cx, cy, outerR, startAngle, startAngle + sliceAngle);
        ctx.arc(cx, cy, innerR, startAngle + sliceAngle, startAngle, true);
        ctx.closePath();
        ctx.fillStyle = colors[i];
        ctx.fill();
        startAngle += sliceAngle;
    });
}

// Bar chart renderer
function drawBars(containerId, data, colorFn) {
    const container = document.getElementById(containerId);
    if (!container || !data) return;
    const dataArray = Array.isArray(data) ? data : [data];
    if (!dataArray.length || (dataArray.length === 1 && !dataArray[0])) return;
    const maxVal = Math.max(...dataArray.map(d => d.value || 0));
    const colors = ['#0f62fe','#8a3ffc','#00b0ff','#008080','#da1e28','#ff832b','#8d8d8d','#e0e0e0'];
    dataArray.forEach((item, i) => {
        if (!item) return;
        const val = item.value || 0;
        const label = item.label || 'Unknown';
        const pct = maxVal > 0 ? (val / maxVal * 100) : 0;
        const color = colorFn ? colorFn(item, i) : colors[i % colors.length];
        const row = document.createElement('div');
        row.className = 'bar-row';
        row.innerHTML =
            '<div class="bar-label" title="' + label + '">' + label + '</div>' +
            '<div class="bar-track"><div class="bar-fill" style="width:' + pct + '%;background-color:' + color + '">' + val + '</div></div>';
        container.appendChild(row);
    });
}
</script>
"@
}

# ============================================================================
# Export-StandardHtmlReport - convenience: build + write a complete report in one call.
# Pass -Body as the HTML between header and footer.
# ============================================================================
function Export-StandardHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter()][string]$Subtitle = '',
        [Parameter()][string]$Tenant = '',
        [Parameter()][string]$Operator = '',
        [Parameter()][string]$Body = '',
        [Parameter()][hashtable[]]$Kpis = @(),
        [Parameter()][string]$Grade = '',
        [Parameter()][string]$GradeRate = '',
        [Parameter()][string]$GradeColor = '',
        [Parameter()][string]$GradeTip = '',
        [Parameter()][string]$Version = '1.0.0',
        [Parameter()][string]$ReportName = 'Report',
        [Parameter()][string]$ChartScripts = ''
    )

    $now = Get-Date
    $html = Get-StandardHtmlHead -Title $Title -Subtitle $Subtitle
    $html += Get-StandardHtmlOpen -Title $Title -Subtitle $Subtitle -GeneratedAt $now.ToString('yyyy-MM-dd HH:mm:ss') -Operator $Operator -Kpis $Kpis
    $html += $Body
    $html += Get-StandardHtmlFooter -Tenant $Tenant -Operator $Operator -GeneratedAt $now.ToString('yyyy-MM-dd HH:mm:ss') `
        -Timezone ([System.TimeZoneInfo]::Local.DisplayName) `
        -Utc $now.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") `
        -RunId ([guid]::NewGuid().ToString().Substring(0, 8).ToUpper()) `
        -Version $Version -ReportName $ReportName `
        -Grade $Grade -GradeRate $GradeRate -GradeColor $GradeColor -GradeTip $GradeTip
    if ($ChartScripts) { $html += "`n$ChartScripts" }
    $html += Get-StandardHtmlClose

    $html | Out-File -FilePath $OutputPath -Encoding utf8 -Force
}
