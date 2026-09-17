<#
.TITLE
    Get-EntraCAReport - Conditional Access Policy Report

.SYNOPSIS
    Generates a comprehensive Conditional Access policy report from Entra ID.

.DESCRIPTION
    Queries Microsoft Graph to retrieve all Conditional Access policies and expands them into a readable report with state, user and group targets, application targets, platform and location conditions, grant controls, and session controls for security auditing and change review.

        Scope & safety:
        - Read-only Graph queries; never modifies policies.
        Degradation behavior:
        - Unresolvable group or app IDs render as raw IDs without failing the report.
        Output contract:
        - Console summary plus CSV beside the script; exit 0 = success, 1 = failure.

.TAGS
    Reporting,EntraID,ConditionalAccess,Graph

.PLATFORM
    Windows

.MINROLE
    Intune Service Administrator

.PERMISSIONS
    Policy.Read.All, Directory.Read.All, Application.Read.All, Group.Read.All

.AUTHOR
    AI Generated

.VERSION
    1.0.1

.CHANGELOG
    1.0.1 (2026-08-26)
    - Migrated to Enterprise Admin standards (canonical header order, structured logging, PS 5.1 contract)
    1.0.0
    - Initial release

.LASTUPDATE
    2026-08-26

.EXAMPLE
    .\Get-EntraCAReport.ps1
    Exports all enabled and report-only Conditional Access policies.

.EXAMPLE
    .\\Get-EntraCAReport.ps1 -PolicyName "MFA" -IncludeDisabled
    Exports policies matching MFA including disabled ones.

.EXAMPLE
    .\\Get-EntraCAReport.ps1 -EnabledOnly -ExportPath "C:\\temp\\ca_policies.csv"
    Exports active policies only to a specific CSV.

.NOTES
    - Requires Microsoft.Graph.Authentication module.
        - Read-only; no policy changes.
        - Logs: C:\ProgramData\Get-EntraCAReport\Logs\
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter()]
    [string]$PolicyName,

    [Parameter()]
    [switch]$EnabledOnly,

    [Parameter()]
    [switch]$IncludeDisabled,

    [Parameter()]
    [string]$ExportPath
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'Get-EntraCAReport'
$ScriptMode   = 'run'

# ============================================================================
# LOGGING BLOCK (embedded canonical scripts/Write-Log.ps1 - copy VERBATIM)
# Single source of truth: Initialize-Log / Write-Banner / Write-Log / Finish-Script.
# ============================================================================

# --- Logging (CLI Configuration) --------------------------------------------
$script:SystemDrive = if ($env:SystemDrive) { $env:SystemDrive.TrimEnd('\') } else {
    [System.IO.Path]::GetPathRoot($env:SystemRoot).TrimEnd('\')
}
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

# ============================================================================
# REPORT OUTPUT ANCHORING (Law 12)
# Anchors relative output paths beside the script using fallback chain.
# ============================================================================

$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot }
elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
else { (Get-Location).Path }

# Resolve relative ExportPath/OutputPath beside the script (Law 12).
if ($PSBoundParameters.ContainsKey('ExportPath') -and $ExportPath -and -not [System.IO.Path]::IsPathRooted($ExportPath)) {
    $ExportPath = Join-Path $scriptDirectory $ExportPath
}
if ($PSBoundParameters.ContainsKey('OutputPath') -and $OutputPath -and -not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path $scriptDirectory $OutputPath
}


# ============================================================================
# MAIN ENTRY LOGGING INITIALIZATION
# ============================================================================

$null = Initialize-Log -SolutionName $SolutionName -ScriptMode $ScriptMode -Type 'General'
Write-Banner
if ($script:LogReady) {
    Write-Log -Message "Log file ready: $($script:LogFile)" -Level 'DEBUG'
}
Write-Log -Message "Script started: Get-EntraCAReport" -Level 'INFO'

# ============================================================================
# AUTHENTICATION - reuses an existing Graph session, else signs in interactively.
# ============================================================================

Write-Log -Message "=== AUTHENTICATION ===" -Level 'INFO'
$mgContext = Get-MgContext
if (-not $mgContext) {
    Connect-MgGraph -Scopes 'Policy.Read.All', 'Directory.Read.All', 'Application.Read.All', 'Group.Read.All' -ErrorAction Stop
    $mgContext = Get-MgContext
}
Write-Log -Message "Signed in as: $($mgContext.Account)" -Level 'SUCCESS'

# ============================================================================
# RESOLVERS - ID-to-name lookups with caches (users/groups/roles, apps, locations).
# ============================================================================

# Resolves Entra object IDs to display names; unknown IDs pass through unchanged.
function Resolve-DirectoryObjectNames {
    param([string[]]$ObjectIds, [hashtable]$Cache)
    $names = @()
    foreach ($id in $ObjectIds) {
        if ($id -eq 'All') { $names += 'All Users'; continue }
        if ($id -eq 'GuestsOrExternalUsers') { $names += 'Guests/External Users'; continue }
        if ($id -eq 'None') { continue }
        if ($Cache.ContainsKey($id)) { $names += $Cache[$id]; continue }
        try {
            $obj = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/directoryObjects/$id?`$select=displayName"
            $first = @($obj)[0]
        } catch [System.Exception] { $first = $null }
        if ($first -and $first.displayName) {
            $odataType = if ($first.'@odata.type') { $first.'@odata.type' } elseif ($first.AdditionalProperties) { $first.AdditionalProperties['@odata.type'] } else { '' }
            $typeShort = switch -Wildcard ("$odataType") {
                '*group*'            { 'Group' }
                '*user*'             { 'User' }
                '*servicePrincipal*' { 'App' }
                '*directoryRole*'    { 'Role' }
                default              { '' }
            }
            $resolved = if ($typeShort) { "$($first.displayName) [$typeShort]" } else { $first.displayName }
            $Cache[$id] = $resolved
            $names += $resolved
        } else {
            $Cache[$id] = $id
            $names += $id
        }
    }
    return $names
}

# Resolves application IDs to names via well-known map, cache, then service principals.
function Resolve-AppNames {
    param([string[]]$AppIds, [hashtable]$Cache)
    $names = @()
    foreach ($id in $AppIds) {
        $wellKnown = switch ($id) {
            'All'                                  { 'All cloud apps' }
            'Office365'                            { 'Office 365' }
            'MicrosoftAdminPortals'                { 'Microsoft Admin Portals' }
            '00000002-0000-0ff1-ce00-000000000000' { 'Office 365 Exchange Online' }
            '00000003-0000-0ff1-ce00-000000000000' { 'Office 365 SharePoint Online' }
            '00000004-0000-0ff1-ce00-000000000000' { 'Skype for Business' }
            '797f4846-ba00-4fd7-ba43-dac1f8f63013' { 'Windows Azure Service Management API' }
            '0000000c-0000-0000-c000-000000000000' { 'Microsoft App Access Panel' }
            '00000002-0000-0000-c000-000000000000' { 'Microsoft Graph (legacy)' }
            '00000003-0000-0000-c000-000000000000' { 'Microsoft Graph' }
            default { $null }
        }
        if ($wellKnown) { $names += $wellKnown; continue }
        if ($Cache.ContainsKey($id)) { $names += $Cache[$id]; continue }
        try {
            $sp = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$id'&`$select=displayName"
            $first = @($sp)[0]
        } catch [System.Exception] { $first = $null }
        if ($first -and $first.displayName) { $Cache[$id] = $first.displayName; $names += $first.displayName }
        else { $Cache[$id] = $id; $names += $id }
    }
    return $names
}

# Resolves named-location IDs via the preloaded map; unknown IDs pass through.
function Resolve-NamedLocations {
    param([string[]]$LocationIds, [hashtable]$LocationMap)
    $names = @()
    foreach ($id in $LocationIds) {
        if ($id -eq 'All') { $names += 'All locations'; continue }
        if ($id -eq 'AllTrusted') { $names += 'All trusted locations'; continue }
        if ($id -eq '00000000-0000-0000-0000-000000000000') { $names += 'MFA Trusted IPs'; continue }
        if ($LocationMap.ContainsKey($id)) { $names += $LocationMap[$id] } else { $names += $id }
    }
    return $names
}

# Joins a list for CSV cells; empty lists render as '-'.
function Format-List { param([array]$Items); if (@($Items).Count -eq 0) { return '-' }; return ($Items -join '; ') }

# ============================================================================
# RETRIEVE - all CA policies with name/state filters.
# ============================================================================

Write-Log -Message "=== RETRIEVING CONDITIONAL ACCESS POLICIES ===" -Level 'INFO'
try {
    $allPolicies = Get-MgGraphAllPages -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'
} catch [System.Exception] {
    Finish-Script -ExitCode 2 -Message "Failed to query CA policies: $($_.Exception.Message)" -Level 'ERROR'
}
$allPolicies = @($allPolicies)
Write-Log -Message "$($allPolicies.Count) total policies found" -Level 'SUCCESS'

if ($PolicyName) {
    $allPolicies = @($allPolicies | Where-Object { $_.displayName -like "*$PolicyName*" })
    Write-Log -Message "Filtered to $($allPolicies.Count) policies matching '$PolicyName'" -Level 'INFO'
}
if ($EnabledOnly) {
    $allPolicies = @($allPolicies | Where-Object { $_.state -eq 'enabled' })
    Write-Log -Message "Filtered to $($allPolicies.Count) enabled policies" -Level 'INFO'
} elseif (-not $IncludeDisabled) {
    $allPolicies = @($allPolicies | Where-Object { $_.state -ne 'disabled' })
    Write-Log -Message "Showing $($allPolicies.Count) enabled/report-only policies (use -IncludeDisabled for all)" -Level 'INFO'
}

if ($allPolicies.Count -eq 0) {
    Write-Log -Message "No policies found matching the specified criteria." -Level 'WARNING'
    return
}

Write-Log -Message "Loading named locations..." -Level 'DEBUG'
try {
    $namedLocations = Get-MgGraphAllPages -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations'
} catch [System.Exception] { $namedLocations = @() }
$namedLocations = @($namedLocations)
$locationMap = @{}
foreach ($nl in $namedLocations) {
    if ($null -ne $nl -and $null -ne $nl.id) { $locationMap[$nl.id] = $nl.displayName }
}
Write-Log -Message "$($namedLocations.Count) named locations loaded" -Level 'DEBUG'

# ============================================================================
# PROCESS - expands each policy into one flat report row.
# ============================================================================

Write-Log -Message "=== PROCESSING POLICIES ===" -Level 'INFO'
$report = [System.Collections.Generic.List[PSCustomObject]]::new()
$nameCache = @{}
$appCache = @{}
$policyIndex = 0

foreach ($policy in $allPolicies) {
    $policyIndex++
    Write-Progress -Activity 'Processing CA policies' -Status "$policyIndex of $($allPolicies.Count) - $($policy.displayName)" -PercentComplete (($policyIndex / [Math]::Max($allPolicies.Count, 1)) * 100)

    $conditions = $policy.conditions
    $grantControls = $policy.grantControls
    $sessionControls = $policy.sessionControls

    $includeUsers = @()
    $excludeUsers = @()
    if ($conditions.users) {
        if ($conditions.users.includeUsers) { $includeUsers += $conditions.users.includeUsers }
        if ($conditions.users.includeGroups) { $includeUsers += Resolve-DirectoryObjectNames -ObjectIds @($conditions.users.includeGroups) -Cache $nameCache }
        if ($conditions.users.includeRoles) { $includeUsers += Resolve-DirectoryObjectNames -ObjectIds @($conditions.users.includeRoles) -Cache $nameCache }
        if ($conditions.users.includeGuestsOrExternalUsers) { $includeUsers += 'Guests/External Users' }
        if ($conditions.users.excludeUsers) { $excludeUsers += Resolve-DirectoryObjectNames -ObjectIds @($conditions.users.excludeUsers) -Cache $nameCache }
        if ($conditions.users.excludeGroups) { $excludeUsers += Resolve-DirectoryObjectNames -ObjectIds @($conditions.users.excludeGroups) -Cache $nameCache }
        if ($conditions.users.excludeRoles) { $excludeUsers += Resolve-DirectoryObjectNames -ObjectIds @($conditions.users.excludeRoles) -Cache $nameCache }
    }

    $includeApps = @()
    $excludeApps = @()
    if ($conditions.applications) {
        if ($conditions.applications.includeApplications) { $includeApps = Resolve-AppNames -AppIds @($conditions.applications.includeApplications) -Cache $appCache }
        if ($conditions.applications.excludeApplications) { $excludeApps = Resolve-AppNames -AppIds @($conditions.applications.excludeApplications) -Cache $appCache }
        if ($conditions.applications.includeUserActions) {
            $includeApps += @($conditions.applications.includeUserActions | ForEach-Object {
                switch ($_) {
                    'urn:user:registersecurityinfo' { 'Register security info' }
                    'urn:user:registerdevice'       { 'Register or join devices' }
                    default { $_ }
                }
            })
        }
    }

    $platforms = '-'
    if ($conditions.platforms) {
        $incPlat = if ($conditions.platforms.includePlatforms) { $conditions.platforms.includePlatforms -join ', ' } else { '' }
        $excPlat = if ($conditions.platforms.excludePlatforms) { " (excl: $($conditions.platforms.excludePlatforms -join ', '))" } else { '' }
        $platforms = "$incPlat$excPlat"
    }

    $includeLocations = '-'
    $excludeLocations = '-'
    if ($conditions.locations) {
        if ($conditions.locations.includeLocations) { $includeLocations = Format-List (Resolve-NamedLocations -LocationIds @($conditions.locations.includeLocations) -LocationMap $locationMap) }
        if ($conditions.locations.excludeLocations) { $excludeLocations = Format-List (Resolve-NamedLocations -LocationIds @($conditions.locations.excludeLocations) -LocationMap $locationMap) }
    }

    $signInRisk = if ($conditions.signInRiskLevels -and @($conditions.signInRiskLevels).Count -gt 0) { $conditions.signInRiskLevels -join ', ' } else { '-' }
    $userRisk = if ($conditions.userRiskLevels -and @($conditions.userRiskLevels).Count -gt 0) { $conditions.userRiskLevels -join ', ' } else { '-' }
    $clientApps = if ($conditions.clientAppTypes -and @($conditions.clientAppTypes).Count -gt 0) { $conditions.clientAppTypes -join ', ' } else { '-' }

    $deviceFilter = '-'
    if ($conditions.devices -and $conditions.devices.deviceFilter) {
        $deviceFilter = "$($conditions.devices.deviceFilter.mode) : $($conditions.devices.deviceFilter.rule)"
    }

    $grantOperator = if ($grantControls.operator) { $grantControls.operator } else { '-' }
    $grants = @()
    if ($grantControls.builtInControls) { $grants += @($grantControls.builtInControls) }
    if ($grantControls.customAuthenticationFactors) { $grants += @($grantControls.customAuthenticationFactors) }
    if ($grantControls.termsOfUse) { $grants += "ToU: $($grantControls.termsOfUse -join ', ')" }
    if ($grantControls.authenticationStrength) { $grants += "Auth Strength: $($grantControls.authenticationStrength.displayName)" }
    $grantText = if ($grants.Count -gt 0) { "($grantOperator) $($grants -join '; ')" } else { 'Block or not configured' }

    $sessionParts = @()
    if ($sessionControls.signInFrequency -and $sessionControls.signInFrequency.isEnabled) {
        $freq = $sessionControls.signInFrequency
        $sessionParts += "Sign-in freq: $($freq.value) $($freq.type)$(if ($freq.frequencyInterval) { " ($($freq.frequencyInterval))" })"
    }
    if ($sessionControls.persistentBrowser -and $sessionControls.persistentBrowser.isEnabled) { $sessionParts += "Persistent browser: $($sessionControls.persistentBrowser.mode)" }
    if ($sessionControls.cloudAppSecurity -and $sessionControls.cloudAppSecurity.isEnabled) { $sessionParts += "Cloud App Security: $($sessionControls.cloudAppSecurity.cloudAppSecurityType)" }
    if ($sessionControls.applicationEnforcedRestrictions -and $sessionControls.applicationEnforcedRestrictions.isEnabled) { $sessionParts += 'App-enforced restrictions' }
    if ($sessionControls.continuousAccessEvaluation -and $sessionControls.continuousAccessEvaluation.mode) { $sessionParts += "CAE: $($sessionControls.continuousAccessEvaluation.mode)" }
    $sessionText = if ($sessionParts.Count -gt 0) { $sessionParts -join '; ' } else { '-' }

    $stateLabel = switch ($policy.state) {
        'enabled'    { 'Enabled' }
        'disabled'   { 'Disabled' }
        'enabledForReportingButNotEnforced' { 'Report-Only' }
        default      { $policy.state }
    }
    Write-Log -Message "[$stateLabel] $($policy.displayName)" -Level $(if ($stateLabel -eq 'Enabled') { 'SUCCESS' } elseif ($stateLabel -eq 'Disabled') { 'ERROR' } else { 'WARNING' })

    $report.Add([PSCustomObject]@{
        PolicyName       = $policy.displayName
        State            = $stateLabel
        CreatedDateTime  = $policy.createdDateTime
        ModifiedDateTime = $policy.modifiedDateTime
        IncludeUsers     = Format-List $includeUsers
        ExcludeUsers     = Format-List $excludeUsers
        IncludeApps      = Format-List $includeApps
        ExcludeApps      = Format-List $excludeApps
        Platforms        = $platforms
        ClientAppTypes   = $clientApps
        IncludeLocations = $includeLocations
        ExcludeLocations = $excludeLocations
        SignInRiskLevels = $signInRisk
        UserRiskLevels   = $userRisk
        DeviceFilter     = $deviceFilter
        GrantControls    = $grantText
        SessionControls  = $sessionText
        PolicyId         = $policy.id
    })
}
Write-Progress -Activity 'Processing CA policies' -Completed

Write-Log -Message "=== CONDITIONAL ACCESS SUMMARY ===" -Level 'INFO'
$stateGroups = @($report | Group-Object State | Sort-Object Name)
foreach ($sg in $stateGroups) {
    Write-Log -Message "$($sg.Name) : $($sg.Count)" -Level $(if ($sg.Name -eq 'Enabled') { 'SUCCESS' } elseif ($sg.Name -eq 'Disabled') { 'ERROR' } else { 'WARNING' })
}

#region --- Helpers ---

# Grant controls breakdown
Write-Log -Message "--- Grant Controls Used ---" -Level 'WARNING'
$grantTypes = @{}
foreach ($r in $report) {
    $r.GrantControls -split ';' | ForEach-Object {
        $g = $_.Trim()
        if ($g -and $g -ne '-' -and $g -ne 'Block or not configured') {
            if (-not $grantTypes.ContainsKey($g)) { $grantTypes[$g] = 0 }
            $grantTypes[$g]++
        }
    }
}
foreach ($gt in ($grantTypes.GetEnumerator() | Sort-Object Value -Descending)) {
    Write-Log -Message "$($gt.Key) : $($gt.Value) policy/policies" -Level 'INFO'
}

# Policies targeting All Users
$allUserPolicies = $report | Where-Object { $_.IncludeUsers -match 'All Users' }
if ($allUserPolicies.Count -gt 0) {
    Write-Log -Message "--- Policies Targeting All Users ($($allUserPolicies.Count)) ---" -Level 'WARNING'
    foreach ($au in $allUserPolicies) {
        Write-Log -Message "[$($au.State)] $($au.PolicyName)" -Level 'INFO'
    }
}

# Policies targeting All Cloud Apps
$allAppPolicies = $report | Where-Object { $_.IncludeApps -match 'All cloud apps' }
if ($allAppPolicies.Count -gt 0) {
    Write-Log -Message "--- Policies Targeting All Cloud Apps ($($allAppPolicies.Count)) ---" -Level 'WARNING'
    foreach ($aa in $allAppPolicies) {
        Write-Log -Message "[$($aa.State)] $($aa.PolicyName)" -Level 'INFO'
    }
}

# Policies with no exclusions (potential lockout risk)
$noExclusions = $report | Where-Object { $_.ExcludeUsers -eq '-' -and $_.State -eq 'Enabled' -and $_.IncludeUsers -match 'All Users' }
if ($noExclusions.Count -gt 0) {
    Write-Log -Message "--- WARNING: Enabled Policies with All Users and No Exclusions ---" -Level 'ERROR'
    foreach ($ne in $noExclusions) {
        Write-Log -Message "$($ne.PolicyName)" -Level 'ERROR'
        Write-Log -Message "Grant: $($ne.GrantControls)" -Level 'DEBUG'
    }
    Write-Log -Message "These policies risk locking out break-glass accounts if misconfigured." -Level 'WARNING'
}

# Report-only policies (might be forgotten)
$reportOnly = $report | Where-Object { $_.State -eq 'Report-Only' }
if ($reportOnly.Count -gt 0) {
    Write-Log -Message "--- Report-Only Policies ($($reportOnly.Count)) ---" -Level 'WARNING'
    Write-Log -Message "Review these to determine if they should be enabled:" -Level 'DEBUG'
    foreach ($ro in $reportOnly) {
        $age = if ($ro.CreatedDateTime) { [math]::Round(((Get-Date) - [datetime]$ro.CreatedDateTime).TotalDays) } else { '?' }
        Write-Log -Message "$($ro.PolicyName) (created ${age}d ago)" -Level 'INFO'
    }
}

# Dual export: raw CSV plus Carbon Dark HTML dashboard (shared timestamp).
if ($report.Count -gt 0) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    if ($ExportPath) {
        $csvPath = $ExportPath
    } else {
        $csvPath = Join-Path $scriptDirectory "CAPolicy_Report_$stamp.csv"
    }
    $htmlPath = [System.IO.Path]::ChangeExtension($csvPath, '.html')
    $report | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    # Defensive re-collection: $report must survive for the HTML table (Rule 26).
    $htmlRows = @($report)
    if (-not $htmlRows) { $htmlRows = @() }
    $tableRows = foreach ($rowRef in ($htmlRows | Select-Object -First 500)) {
        $stateBadge = switch ("$($rowRef.State)") { 'Enabled' { 'low' } 'Report-Only' { 'medium' } 'Disabled' { 'critical' } default { 'low' } }
        '<tr><td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.PolicyName)") + '</td>' +
        '<td><span class="badge ' + $stateBadge + '">' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.State)") + '</span></td>' +
        '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.IncludeUsers)") + '</td>' +
        '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.IncludeApps)") + '</td>' +
        '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.GrantControls)") + '</td></tr>'
    }
    $tableNote = if ($htmlRows.Count -gt 500) { "<p>Showing 500 of $($htmlRows.Count) policies; full data is in the CSV.</p>" } else { '' }
    $tableHtml = '<div class="section-title">Policy Detail</div>' +
        '<div class="card"><h2>All Policies (' + $htmlRows.Count + ')</h2>' +
        '<table><thead><tr><th>Policy</th><th>State</th><th>Users</th><th>Apps</th><th>Grant Controls</th></tr></thead><tbody>' +
        ($tableRows -join "`n") + '</tbody></table>' + $tableNote + '</div>'

    $enabledCount = @($htmlRows | Where-Object { $_.State -eq 'Enabled' }).Count
    $reportOnlyCount = @($htmlRows | Where-Object { $_.State -eq 'Report-Only' }).Count
    $disabledCount = @($htmlRows | Where-Object { $_.State -eq 'Disabled' }).Count
    $riskyCount = @($htmlRows | Where-Object { $_.ExcludeUsers -eq '-' -and $_.State -eq 'Enabled' -and $_.IncludeUsers -match 'All Users' }).Count
    $kpis = @(
        @{ value = "$($htmlRows.Count)"; label = 'Total policies'; color = '' },
        @{ value = "$enabledCount"; label = 'Enabled'; color = '#24a148' },
        @{ value = "$reportOnlyCount"; label = 'Report-only'; color = '#f1c21b' },
        @{ value = "$disabledCount"; label = 'Disabled'; color = '' },
        @{ value = "$riskyCount"; label = 'All-users, no exclusions'; color = '#da1e28' }
    )
    $tenantId = if ($mgContext) { $mgContext.TenantId } else { '' }
    Export-StandardHtmlReport -OutputPath $htmlPath -Title 'Entra CA Policy Report' -Subtitle ("Policies: $($htmlRows.Count) | Enabled: $enabledCount | Lockout-risk: $riskyCount") `
        -Tenant $tenantId -Body $tableHtml -Kpis $kpis -Version '1.0.1' -ReportName 'Entra CA Policy Report'
    Write-Log -Message "CSV:  $csvPath ($($report.Count) rows)" -Level 'INFO'
    Write-Log -Message "HTML: $htmlPath" -Level 'INFO'
}

Write-Log -Message "`n$('='*60)" -Level 'DEBUG'
#endregion

# ============================================================================
# GRAPH PAGINATION (embedded canonical Get-MgGraphAllPages v1.1.0, function only).
# Follows @odata.nextLink with 429 backoff; uses Invoke-MgGraphRequest.
# ============================================================================

function Get-MgGraphAllPages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [int]$DelayMs = 100,
        [hashtable]$Headers = @{},
        [int]$Max429Retries = 3
    )

    [System.Collections.Generic.List[PSCustomObject]]$allResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $nextLink = $Uri
    $requestCount = 0
    $consecutive429 = 0

    do {
        try {
            if ($requestCount -gt 0 -and $DelayMs -gt 0) {
                Start-Sleep -Milliseconds $DelayMs
            }

            $params = @{
                Uri         = $nextLink
                Method      = 'GET'
                Headers     = $Headers
                ErrorAction = 'Stop'
            }

            $response = Invoke-MgGraphRequest @params
            $requestCount++
            $consecutive429 = 0 # Reset throttle counter on success

            if ($null -ne $response.value) {
                foreach ($item in $response.value) {
                    $allResults.Add($item)
                }
            }
            else {
                $allResults.Add($response)
            }

            $nextLink = $response.'@odata.nextLink'

            if ($requestCount % 10 -eq 0) {
                Write-Verbose "Processed $requestCount API pages, retrieved $($allResults.Count) items..."
            }
        }
        catch {
            $is429 = ($_.Exception.Message -like '*429*') -or ($_.Exception.Message -like '*throttled*')
            if ($is429) {
                $consecutive429++
                if ($consecutive429 -gt $Max429Retries) {
                    throw "Rate limit exceeded (HTTP 429). Maximum retries ($Max429Retries) reached for $nextLink"
                }
                # Honor Retry-After header if present, else exponential backoff capped at 60s
                $retryAfter = $null
                try {
                    if ($_.Exception.Response -and $_.Exception.Response.Headers) {
                        $retryAfter = $_.Exception.Response.Headers['Retry-After']
                        if (-not $retryAfter) { $retryAfter = $_.Exception.Response.Headers['retry-after'] }
                    }
                } catch [System.Exception] {
                    $retryAfter = $null
                }
                $delaySec = if ($retryAfter -and [int]::TryParse($retryAfter.ToString().Split(',')[0], [ref]$null)) { [int]$retryAfter.ToString().Split(',')[0] } else { [Math]::Min(60, [Math]::Pow(2, $consecutive429) * 5) }
                Write-Warning "Rate limit hit (attempt $consecutive429/$Max429Retries), waiting $delaySec seconds..."
                Start-Sleep -Seconds $delaySec
                continue
            }
            throw "Error fetching data from $nextLink : $($_.Exception.Message)"
        }
    } while ($nextLink)

    return $allResults
}

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
