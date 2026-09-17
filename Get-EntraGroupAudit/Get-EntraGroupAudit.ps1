<#
.TITLE
    Get-EntraGroupAudit - Entra ID Group Membership and Configuration Audit

.SYNOPSIS
    Audits Entra ID group membership, ownership, nesting, and configuration.

.DESCRIPTION
    Provides detailed group analysis including members, direct versus transitive membership, owners, dynamic rules, nested hierarchy, license assignments, and group type classification. Supports single-group deep dive or bulk health audit for empty, large, or ownerless groups.

        Scope & safety:
        - Read-only Graph queries; never modifies groups or memberships.
        Degradation behavior:
        - Missing owners or members render as zero counts without failing the report.
        Output contract:
        - Console summary plus CSV beside the script; exit 0 = success, 1 = failure.

.TAGS
    Reporting,EntraID,Groups,Graph

.PLATFORM
    Windows

.MINROLE
    Intune Service Administrator

.PERMISSIONS
    Directory.Read.All, Group.Read.All, GroupMember.Read.All, User.Read.All

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
    .\\Get-EntraGroupAudit.ps1 -GroupName "SG-Intune-Windows-Devices"
    Runs a deep dive on a single group.

.EXAMPLE
    .\\Get-EntraGroupAudit.ps1 -GroupName "SG-Intune-Pilot" -IncludeMembers -ExportPath "C:\\temp\\members.csv"
    Exports members of a group to CSV.

.EXAMPLE
    .\\Get-EntraGroupAudit.ps1 -BulkAudit -ExportPath "C:\\temp\\group_health.csv"
    Runs a health audit across all groups.

.NOTES
    - Requires Microsoft.Graph.Authentication module.
        - Read-only; no group modifications.
        - Logs: C:\ProgramData\Get-EntraGroupAudit\Logs\
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ByName')]
    [string]$GroupName,

    [Parameter(Mandatory, ParameterSetName = 'ById')]
    [string]$GroupId,

    [Parameter(Mandatory, ParameterSetName = 'BulkAudit')]
    [switch]$BulkAudit,

    [Parameter(ParameterSetName = 'ByName')]
    [Parameter(ParameterSetName = 'ById')]
    [switch]$IncludeMembers,

    [Parameter()]
    [string]$ExportPath
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'Get-EntraGroupAudit'
$ScriptMode   = 'run'

# ============================================================================
# LOGGING BLOCK (embedded canonical scripts/Write-Log.ps1 - copy VERBATIM)
# Single source of truth: Initialize-Log / Write-Banner / Write-Log / Finish-Script.
# ============================================================================

# --- Logging (CLI Configuration) --------------------------------------------
$script:SystemDrive = if ($env:SystemDrive) { $env:SystemDrive.TrimEnd('') } else {
    [System.IO.Path]::GetPathRoot($env:SystemRoot).TrimEnd('')
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
        Write-Log -Message "Log initialization failed: $($_.Exception.Message)" -Level 'ERROR'
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
Write-Log -Message "Script started: Get-EntraGroupAudit" -Level 'INFO'



#region --- Helpers ---

function Get-MgGraphAllPages {
    param([string]$Uri, [string]$Method = 'GET')
    try {
        $response = Invoke-MgGraphRequest -Uri $Uri -Method $Method -ErrorAction Stop
        $results = @()
        if ($null -ne $response.value) { $results += $response.value }
        elseif ($response) { $results += $response }
        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Uri $response.'@odata.nextLink' -Method GET -ErrorAction Stop
            if ($null -ne $response.value) { $results += $response.value }
        }
        return ,$results
    }
    catch [System.Exception] {
        # Graph call failed - logged as verbose
        Write-Verbose "Graph call failed for $Uri : $_"
        return @()
    }
}
#endregion

#region --- Authentication ---
Write-Log -Message "=== AUTHENTICATION ===" -Level 'INFO'
$context = Get-MgContext
if (-not $context) {
    Write-Log -Message "Connecting to Microsoft Graph..." -Level 'INFO'
    Connect-MgGraph -Scopes @(
        'Directory.Read.All',
        'Group.Read.All',
        'GroupMember.Read.All',
        'User.Read.All'
    ) -ErrorAction Stop
    $context = Get-MgContext
}
Write-Log -Message "Signed in as: $($context.Account)" -Level 'INFO'
#endregion

if ($BulkAudit) {
    #region --- Bulk Audit Mode ---
    Write-Log -Message "=== BULK GROUP HEALTH AUDIT ===" -Level 'INFO'

    Write-Log -Message "Fetching all groups..." -Level 'INFO'
    $allGroups = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,groupTypes,securityEnabled,mailEnabled,membershipRule,membershipRuleProcessingState,createdDateTime,renewedDateTime,description,isAssignableToRole"
    Write-Log -Message "$($allGroups.Count) groups found" -Level 'INFO'

    $auditReport = [System.Collections.Generic.List[PSCustomObject]]::new()
    $emptyCount = 0
    $noOwnerCount = 0
    $largeCount = 0
    $dynamicCount = 0
    $groupIndex = 0

    foreach ($g in $allGroups) {
        $groupIndex++
        if ($groupIndex % 50 -eq 0) {
            Write-Progress -Activity "Auditing groups" -Status "$groupIndex of $($allGroups.Count) - $($g.displayName)" -PercentComplete (($groupIndex / $allGroups.Count) * 100)
        }

        # Classify group type
        $isDynamic = $g.groupTypes -contains 'DynamicMembership'
        $isM365 = $g.groupTypes -contains 'Unified'
        $groupTypeLabel = if ($isM365 -and $isDynamic) { 'M365 Dynamic' }
                          elseif ($isM365) { 'M365 Assigned' }
                          elseif ($isDynamic -and $g.securityEnabled) { 'Security Dynamic' }
                          elseif ($g.securityEnabled -and $g.mailEnabled) { 'Mail-Enabled Security' }
                          elseif ($g.securityEnabled) { 'Security Assigned' }
                          elseif ($g.mailEnabled) { 'Distribution' }
                          else { 'Other' }

        if ($isDynamic) { $dynamicCount++ }

        # Get member count (using $count for efficiency)
        $memberCount = 0
        try {
            $countResult = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/members/`$count" -Method GET -Headers @{ 'ConsistencyLevel' = 'eventual' } -ErrorAction Stop
            $memberCount = [int]$countResult
        } catch [System.Exception] { # typed catch - handles Graph or runtime errors
            # Fallback - fetch a page and count
            $membersPage = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/members?`$top=1&`$select=id"
            $memberCount = $membersPage.Count
        }

        if ($memberCount -eq 0) { $emptyCount++ }
        if ($memberCount -ge 500) { $largeCount++ }

        # Get owner count
        $owners = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)/owners?`$select=id,displayName"
        $ownerCount = $owners.Count
        $ownerNames = ($owners | ForEach-Object { $_.displayName }) -join '; '
        if ($ownerCount -eq 0) { $noOwnerCount++ }

        # Determine age
        $groupAge = if ($g.createdDateTime) { [math]::Round(((Get-Date) - [datetime]$g.createdDateTime).TotalDays) } else { 'N/A' }

        # Issues detection
        $issues = @()
        if ($memberCount -eq 0) { $issues += 'Empty group' }
        if ($ownerCount -eq 0) { $issues += 'No owners' }
        if ($memberCount -ge 5000) { $issues += 'Very large (5000+)' }
        if ($isDynamic -and $g.membershipRuleProcessingState -eq 'Paused') { $issues += 'Dynamic rule paused' }

        $auditReport.Add([PSCustomObject]@{
            GroupName             = $g.displayName
            GroupType             = $groupTypeLabel
            MemberCount           = $memberCount
            OwnerCount            = $ownerCount
            Owners                = if ($ownerNames) { $ownerNames } else { '-' }
            SecurityEnabled       = $g.securityEnabled
            MailEnabled           = $g.mailEnabled
            IsRoleAssignable      = $g.isAssignableToRole
            IsDynamic             = $isDynamic
            MembershipRule        = if ($g.membershipRule) { $g.membershipRule } else { '-' }
            RuleProcessingState   = if ($g.membershipRuleProcessingState) { $g.membershipRuleProcessingState } else { '-' }
            CreatedDateTime       = $g.createdDateTime
            GroupAgeDays          = $groupAge
            Description           = if ($g.description) { $g.description } else { '-' }
            Issues                = if ($issues.Count -gt 0) { $issues -join '; ' } else { '-' }
            GroupId               = $g.id
        })
    }

    Write-Progress -Activity "Auditing groups" -Completed

    # Summary
    Write-Log -Message "=== GROUP HEALTH SUMMARY ===" -Level 'INFO'
    Write-Log -Message "Total groups       : $($allGroups.Count)" -Level 'INFO'

    # Type breakdown
    Write-Log -Message "--- Group Types ---" -Level 'WARNING'
    $typeGroups = $auditReport | Group-Object GroupType | Sort-Object Count -Descending
    foreach ($tg in $typeGroups) {
        Write-Log -Message "$($tg.Name) : $($tg.Count)" -Level 'INFO'
    }

    # Health flags
    Write-Log -Message "--- Health Flags ---" -Level 'WARNING'
    Write-Log -Message "Empty groups (0 members)  : $emptyCount" -Level 'INFO'
    Write-Log -Message "No owners assigned        : $noOwnerCount" -Level 'INFO'
    Write-Log -Message "Large groups (500+)       : $largeCount" -Level 'INFO'
    Write-Log -Message "Dynamic groups            : $dynamicCount" -Level 'INFO'

    # Show empty groups
    $emptyGroups = $auditReport | Where-Object { $_.MemberCount -eq 0 } | Sort-Object GroupName
    if ($emptyGroups.Count -gt 0) {
        Write-Log -Message "=== EMPTY GROUPS ($($emptyGroups.Count)) ===" -Level 'INFO'
        foreach ($eg in ($emptyGroups | Select-Object -First 20)) {
            Write-Log -Message "$($eg.GroupName)" -Level 'WARNING'
            Write-Log -Message "| $($eg.GroupType) | Age: $($eg.GroupAgeDays)d" -Level 'DEBUG'
        }
        if ($emptyGroups.Count -gt 20) {
            Write-Log -Message "... and $($emptyGroups.Count - 20) more" -Level 'DEBUG'
        }
    }

    # Show ownerless groups
    $ownerlessGroups = $auditReport | Where-Object { $_.OwnerCount -eq 0 } | Sort-Object GroupName
    if ($ownerlessGroups.Count -gt 0) {
        Write-Log -Message "=== GROUPS WITH NO OWNERS ($($ownerlessGroups.Count)) ===" -Level 'INFO'
        foreach ($og in ($ownerlessGroups | Select-Object -First 20)) {
            Write-Log -Message "$($og.GroupName)" -Level 'WARNING'
            Write-Log -Message "| $($og.GroupType) | Members: $($og.MemberCount)" -Level 'DEBUG'
        }
        if ($ownerlessGroups.Count -gt 20) {
            Write-Log -Message "... and $($ownerlessGroups.Count - 20) more" -Level 'DEBUG'
        }
    }

    # Dual export: raw CSV plus Carbon Dark HTML dashboard (shared timestamp).
    if ($auditReport.Count -gt 0) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $csvPath = if ($ExportPath) { $ExportPath } else { Join-Path $scriptDirectory "GroupHealthAudit_$stamp.csv" }
        $htmlPath = [System.IO.Path]::ChangeExtension($csvPath, '.html')
        $auditReport | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

        # Defensive re-collection: $auditReport must survive for the HTML table (Rule 26).
        $htmlRows = @($auditReport)
        if (-not $htmlRows) { $htmlRows = @() }
        $tableRows = foreach ($rowRef in ($htmlRows | Select-Object -First 500)) {
            $issueBadge = if ($rowRef.Issues -and $rowRef.Issues -ne '-') { 'high' } else { 'low' }
            '<tr><td><code>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.GroupName)") + '</code></td>' +
            '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.GroupType)") + '</td>' +
            '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.MemberCount)") + '</td>' +
            '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.OwnerCount)") + '</td>' +
            '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.GroupAgeDays)") + '</td>' +
            '<td><span class="badge ' + $issueBadge + '">' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.Issues)") + '</span></td></tr>'
        }
        $tableNote = if ($htmlRows.Count -gt 500) { "<p>Showing 500 of $($htmlRows.Count) groups; full data is in the CSV.</p>" } else { '' }
        $tableHtml = '<div class="section-title">Group Health Detail</div>' +
            '<div class="card"><h2>All Groups (' + $htmlRows.Count + ')</h2>' +
            '<table><thead><tr><th>Group</th><th>Type</th><th>Members</th><th>Owners</th><th>Age (days)</th><th>Issues</th></tr></thead><tbody>' +
            ($tableRows -join "`n") + '</tbody></table>' + $tableNote + '</div>'

        $kpis = @(
            @{ value = "$($auditReport.Count)"; label = 'Groups audited'; color = '' },
            @{ value = "$emptyCount"; label = 'Empty groups'; color = '#da1e28' },
            @{ value = "$noOwnerCount"; label = 'Groups without owners'; color = '#f1c21b' },
            @{ value = "$dynamicCount"; label = 'Dynamic groups'; color = '#0f62fe' },
            @{ value = "$largeCount"; label = 'Large groups (500+)'; color = '#8a3ffc' }
        )
        $mgContext = try { Get-MgContext } catch [System.Exception] { $null }
        $tenantId = if ($mgContext) { $mgContext.TenantId } else { '' }
        Export-StandardHtmlReport -OutputPath $htmlPath -Title 'Group Health Audit' -Subtitle ("Groups: $($auditReport.Count) | Empty: $emptyCount | Ownerless: $noOwnerCount") `
            -Tenant $tenantId -Body $tableHtml -Kpis $kpis -Version '1.0.1' -ReportName 'Group Health Audit'
        Write-Log -Message "CSV:  $csvPath ($($auditReport.Count) rows)" -Level 'INFO'
        Write-Log -Message "HTML: $htmlPath" -Level 'INFO'
    }
    #endregion

} else {
    #region --- Single Group Deep Dive ---
    Write-Log -Message "=== RESOLVING GROUP ===" -Level 'INFO'

    if ($GroupName) {
        Write-Log -Message "Searching for group: $GroupName" -Level 'INFO'
        $groups = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$($GroupName -replace "'","''")'"
        if ($groups.Count -eq 0) {
            Write-Log -Message "ERROR: Group '$GroupName' not found." -Level 'ERROR'
            return
        }
        $group = $groups[0]
    } else {
        Write-Log -Message "Looking up group ID: $GroupId" -Level 'INFO'
        $group = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId"
        if (-not $group -or $group.Count -eq 0) {
            Write-Log -Message "ERROR: Group ID '$GroupId' not found." -Level 'ERROR'
            return
        }
        if ($group -is [array]) { $group = $group[0] }
    }

    $gId = $group.id
    $isDynamic = $group.groupTypes -contains 'DynamicMembership'
    $isM365 = $group.groupTypes -contains 'Unified'
    $groupTypeLabel = if ($isM365 -and $isDynamic) { 'M365 Dynamic' }
                      elseif ($isM365) { 'M365 Assigned' }
                      elseif ($isDynamic -and $group.securityEnabled) { 'Security Dynamic' }
                      elseif ($group.securityEnabled -and $group.mailEnabled) { 'Mail-Enabled Security' }
                      elseif ($group.securityEnabled) { 'Security Assigned' }
                      elseif ($group.mailEnabled) { 'Distribution' }
                      else { 'Other' }

    Write-Log -Message "Group Name         : $($group.displayName)" -Level 'INFO'
    Write-Log -Message "Group ID           : $gId" -Level 'INFO'
    Write-Log -Message "Group Type         : $groupTypeLabel" -Level 'INFO'
    Write-Log -Message "Security Enabled   : $($group.securityEnabled)" -Level 'INFO'
    Write-Log -Message "Mail Enabled       : $($group.mailEnabled)" -Level 'INFO'
    Write-Log -Message "Role Assignable    : $($group.isAssignableToRole)" -Level 'INFO'
    Write-Log -Message "Created            : $($group.createdDateTime)" -Level 'INFO'
    if ($group.description) {
        Write-Log -Message "Description        : $($group.description)" -Level 'INFO'
    }
    if ($isDynamic -and $group.membershipRule) {
        Write-Log -Message "Membership Rule    : $($group.membershipRule)" -Level 'INFO'
        Write-Log -Message "Rule Processing    : $($group.membershipRuleProcessingState)" -Level 'INFO'
    }

    # --- Owners ---
    Write-Log -Message "=== GROUP OWNERS ===" -Level 'INFO'
    $owners = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$gId/owners?`$select=id,displayName,userPrincipalName,@odata.type"
    if ($owners.Count -eq 0) {
        Write-Log -Message "No owners assigned." -Level 'WARNING'
    } else {
        Write-Log -Message "$($owners.Count) owner(s):" -Level 'INFO'
        foreach ($o in $owners) {
            $ownerType = switch -Wildcard ($o.'@odata.type') { '*user*' { 'User' } ; '*servicePrincipal*' { 'App' } ; default { '' } }
            Write-Log -Message "$($o.displayName)" -Level 'INFO'
            Write-Log -Message "($($o.userPrincipalName)) [$ownerType]" -Level 'INFO'
        }
    }

    # --- Members ---
    Write-Log -Message "=== GROUP MEMBERS ===" -Level 'INFO'
    $members = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$gId/members?`$select=id,displayName,userPrincipalName,mail,@odata.type,accountEnabled,deviceId,operatingSystem"

    $userMembers = $members | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' }
    $deviceMembers = $members | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.device' }
    $groupMembers = $members | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
    $spMembers = $members | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.servicePrincipal' }

    Write-Log -Message "Total members      : $($members.Count)" -Level 'INFO'
    Write-Log -Message "Users              : $($userMembers.Count)" -Level 'INFO'
    Write-Log -Message "Devices            : $($deviceMembers.Count)" -Level 'INFO'
    Write-Log -Message "Nested groups      : $($groupMembers.Count)" -Level 'INFO'
    Write-Log -Message "Service principals : $($spMembers.Count)" -Level 'INFO'

    # Show nested groups
    if ($groupMembers.Count -gt 0) {
        Write-Log -Message "Nested groups:" -Level 'WARNING'
        foreach ($ng in $groupMembers) {
            Write-Log -Message "$($ng.displayName) ($($ng.id))" -Level 'INFO'
        }
    }

    # User account status
    if ($userMembers.Count -gt 0) {
        $disabledUsers = $userMembers | Where-Object { $_.accountEnabled -eq $false }
        if ($disabledUsers.Count -gt 0) {
            Write-Log -Message "WARNING: $($disabledUsers.Count) disabled user account(s) in this group:" -Level 'WARNING'
            foreach ($du in ($disabledUsers | Select-Object -First 10)) {
                Write-Log -Message "$($du.displayName) ($($du.userPrincipalName))" -Level 'INFO'
            }
            if ($disabledUsers.Count -gt 10) {
                Write-Log -Message "... and $($disabledUsers.Count - 10) more" -Level 'DEBUG'
            }
        }
    }

    # Device OS breakdown
    if ($deviceMembers.Count -gt 0) {
        Write-Log -Message "Device OS breakdown:" -Level 'WARNING'
        $osGroups = $deviceMembers | Group-Object operatingSystem | Sort-Object Count -Descending
        foreach ($os in $osGroups) {
            Write-Log -Message "$($os.Name) : $($os.Count)" -Level 'INFO'
        }
    }

    # --- Parent Groups (nesting) ---
    Write-Log -Message "=== PARENT GROUP MEMBERSHIPS ===" -Level 'INFO'
    $parentGroups = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$gId/transitiveMemberOf?`$select=id,displayName,@odata.type"
    $parentGroupList = $parentGroups | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }

    if ($parentGroupList.Count -eq 0) {
        Write-Log -Message "This group is not nested inside any other groups." -Level 'DEBUG'
    } else {
        Write-Log -Message "Nested inside $($parentGroupList.Count) parent group(s):" -Level 'INFO'
        foreach ($pg in $parentGroupList) {
            Write-Log -Message "$($pg.displayName) ($($pg.id))" -Level 'INFO'
        }
    }

    # --- License Assignments ---
    Write-Log -Message "=== LICENSE ASSIGNMENTS ===" -Level 'INFO'
    try {
        $groupDetail = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$gId`?`$select=assignedLicenses" -ErrorAction Stop
        if ($groupDetail.assignedLicenses -and $groupDetail.assignedLicenses.Count -gt 0) {
            Write-Log -Message "$($groupDetail.assignedLicenses.Count) license(s) assigned to this group:" -Level 'INFO'
            # Resolve SKU IDs to names
            $skus = Get-MgGraphAllPages -Uri "https://graph.microsoft.com/v1.0/subscribedSkus?`$select=skuId,skuPartNumber"
            $skuMap = @{}
            foreach ($sku in $skus) { $skuMap[$sku.skuId] = $sku.skuPartNumber }

            foreach ($lic in $groupDetail.assignedLicenses) {
                $skuName = if ($skuMap.ContainsKey($lic.skuId)) { $skuMap[$lic.skuId] } else { $lic.skuId }
                $disabledPlans = if ($lic.disabledPlans -and $lic.disabledPlans.Count -gt 0) { " ($($lic.disabledPlans.Count) plans disabled)" } else { '' }
                Write-Log -Message "$skuName$disabledPlans" -Level 'INFO'
            }
        } else {
            Write-Log -Message "No licenses assigned to this group." -Level 'DEBUG'
        }
    } catch [System.Exception] { # typed catch - handles Graph or runtime errors
        Write-Log -Message "Could not retrieve license information." -Level 'DEBUG'
    }

    # --- Export Members ---
    if ($IncludeMembers -and $members.Count -gt 0) {
        $memberReport = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($m in $members) {
            $memberType = switch -Wildcard ($m.'@odata.type') {
                '*user*'             { 'User' }
                '*device*'           { 'Device' }
                '*group*'            { 'Nested Group' }
                '*servicePrincipal*' { 'Service Principal' }
                default              { 'Unknown' }
            }

            $memberReport.Add([PSCustomObject]@{
                GroupName         = $group.displayName
                GroupId           = $gId
                GroupType         = $groupTypeLabel
                MemberName        = $m.displayName
                MemberType        = $memberType
                UserPrincipalName = if ($m.userPrincipalName) { $m.userPrincipalName } else { '-' }
                Mail              = if ($m.mail) { $m.mail } else { '-' }
                AccountEnabled    = if ($null -ne $m.accountEnabled) { $m.accountEnabled } else { '-' }
                OperatingSystem   = if ($m.operatingSystem) { $m.operatingSystem } else { '-' }
                MemberId          = $m.id
            })
        }

        if ($ExportPath) {
            $memberReport | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Log -Message "Exported $($memberReport.Count) members to: $ExportPath" -Level 'INFO'
        } else {
            $safeName = $group.displayName -replace '[^\w\-]','_'
            $defaultPath = Join-Path $scriptDirectory "$safeName`_Members_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $memberReport | Export-Csv -Path $defaultPath -NoTypeInformation -Encoding UTF8
            Write-Log -Message "Auto-exported $($memberReport.Count) members to: $defaultPath" -Level 'INFO'
        }
    } elseif (-not $IncludeMembers -and $members.Count -gt 0) {
        Write-Log -Message "Use -IncludeMembers to export the full member list to CSV" -Level 'INFO'
    }
    #endregion
}

Write-Log -Message "`n$('='*60)" -Level 'DEBUG'

# ============================================================================
# HTML REPORT HELPERS (embedded canonical EnterpriseHtmlReport.template.ps1 v1.0.1).
# Six functions copied verbatim (Get-StandardHtmlHead/Open/Footer/Close/ChartScripts + Export-StandardHtmlReport);
# file-level header omitted. IBM Carbon Dark is the only approved HTML design system.
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
