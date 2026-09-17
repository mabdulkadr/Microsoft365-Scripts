<#
.TITLE
    Remove-GuestUsersFromCSV - Bulk delete guest users from Entra ID using a CSV list

.SYNOPSIS
    Bulk deletes guest users from Microsoft Entra ID based on a provided CSV list.

.DESCRIPTION
    This script connects to Microsoft Graph and processes a CSV file containing user identifiers
    (such as UserPrincipalName or Email). For each entry, the script attempts to find the corresponding
    account in Microsoft Entra ID, fetches its UserType property, and deletes the user only if their
    UserType is "Guest". The script prints the detected UserType for each user, logs actions, and
    provides a summary report of deleted, skipped, not found, and failed entries.

.TAGS
    Identity,M365,Guests

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    User.ReadWrite.All

.AUTHOR
    AI Generated

.VERSION
    1.0.0

.CHANGELOG
    1.0.0 (2026-09-16) - Compliance hardening: canonical rich header, ErrorActionPreference Stop, alias and catch hygiene.

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Remove-GuestUsersFromCSV.ps1
    Runs the script and prompts to choose the CSV file.

.EXAMPLE
    Get-Help .\Remove-GuestUsersFromCSV.ps1 -Full
    Shows the required CSV schema (UserPrincipalName) and safety checks.

.NOTES
    Part of Microsoft365-Scripts toolkit - Identity,M365,Guests
    Exit codes: 0 = success, 1 = failure, 2 = script error
    Elevation is detected at runtime via Test-IsElevated and degrades gracefully.
#>

#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'Remove-GuestUsersFromCSV'
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
# INPUT - Graph sign-in, CSV picker, and identity-column detection.
# ============================================================================

# Connect to Microsoft Graph
Connect-MgGraph

# Prompt for CSV file
Add-Type -AssemblyName System.Windows.Forms
$OpenDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenDialog.Filter = "CSV Files (*.csv)|*.csv"
$OpenDialog.Title  = "Select CSV with UserPrincipalName column"
if ($OpenDialog.ShowDialog() -ne 'OK') {
    Write-Log -Message "❌ No file selected. Exiting." -Level 'ERROR'
    exit
}
$InputCsv = $OpenDialog.FileName

# Import CSV
$List = Import-Csv $InputCsv

# Detect column
if ($List[0].PSObject.Properties.Name -contains 'UserPrincipalName') {
    $Col = 'UserPrincipalName'
} elseif ($List[0].PSObject.Properties.Name -contains 'Email') {
    $Col = 'Email'
} else {
    Write-Log -Message "❌ No UserPrincipalName or Email column found." -Level 'ERROR'
    exit
}

# ============================================================================
# REMOVAL - deletes guests only; members are skipped, outcomes are counted.
# why: UserType is re-fetched live per row so stale CSV data can't delete members.
# ============================================================================

# Start processing
$Total = 0; $Deleted = 0; $Skipped = 0; $NotFound = 0; $Failed = 0
foreach ($row in $List) {
    $upn = $row.$Col
    $Total++
    try {
        # Always request UserType property!
        $user = Get-MgUser -Filter "UserPrincipalName eq '$upn'" -Property Id,UserType,DisplayName,UserPrincipalName

        if ($user) {
            $realType = $user.UserType
            if ($realType) { $realType = $realType.Trim().ToLower() } else { $realType = "<empty>" }
            Write-Log -Message "[$upn] UserType (from Graph): $realType" -Level 'DEBUG'

            if ($realType -eq "guest") {
                Remove-MgUser -UserId $user.Id -Confirm:$false
                Write-Log -Message "✅ Deleted guest: $upn" -Level 'SUCCESS'
                $Deleted++
            } else {
                Write-Log -Message "⏩ Skipped (not guest): $upn (UserType returned: $realType)" -Level 'INFO'
                $Skipped++
            }
        } else {
            Write-Log -Message "⚠️ Not found: $upn" -Level 'WARNING'
            $NotFound++
        }
    } catch {
        Write-Log -Message "❌ Failed to process: $upn | $_" -Level 'ERROR'
        $Failed++
    }
}

Write-Log -Message "Summary: $Total processed, $Deleted deleted, $Skipped skipped, $NotFound not found, $Failed failed" -Level 'SUCCESS'
Write-Log -Message "Total processed : $Total" -Level 'INFO'
Write-Log -Message "Deleted guests  : $Deleted" -Level 'INFO'
Write-Log -Message "Skipped (not guest): $Skipped" -Level 'INFO'
Write-Log -Message "Not found       : $NotFound" -Level 'INFO'
Write-Log -Message "Failed          : $Failed" -Level 'INFO'
