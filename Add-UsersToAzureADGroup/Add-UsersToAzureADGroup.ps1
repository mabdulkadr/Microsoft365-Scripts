<#
.TITLE
    Add-UsersToAzureADGroup - Add users or devices from CSV to Intune security group via Graph

.SYNOPSIS
    This script adds users or devices from a CSV file to an Intune Security Group using Microsoft Graph API.

.DESCRIPTION
    - Provides an interactive menu to choose between adding Users or Devices.
    - Uses Microsoft Graph API (modern replacement for AzureAD module).
    - Allows the user to browse for a CSV file instead of hardcoding paths.
    - Checks if each user/device exists in Azure AD before adding them.
    - Verifies if the user/device is already a member of the group to prevent duplicates.
    - Logs successful and failed additions for audit purposes.
    - Implements error handling to log issues encountered.

.TAGS
    Identity,M365,Intune

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    GroupMember.ReadWrite.All, User.Read.All, Device.Read.All

.AUTHOR
    AI Generated

.VERSION
    1.0.0

.CHANGELOG
    1.0.0 (2026-09-16) - Compliance hardening: canonical rich header + ErrorActionPreference Stop

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Add-UsersToAzureADGroup.ps1
    Launches the interactive menu to add users or devices from a chosen CSV file.

.EXAMPLE
    Get-Help .\Add-UsersToAzureADGroup.ps1 -Full
    Shows full help including notes on group ID and CSV path customization.

.NOTES
    Exit codes: 0 = success, 1 = failure.
    Log path: C:\ProgramData\Microsoft365Scripts\Logs\
#>

#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'Add-UsersToAzureADGroup'
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

# Anchor beside-script outputs (Law 12).
$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$intuneLogPath = Join-Path $scriptDirectory 'Intune_Group_Addition_Log.csv'

# Install & Import Microsoft Graph Module if needed
if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    Install-Module Microsoft.Graph -Force
}
Import-Module Microsoft.Graph

# Authenticate with Microsoft Graph (Interactive login)
Try {
    Connect-MgGraph -Scopes "GroupMember.ReadWrite.All", "User.Read.All", "Device.Read.All" -ErrorAction Stop
} Catch {
    Finish-Script -ExitCode 1 -Message "Error connecting to Microsoft Graph: $($_.Exception.Message)" -Level 'ERROR'
}

# Get the Tenant Information
$Tenant = Get-MgOrganization
Write-Log -Message "Connected to Tenant: $($Tenant.DisplayName)" -Level 'INFO'

# Fancy Menu Selection
$selection = @("Add Users to Group", "Add Devices to Group") | Out-GridView -Title "Select Operation" -OutputMode Single

if (-not $selection) {
    Finish-Script -ExitCode 0 -Message "Operation cancelled by user." -Level 'WARNING'
}

# Browse for CSV File
Add-Type -AssemblyName System.Windows.Forms
$FileDialog = New-Object System.Windows.Forms.OpenFileDialog
$FileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$FileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$FileDialog.Title = "Select CSV File"
$null = $FileDialog.ShowDialog()
$CSVFilePath = $FileDialog.FileName

if (-not $CSVFilePath) {
    Finish-Script -ExitCode 1 -Message "No file selected. Exiting script." -Level 'ERROR'
}

# Import CSV File
$Entries = Import-Csv -Path $CSVFilePath -Delimiter ","

# Check if CSV is empty
if ($Entries.Count -eq 0) {
    Finish-Script -ExitCode 1 -Message "Error: No entries found in the CSV file." -Level 'ERROR'
}

# Define Intune Security Group ID (Change This)
$GroupID = "your-group-id"  # Replace with your actual Group ObjectId

# ============================================================================
# MEMBERSHIP ACTIONS - existence-checked, duplicate-safe Graph adds per CSV row.
# ============================================================================

# Adds each CSV user to the group after existence and duplicate checks.
function Add-UserToIntuneGroup {
    foreach ($Entry in $Entries) {
        $UPN = $Entry.UPN  # Extract UserPrincipalName from CSV
        Write-Progress -Activity "Processing: $UPN"

        Try {
            # Step 1: Get User's ObjectId from Microsoft Graph
            $User = Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction Stop

            if (-not $User) {
                Write-Log -Message "User does not exist in Azure AD: $UPN" -Level 'WARNING'
                continue
            }

            # Step 2: Check if the user is already a member of the group
            $ExistingMembers = Get-MgGroupMember -GroupId $GroupID -All | Select-Object -ExpandProperty Id
            if ($ExistingMembers -contains $User.Id) {
                Write-Log -Message "$UPN is already a member of the security group." -Level 'INFO'
                continue
            }

            # Step 3: Add User to Security Group
            New-MgGroupMember -GroupId $GroupID -DirectoryObjectId $User.Id -ErrorAction Stop
            Write-Log -Message "$UPN successfully added to the security group." -Level 'SUCCESS'

            # Log Success
            "$UPN,Success" | Out-File -Append -FilePath $intuneLogPath
        }
        Catch {
            Write-Log -Message "Error adding $UPN : $($_.Exception.Message)" -Level 'ERROR'
            "$UPN,Failed,$($_.Exception.Message)" | Out-File -Append -FilePath $intuneLogPath
        }
    }
}

# Adds each CSV device to the group after existence and duplicate checks.
function Add-DeviceToIntuneGroup {
    foreach ($Entry in $Entries) {
        $DeviceID = $Entry.DeviceID  # Extract Device ID from CSV
        Write-Progress -Activity "Processing: $DeviceID"

        Try {
            # Step 1: Get Device's ObjectId from Microsoft Graph
            $Device = Get-MgDevice -Filter "deviceId eq '$DeviceID'" -ErrorAction Stop

            if (-not $Device) {
                Write-Log -Message "Device does not exist in Azure AD: $DeviceID" -Level 'WARNING'
                continue
            }

            # Step 2: Check if the device is already a member of the group
            $ExistingMembers = Get-MgGroupMember -GroupId $GroupID -All | Select-Object -ExpandProperty Id
            if ($ExistingMembers -contains $Device.Id) {
                Write-Log -Message "$DeviceID is already a member of the security group." -Level 'INFO'
                continue
            }

            # Step 3: Add Device to Security Group
            New-MgGroupMember -GroupId $GroupID -DirectoryObjectId $Device.Id -ErrorAction Stop
            Write-Log -Message "$DeviceID successfully added to the security group." -Level 'SUCCESS'

            # Log Success
            "$DeviceID,Success" | Out-File -Append -FilePath $intuneLogPath
        }
        Catch {
            Write-Log -Message "Error adding $DeviceID : $($_.Exception.Message)" -Level 'ERROR'
            "$DeviceID,Failed,$($_.Exception.Message)" | Out-File -Append -FilePath $intuneLogPath
        }
    }
}

# Execute Based on User Selection
if ($selection -eq "Add Users to Group") {
    Add-UserToIntuneGroup
} elseif ($selection -eq "Add Devices to Group") {
    Add-DeviceToIntuneGroup
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph

Finish-Script -ExitCode 0 -Message "Process completed. Per-row details: $intuneLogPath" -Level 'SUCCESS'
