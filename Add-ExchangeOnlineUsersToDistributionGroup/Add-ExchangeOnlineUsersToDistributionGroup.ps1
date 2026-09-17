<#
.TITLE
    Add-ExchangeOnlineUsersToDistributionGroup - Add users from CSV to Exchange distribution group

.SYNOPSIS
    This script adds multiple users from a CSV file to a specified distribution group in Exchange Online.

.DESCRIPTION
    The script reads user email addresses from a CSV file and adds each user to a specified distribution group in Exchange Online.
    It ensures that the necessary modules are installed, prompts the user to choose the CSV file, and to provide the Group ID or name.
    It also asks the user if they want to save a log file and, if confirmed, prompts for the location to save the log file.

.TAGS
    Identity,M365,Exchange

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    ExchangeOnlineManagement

.AUTHOR
    AI Generated

.VERSION
    1.0.0

.CHANGELOG
    1.0.0 (2026-09-16) - Compliance hardening: canonical rich header + ErrorActionPreference Stop

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Add-ExchangeOnlineUsersToDistributionGroup.ps1
    Adds users listed in the chosen CSV file to the target distribution group.

.EXAMPLE
    .\Add-ExchangeOnlineUsersToDistributionGroup.ps1 -ModuleName "ExchangeOnlineManagement"
    Verifies the named module before launching the interactive picker flow.

.NOTES
    CSV input must include a header named Email.
    Exit codes: 0 = success, 1 = failure.
    Log path: C:\ProgramData\Microsoft365Scripts\Logs\
#>

#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'Add-ExchangeOnlineUsersToDistributionGroup'
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
# INTERACTIVE HELPERS - dialogs and safe-connection wrappers for the guided flow.
# ============================================================================

# Installs (if missing) and imports the named module.
function Ensure-Module {
    param (
        [string]$ModuleName
    )


    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Install-Module -Name $ModuleName -Force
    }
    Import-Module $ModuleName
}

# Ensure ExchangeOnlineManagement module is installed
Ensure-Module -ModuleName "ExchangeOnlineManagement"

# Import necessary types for dialogs
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# Connects to Exchange Online, exiting with a message when auth fails.
function Connect-ExchangeOnlineSafely {
    try {
        Connect-ExchangeOnline -UserPrincipalName (Get-Credential).UserName
    } catch {
        Write-Log -Message "Failed to connect to Exchange Online. Exiting." -Level 'ERROR'
        exit
    }
}
Connect-ExchangeOnlineSafely

# Opens a file dialog and returns the chosen CSV path, exiting when cancelled.
function Get-CSVFilePath {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $OpenFileDialog.FilterIndex = 1
    $OpenFileDialog.Multiselect = $false
    
    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $OpenFileDialog.FileName
    } else {
        Write-Log -Message "No file selected. Exiting." -Level 'ERROR'
        exit
    }
}

# Opens a save dialog and returns the chosen log path, exiting when cancelled.
function Get-LogFilePath {
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $SaveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $SaveFileDialog.FilterIndex = 1

    if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $SaveFileDialog.FileName
    } else {
        Write-Log -Message "No save file selected. Exiting." -Level 'ERROR'
        exit
    }
}

# Prompts for one free-text value via input box.
function Get-UserInput {
    param (
        [string]$message
    )
    return [Microsoft.VisualBasic.Interaction]::InputBox($message, "Input Required")
}

# Prompts a Yes/No dialog; returns $true only on Yes.
function Get-UserConfirmation {
    param (
        [string]$message
    )
    $result = [System.Windows.Forms.MessageBox]::Show($message, "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

# ============================================================================
# MAIN - guided flow: pick CSV, name the group, add members, optionally log.
# ============================================================================

# Get the CSV file path
$csvPath = Get-CSVFilePath

# Prompt user for Distribution Group Name
$groupName = Get-UserInput -message "Please enter the Distribution Group Name:"

if ([string]::IsNullOrWhiteSpace($groupName)) {
    Write-Log -Message "No Distribution Group Name provided. Exiting." -Level 'ERROR'
    exit
}

# Import the CSV file
try {
    $users = Import-Csv -Path $csvPath
} catch {
    Write-Log -Message "Failed to import CSV file. Exiting." -Level 'ERROR'
    exit
}

# Prepare log file
$log = @()

# Appends one per-user result row to the in-memory log for optional export.
function Log-Result {
    param (
        [string]$Status,
        [string]$Email,
        [string]$Message
    )
    $log += [PSCustomObject]@{
        Status = $Status
        Email = $Email
        Message = $Message
    }
}

# Add users to the distribution group
foreach ($user in $users) {
    $userPrincipalName = $user.Email
    try {
        # Validate email format
        if (-not [System.Text.RegularExpressions.Regex]::IsMatch($userPrincipalName, "^[^@\s]+@[^@\s]+\.[^@\s]+$")) {
            throw "Invalid email format: $userPrincipalName"
        }

        # Add user to distribution group
        Add-DistributionGroupMember -Identity $groupName -Member $userPrincipalName
        
        # Log success
        Log-Result -Status "SUCCESS" -Email $userPrincipalName -Message "Added to the distribution group."
    } catch {
        # Log failure
        Log-Result -Status "FAILURE" -Email $userPrincipalName -Message $_.Exception.Message
    }
}

# Display results in a table
$log | Format-Table -AutoSize

# Ask the user if they want to save the log file
if (Get-UserConfirmation -message "Do you want to save the log file?") {
    # Get the log file path
    $logFilePath = Get-LogFilePath

    # Save log to file in table format
    try {
        $log | Out-File -FilePath $logFilePath -Force
        Write-Log -Message "Process completed. Log file saved to $logFilePath." -Level 'SUCCESS'
    } catch {
        Write-Log -Message "Failed to save log file. Exiting." -Level 'ERROR'
    }
} else {
    Write-Log -Message "Process completed. Log file not saved." -Level 'SUCCESS'
}
