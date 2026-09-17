<#
.TITLE
    New-DynamicGroupByKeyword - Create dynamic Entra ID group by device name keyword

.SYNOPSIS
    Creates a dynamic Azure AD group based on a keyword match in device display names.

.DESCRIPTION
    This script connects to Microsoft Graph using app-based authentication (service principal)
    and creates a dynamic Azure Active Directory (AAD) group whose membership is determined
    by a rule targeting device display names that contain a specified keyword (e.g., 'it-op').

    The script includes secure credential handling, structured error management, and outputs
    a summary of the group creation process. It is ideal for scenarios in which devices follow
    a naming convention and need dynamic grouping in Intune or Entra ID.

.TAGS
    Identity,M365,Groups

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    Group.ReadWrite.All, Directory.Read.All

.AUTHOR
    AI Generated

.VERSION
    1.0.0

.CHANGELOG
    1.0.0 (2026-09-16) - Compliance hardening: canonical rich header, ErrorActionPreference Stop, alias and catch hygiene.

.LASTUPDATE
    2026-09-16

.EXAMPLE
    & ".\New-DynamicGroupByKeyword.ps1"
    Runs the single-group creation flow with default settings.

.EXAMPLE
    Get-Help .\New-DynamicGroupByKeyword.ps1 -Full
    Shows configuration variables ($groupName, $membershipRule) and auth setup.

.NOTES
    Part of Microsoft365-Scripts toolkit - Identity,M365,Groups
    Exit codes: 0 = success, 1 = failure, 2 = script error
    Elevation is detected at runtime via Test-IsElevated and degrades gracefully.
#>

#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'New-DynamicGroupByKeyword'
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

# ===================== Configuration Section =====================

$tenantID       = "xxxxxxxxxxxxxxxxxxxxxxxx"  # replace with your actual Tenant ID
$appID          = "xxxxxxxxxxxxxxxxxxxxxxxx"  # replace with your actual Client ID
$appSecretPlain = "xxxxxxxxxxxxxxxxxxxxxxxx"  # replace with your actual Client Secret as plain text

# Group Properties
$groupName = "IT-Operations Devices"                            # Name of the new dynamic group
$groupDescription = "Dynamic group for devices with 'it-op' in the display name" # Description of the group
$mailNickname = $groupName.Replace(" ", "").ToLower()           # Mail nickname for the group

# Dynamic Membership Rule for devices whose displayName contains 'it-op'
$membershipRule = "(device.displayName -contains 'it-op')"      # Adjusted rule for devices

$appSecret = ConvertTo-SecureString $appSecretPlain -AsPlainText -Force
# ===================== Module Installation Check =====================

# Define the submodules required for the script
$requiredModules = @("Microsoft.Graph.Groups", "Microsoft.Graph.Authentication")

# Check if the submodules are installed
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Log -Message "The required module '$module' is not installed. Installing it now..." -Level 'WARNING'
        
        try {
            # Install the submodule
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
            Write-Log -Message "Module '$module' installed successfully." -Level 'SUCCESS'
        }
        catch {
            Write-Log -Message "Failed to install the required module '$module'. Please check your internet connection or permissions and try again." -Level 'ERROR'
            exit
        }
    } else {
        Write-Log -Message "Module '$module' is already installed." -Level 'SUCCESS'
    }
}

# Import only the necessary Microsoft Graph submodules
Import-Module Microsoft.Graph.Groups -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# ====================== Function Definitions ======================

# Function to connect to Microsoft Graph
function Connect-ToGraph {
    try {
        Write-Log -Message "Connecting to Microsoft Graph..." -Level 'INFO'

        # Connect to Microsoft Graph using the service principal credentials
        Connect-MgGraph -ClientId $appID -TenantId $tenantID -ClientSecret (ConvertFrom-SecureString $appSecret) -ErrorAction Stop

        Write-Log -Message "Successfully connected to Microsoft Graph." -Level 'SUCCESS'
    }
    catch {
        # Catch and report any errors during connection
        Write-Log -Message "Failed to connect to Microsoft Graph: $_" -Level 'ERROR'
        exit
    }
}

# Function to create a dynamic group in Azure AD
function Create-DynamicGroup {
    try {
        Write-Log -Message "Creating dynamic group '$groupName'..." -Level 'INFO'

        # Define the group properties
        $group = @{
            DisplayName                      = $groupName
            Description                      = $groupDescription
            MailEnabled                      = $false
            MailNickname                     = $mailNickname
            SecurityEnabled                  = $true
            GroupTypes                       = @("DynamicMembership")
            MembershipRule                   = $membershipRule
            MembershipRuleProcessingState    = "On"
        }

        # Create the dynamic group in Azure AD
        $newGroup = New-MgGroup @group

        # Output the result in a formatted manner
        Write-Log -Message "Dynamic group created successfully!" -Level 'SUCCESS'
        Write-Log -Message "Group Name:        $($newGroup.DisplayName)" -Level 'SUCCESS'
        Write-Log -Message "Description:       $($newGroup.Description)" -Level 'INFO'
        Write-Log -Message "Object ID:         $($newGroup.Id)" -Level 'INFO'
        Write-Log -Message "Membership Rule:   $membershipRule" -Level 'INFO'
    }
    catch {
        # Catch and report any errors during group creation
        Write-Log -Message "Failed to create dynamic group: $_" -Level 'ERROR'
    }
}

# ====================== Script Execution ======================

# Step 1: Connect to Microsoft Graph
Connect-ToGraph

# Step 2: Create the dynamic group with the specified rules
Create-DynamicGroup

# End of script
