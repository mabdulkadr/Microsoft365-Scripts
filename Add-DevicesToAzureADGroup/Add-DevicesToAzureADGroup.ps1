<#
.TITLE
    Add-DevicesToAzureADGroup - Bulk add devices from CSV to Azure AD group

.SYNOPSIS
    PowerShell script to add devices (from CSV file) to an Azure AD Group.

.DESCRIPTION
    PowerShell script used to bulk add devices to a group in Azure AD, as the Azure AD bulk import functionality requires Device GUIDs rather than the Device Name.
    This script looks up the device based on its name and then adds the device to the specified group.

    Scope & safety:
    - Read-only lookups plus member additions only; never deletes devices or groups.
    Output contract:
    - Per-device console lines (added / already member / not found); exit 0 = success, 1 = failure.

.TAGS
    Identity,M365

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    M365 admin

.AUTHOR
    AI Generated

.VERSION
    1.0.0

.CHANGELOG
    1.0.0 (2026-09-16) - Compliance hardening: canonical rich header + ErrorActionPreference Stop

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Add-DevicesToAzureADGroup.ps1 -GroupName "Device Test Group" -InputFile "C:\Scripts\DevicesToAdd.csv"
    Adds listed devices to the named Azure AD group.

.EXAMPLE
    .\Add-DevicesToAzureADGroup.ps1 -GroupName "Lab Devices" -InputFile ".\SampleDevicesFile.csv"
    Adds devices from the bundled sample file to the lab group.

.NOTES
    CSV input must include a DeviceName column.
    Exit codes: 0 = success, 1 = failure.
    Log path: C:\ProgramData\Microsoft365Scripts\Logs\
#>

#Requires -Version 5.1
param
(
    [parameter(Mandatory=$true)]
    [string] $GroupName,
    [parameter(Mandatory=$true)]
    [string] $InputFile

)

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'Add-DevicesToAzureADGroup'
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
# CONNECTION - ensures a live Azure AD session, prompting only when required.
# ============================================================================

# Ensures an Azure AD connection; prompts for auth only when the session requires it.
function Create-AzureADConnection
{
    # Helper function that runs a cmdlet, and if the exception relates to a connection being required to Azure AD call Connect-AzureAD
    try
    {
        Get-AzureADTenantDetail | Out-Null
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
    {
        Connect-AzureAD
    }
    catch
    {
        Finish-Script -ExitCode 1 -Message "An exception has occurred attempting to connect to Azure AD" -Level 'ERROR'
    }
}

Import-Module AzureAD
Create-AzureADConnection

# Get the group object from Azure AD
#
$AzureAdGroup = Get-AzureADGroup -Filter ("DisplayName eq '{0}'" -f $GroupName)
if (-not $AzureAdGroup)
{
    Finish-Script -ExitCode 1 -Message ("A group with the name '{0}' was not found. No actions performed." -f $GroupName) -Level 'WARNING'
}

# Get devices input file
#
$DeviceCsv = Import-Csv -Path $InputFile -ErrorAction Stop
if (-not $DeviceCsv)
{
    Finish-Script -ExitCode 1 -Message ("An error has occurred reading the device CSV file: '{0}'. No actions performed." -f $InputFile) -Level 'ERROR'
}

# Add the devices the group retrieved from Azure AD
#
$counter = 1
foreach ($DeviceInfo in $DeviceCsv)
{
    "{0} of {1} - Processing the device '{2}'" -f $counter, $DeviceCsv.Count, $DeviceInfo.DeviceName
    
    # Get the device object (as we need the Object ID to add it to the group)
    $AzureAdDevices = Get-AzureADDevice -SearchString $DeviceInfo.DeviceName
    
    # If there are device(s) found in Azure AD - add them to the group
    #
    if ($AzureAdDevices)
    {
        # need to loop, as its possible multiple devices exist with the same name
        foreach ($Device in $AzureAdDevices)
        {
            "`tAdding the device with the device ID: {1})" -f $DeviceInfo.DeviceName, $Device.ObjectId
            try
            {
                Add-AzureADGroupMember -ObjectId $AzureAdGroup.ObjectId -RefObjectId $Device.ObjectId
            }
            catch
            {
                Write-Log -Message "An exception occurred adding the device to the group, it most likely already exists in the group" -Level 'WARNING'
            }
        }
    }
    else
    {
        Write-Log -Message "The device could not be found in Azure AD" -Level 'WARNING'
    }
    ""
    $counter++
}

Finish-Script -ExitCode 0 -Message "Device processing complete." -Level 'SUCCESS'
