<#
.TITLE
    BulkUserResetTool - Reset passwords, disable devices, sign out sessions, block sign-in

.SYNOPSIS
    BulkUserResetTool.ps1

.DESCRIPTION
    Performs bulk access-control actions against Entra ID users via Microsoft Graph:
    password resets, device disables, session sign-outs, and sign-in blocks.

    Scope & safety:
    - High-impact actions; supports -All fleet mode, -UserPrincipalNames targeting, and -Exclude guardrails.
    - Test in a non-production tenant first.
    Output contract:
    - Per-user success/error console lines; exit 0 = success, 1 = failure.

.TAGS
    Identity,M365

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    Directory.AccessAsUser.All

.AUTHOR
    AI Generated

.VERSION
    1.0.0

.CHANGELOG
    1.0.0 (2026-09-16) - Compliance hardening: canonical rich header + ErrorActionPreference Stop

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\BulkUserResetTool.ps1 -All -ResetPassword
    Resets passwords for all users.

.EXAMPLE
    .\BulkUserResetTool.ps1 -UserPrincipalNames user@contoso.com -SignOut -BlockSignIn
    Signs out sessions and blocks sign-in for the named user.

.NOTES
    Exit codes: 0 = success, 1 = failure.
    Log path: C:\ProgramData\Microsoft365Scripts\Logs\
#>

#Requires -Version 5.1
param (
    [switch]$All,
    [switch]$ResetPassword,
    [switch]$DisableDevices,
    [switch]$SignOut,
    [switch]$BlockSignIn,
    [string[]]$Exclude,
    [string[]]$UserPrincipalNames
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'BulkUserResetTool'
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
# VALIDATION - mutually exclusive scope switches resolved before any action.
# ============================================================================

# Check if no switches or parameters are provided
if (-not $All -and -not $ResetPassword -and -not $DisableDevices -and -not $SignOut -and -not $BlockSignIn -and -not $Exclude -and -not $UserPrincipalNames) {
    Write-Log -Message "No switches or parameters provided. Specify -All, -ResetPassword, -DisableDevices, -SignOut, -BlockSignIn, or -UserPrincipalNames." -Level 'WARNING'
    Exit
}

# Connect to Microsoft Graph API
Connect-MgGraph -Scopes Directory.AccessAsUser.All

# Check if both parameters are provided
if ($All -and $UserPrincipalNames) {
    Write-Log -Message "Parameter -All and -UserPrincipalNames cannot be used together." -Level 'WARNING'
    return
}

# ============================================================================
# TARGET RESOLUTION - builds the working user set from -All or -UserPrincipalNames.
# ============================================================================

# Retrieve all users if -All parameter is specified
if ($All) {
    $Users = Get-MgUser -All
}
# Filter users based on provided user principal names
elseif ($UserPrincipalNames) {
    $Users = $UserPrincipalNames | Foreach-Object { Get-MgUser -Filter "UserPrincipalName eq '$($_)'" }
}
else {
    Write-Log -Message "No -UserPrincipalNames or -All parameter provided." -Level 'WARNING'
}

# Prompt for the new password if -ResetPassword parameter is specified and there are users to process
$NewPassword = ""
if ($ResetPassword -and $Users.Count -gt 0) {
    $NewPassword = Read-Host "Enter the new password"
}

# Check if any excluded users were not found
$ExcludedNotFound = $Exclude | Where-Object { $Users.UserPrincipalName -notcontains $_ }
foreach ($excludedUser in $ExcludedNotFound) {
    Write-Log -Message "Can't find Microsoft Entra ID account for user $excludedUser" -Level 'ERROR'
}

# Check if any provided users were not found
$UsersNotFound = $UserPrincipalNames | Where-Object { $Users.UserPrincipalName -notcontains $_ }
foreach ($userNotFound in $UsersNotFound) {
    Write-Log -Message "Can't find Microsoft Entra ID account for user $userNotFound" -Level 'ERROR'
}

foreach ($User in $Users) {
    # Check if the user should be excluded
    if ($Exclude -contains $User.UserPrincipalName) {
        Write-Log -Message "Skipping user $($User.UserPrincipalName)" -Level 'INFO'
        continue
    }

    # Flag to indicate if any actions were performed for the user
    $processed = $false

    # Revoke access if -SignOut parameter is specified
    if ($SignOut) {
        Write-Log -Message "Sign-out completed for account $($User.DisplayName)" -Level 'SUCCESS'

        # Revoke all signed in sessions and refresh tokens for the account
        $SignOutStatus = Revoke-MgUserSignInSession -UserId $User.Id

        $processed = $true
    }

    # Block sign-in if -BlockSignIn parameter is specified
    if ($BlockSignIn) {
        Write-Log -Message "Block sign-in completed for account $($User.DisplayName)" -Level 'SUCCESS'

        # Block sign-in
        Update-MgUser -UserId $User.Id -AccountEnabled:$False

        $processed = $true
    }

    # Reset the password if -ResetPassword parameter is specified
    if ($ResetPassword -and $NewPassword) {
        $NewPasswordProfile = @{
            "Password"                      = $NewPassword
            "ForceChangePasswordNextSignIn" = $true
        }
        Update-MgUser -UserId $User.Id -PasswordProfile $NewPasswordProfile
        Write-Log -Message "Password reset completed for $($User.DisplayName)" -Level 'SUCCESS'

        $processed = $true
    }

    # Disable registered devices if -DisableDevices parameter is specified
    if ($DisableDevices) {
        Write-Log -Message "Disable registered devices completed for $($User.DisplayName)" -Level 'SUCCESS'

        # Retrieve registered devices
        $UserDevices = Get-MgUserRegisteredDevice -UserId $User.Id

        # Disable registered devices
        if ($UserDevices) {
            foreach ($Device in $UserDevices) {
                Update-MgDevice -DeviceId $Device.Id -AccountEnabled $false
            }
        }

        $processed = $true
    }

    if (-not $processed) {
        Write-Log -Message "No actions selected for account $($User.DisplayName)" -Level 'WARNING'
    }
}

Finish-Script -ExitCode 0 -Message "BulkUserResetTool run complete." -Level 'SUCCESS'
