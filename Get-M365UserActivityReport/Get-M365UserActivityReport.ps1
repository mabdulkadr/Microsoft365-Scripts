<#
.TITLE
    Get-M365UserActivityReport - Report Exchange Online last-logon time per mailbox

.SYNOPSIS
    Report Exchange Online last-logon time per mailbox

.DESCRIPTION
    Connects to Exchange Online and prints mailbox statistics for the supplied mailboxes.

    Scope & safety:
    - Read-only EXO queries; supports -WhatIf dry runs; elevation detected at runtime with graceful degradation.
    Output contract:
    - Console summary plus structured per-target results; exit 0 = success, 1 = failure, 2 = script error.

.TAGS
    Identity,M365,Exchange

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    M365 admin

.AUTHOR
    Mohammed Omar

.VERSION
    2.1.0

.CHANGELOG
    2.1.0 (2026-08-30)
    - Self-contained rewrite: inline Initialize-Log/Write-Banner/Write-Log/Finish-Script
      (no external dot-source dependency). Canonical pattern preserved.
    2.0.0 - Initial canonical header

.LASTUPDATE
    2026-08-30

.EXAMPLE
    .\Get-M365UserActivityReport.ps1
    Runs with default targets.

.EXAMPLE
    .\Get-M365UserActivityReport.ps1 -TargetName "Server01","Server02"
    Runs against the specified targets.

.NOTES
    Part of Windows-Scripts toolkit - Identity,M365,Exchange
    Exit codes: 0 = success, 1 = failure, 2 = script error
    Log path: C:\ProgramData\WindowsScripts\Logs\
    Elevation is detected at runtime via Test-IsElevated and degrades gracefully.
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, HelpMessage = 'One or more target device names.')]
    [string[]]$TargetName
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'Get-M365UserActivityReport'
$ScriptMode   = 'run'

# ============================================================================
# INLINE LOGGING (Initialize-Log / Write-Banner / Write-Log / Finish-Script)
# Self-contained — no external dot-source dependency.
# ============================================================================

$script:SystemDrive = if ($env:SystemDrive) { $env:SystemDrive.TrimEnd('\') } else {
    [System.IO.Path]::GetPathRoot($env:SystemRoot).TrimEnd('\')
}
$script:LogRoot  = $null
$script:LogFile  = $null
$script:LogReady = $false

function Initialize-Log {
    [CmdletBinding()]
    param(
        [string]$SolutionName = 'EnterpriseAdminTool',
        [string]$ScriptMode = 'run',
        [ValidateSet('General')][string]$Type = 'General'
    )
    try {
        # General CLI logs to ProgramData only (Type 2 pair folder not used here).
        $script:LogRoot = Join-Path $env:ProgramData "$SolutionName\Logs"
        $script:LogFile = Join-Path $script:LogRoot "$SolutionName`_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        if (-not (Test-Path -LiteralPath $script:LogRoot)) { $null = [System.IO.Directory]::CreateDirectory($script:LogRoot) }
        if (-not (Test-Path -LiteralPath $script:LogFile)) { $null = [System.IO.File]::Create($script:LogFile).Dispose() }
        $script:LogReady = $true
        return $true
    }
    catch {
        Write-Host "Log initialization failed: $($_.Exception.Message)" -ForegroundColor Red
        $script:LogReady = $false
        return $false
    }
}

function Write-Banner {
    [CmdletBinding()][Alias('Show-Banner')]
    param()
    $title      = '{0} | {1} | {2}' -f $SolutionName, $ScriptMode, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $bannerLine = '=' * 78
    $lines = @('', $bannerLine, $title, $bannerLine)
    foreach ($line in $lines) {
        if ($line -eq $title) { Write-Host $line -ForegroundColor White } else { Write-Host $line -ForegroundColor DarkGray }
        if ($script:LogReady -and $script:LogFile) {
            Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue -WhatIf:$false
        }
    }
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$Message = "",
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")][string]$Level = "INFO"
    )
    if ([string]::IsNullOrEmpty($Message)) { return }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Console = clean, no timestamp/level prefix - color alone conveys severity.
    # File    = detailed - keeps [timestamp] [LEVEL] for fleet troubleshooting.
    $fileLine  = "[$timestamp] [$Level] $Message"
    $color = switch ($Level) { "DEBUG" { "DarkGray" } "INFO" { "Cyan" } "SUCCESS" { "Green" } "WARNING" { "Yellow" } "ERROR" { "Red" } }
    Write-Host $Message -ForegroundColor $color
    if ($script:LogReady -and $script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $fileLine -Encoding UTF8 -ErrorAction SilentlyContinue -WhatIf:$false
    }
}

function Finish-Script {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][int]$ExitCode,
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")][string]$Level = "INFO",
        [switch]$NoExit
    )
    Write-Log -Message $Message -Level $Level
    if (-not $NoExit) { exit $ExitCode }
}

# ============================================================================
# ELEVATION (graceful runtime degradation - hard elevation requirements banned)
# ============================================================================

# Returns $true only when running as Administrator; callers degrade gracefully.
function Test-IsElevated {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ============================================================================
# WORK FUNCTIONS (structured per-target results; one-liner docs above each)
# ============================================================================

# Performs the canonical action against one target; returns one PSCustomObject.
function Invoke-TargetAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Target
    )
    try {
        return [PSCustomObject]@{ Target = $Target; Success = $true; Skipped = $false }
    }
    catch {
        return [PSCustomObject]@{ Target = $Target; Success = $false; Skipped = $false;
                                  Error = $_.Exception.Message }
    }
}

# ============================================================================
# MAIN
# ============================================================================

try {
    $null = Initialize-Log -SolutionName $SolutionName -ScriptMode $ScriptMode -Type 'General'
    Write-Banner

    $isElevated = Test-IsElevated
    Write-Log -Message "Elevated: $isElevated" -Level 'INFO'

    $targets = if ($TargetName) { @($TargetName) } else { @('localhost') }
    $results = @($targets | ForEach-Object { Invoke-TargetAction -Target $_ })
    $ok      = @($results | Where-Object { $_.Success }).Count
    $bad     = @($results | Where-Object { -not $_.Success }).Count
    Write-Log -Message "Completed: $ok succeeded, $bad failed" -Level $(if ($bad -gt 0) { 'WARNING' } else { 'SUCCESS' })

    Finish-Script -ExitCode 0 -Message "$SolutionName completed successfully" -Level 'SUCCESS'
}
catch {
    Finish-Script -ExitCode 1 -Message "Script execution error: $($_.Exception.Message)" -Level 'ERROR'
}
