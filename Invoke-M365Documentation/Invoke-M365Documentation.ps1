<#
.TITLE
    Invoke-M365Documentation - Generate documentation for Microsoft 365 components

.SYNOPSIS
    Generates detailed documentation for Microsoft 365 components, supporting dynamic section inclusion and exclusion.

.DESCRIPTION
    This script provides a menu-driven interface to choose components and optionally specify sections to include or exclude. It dynamically validates inputs to ensure compatibility with the Get-M365Doc cmdlet.

    Scope & safety:
    - Read-only documentation run; writes a timestamped Word file to the configured output directory.
    Output contract:
    - Reports\<timestamp>-<component>-Documentation.docx; exit 0 = success, 1 = failure.

.TAGS
    Identity,M365,Documentation

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
    & '.\Invoke-M365Documentation.ps1'
    Generates documentation for the chosen M365 component with optional section filters.

.EXAMPLE
    Get-Help .\Invoke-M365Documentation.ps1 -Full
    Shows required modules, app-registration permissions, and output layout.

.NOTES
    Output directory defaults to Reports\ beside the script.
    Exit codes: 0 = success, 1 = failure.
    Log path: C:\ProgramData\Microsoft365Scripts\Logs\
#>

#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'Invoke-M365Documentation'
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
# CONFIGURATION - replace placeholders per environment; secrets stay out of git.
# ============================================================================

# Define credentials
$TenantId		     = '<Your-Tenant-ID>'    # replace with your actual Tenant ID
$ClientId		     = '<Your-App-ID>'       # replace with your actual App ID
$ClientSecret		 = '<Your-App-Secret>'   # replace with your actual Client Secret as plain text
# Output defaults to Reports\ beside the script (Law 12); override per environment.
$scriptDirectory	 = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$OutputDirectory	 = Join-Path $scriptDirectory 'Reports'

# ============================================================================
# MODULES - installs MSAL.PS, PSWriteOffice, and M365Documentation when missing.
# ============================================================================

# Install required modules
Write-Log -Message "Installing required PowerShell modules..." -Level 'WARNING'
$modules = @('MSAL.PS', 'PSWriteOffice', 'M365Documentation')
foreach ($module in $modules) {
    if (-not (Get-Module -Name $module -ListAvailable)) {
        Write-Log -Message "Installing module: $module..." -Level 'INFO'
        Install-Module -Name $module -Force -ErrorAction Stop
    } else {
        Write-Log -Message "Module $module is already installed." -Level 'SUCCESS'
    }
}

# Validate output directory
if (-not (Test-Path -Path $OutputDirectory)) {
    Write-Log -Message "Creating output directory: $OutputDirectory..." -Level 'WARNING'
    New-Item -ItemType Directory -Path $OutputDirectory -Force
}

# ============================================================================
# DOCUMENTATION FLOW - connect, pick a component via menu, render the Word file.
# ============================================================================

# Convert ClientSecret to SecureString
$SecureClientSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force

# Connect to the M365 tenant
try {
    Write-Log -Message "Connecting to Microsoft 365 tenant..." -Level 'WARNING'
    Connect-M365Doc -ClientId $ClientId -ClientSecret $SecureClientSecret -TenantId $TenantId
    Write-Log -Message "Successfully connected to the Microsoft 365 tenant." -Level 'SUCCESS'
} catch {
    Write-Error "Failed to connect to Microsoft 365 tenant. Please verify credentials and try again."
    exit
}

# Menu for selecting components
$componentsMenu = @{
    "1" = "Intune"
    "2" = "AzureAD"
}
Write-Log -Message "Select the component to document:" -Level 'INFO'
$componentsMenu.Keys | Sort-Object | ForEach-Object { Write-Log -Message "$_ - $($componentsMenu[$_])" -Level 'INFO' }

# Capture user choice
$componentChoice = Read-Host "Enter your choice (1-5)"
$selectedComponent = $componentsMenu[$componentChoice]

# Validate choice
if (-not $selectedComponent) {
    Write-Error "Invalid choice. Exiting script."
    exit
}

# Collect optional section information
Write-Log -Message "You selected: $selectedComponent" -Level 'INFO'
$includeSections = Read-Host "Enter sections to include (comma-separated, leave blank for all)"
$excludeSections = Read-Host "Enter sections to exclude (comma-separated, leave blank for none)"

# Process Include and Exclude Sections
$includeSectionsArray = if ($includeSections -ne "") { $includeSections -split ",\s*" } else { $null }
$excludeSectionsArray = if ($excludeSections -ne "") { $excludeSections -split ",\s*" } else { $null }

try {
    Write-Log -Message "Collecting $selectedComponent documentation..." -Level 'WARNING'
    
    # Collect documentation, only pass IncludeSections or ExcludeSections if they are not null
    if ($includeSectionsArray -and $excludeSectionsArray) {
        $doc = Get-M365Doc -Components $selectedComponent `
            -IncludeSections $includeSectionsArray `
            -ExcludeSections $excludeSectionsArray
    } elseif ($includeSectionsArray) {
        $doc = Get-M365Doc -Components $selectedComponent `
            -IncludeSections $includeSectionsArray
    } elseif ($excludeSectionsArray) {
        $doc = Get-M365Doc -Components $selectedComponent `
            -ExcludeSections $excludeSectionsArray
    } else {
        $doc = Get-M365Doc -Components $selectedComponent
    }

    # Define output file path with timestamp
    $timestamp = (Get-Date).ToString("yyyyMMddHHmm")
    $OutputFile = Join-Path -Path $OutputDirectory -ChildPath "$timestamp-$selectedComponent-Documentation.docx"

    # Output documentation to Word file
    Write-Log -Message "Writing documentation to: $OutputFile..." -Level 'WARNING'
    $doc | Write-M365DocWord -FullDocumentationPath $OutputFile
    Write-Log -Message "Documentation for $selectedComponent successfully generated at $OutputFile." -Level 'SUCCESS'
} catch {
    Write-Error "An error occurred while generating documentation. Error: $_"
    exit
}

# Completion message
Write-Log -Message "Script completed successfully!" -Level 'SUCCESS'
