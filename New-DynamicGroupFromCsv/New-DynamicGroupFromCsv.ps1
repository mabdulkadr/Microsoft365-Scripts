<#
.TITLE
    New-DynamicGroupFromCsv - Create dynamic Entra ID groups from CSV input

.SYNOPSIS
    Creates multiple dynamic Azure AD groups from a CSV input.

.DESCRIPTION
    This script connects to Microsoft Graph using either app-based (client credentials) or user-based (delegated) authentication.
    It reads a CSV file containing group details (group name, description, and membership rule criteria) and automatically
    creates dynamic Azure Active Directory (AAD) groups based on device display names or other specified properties.
    Dynamic membership rules are generated per the CSV input, enabling automated and scalable group creation across an Intune-managed environment.

.TAGS
    Identity,M365

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    Group.ReadWrite.All, Directory.Read.All

.AUTHOR
    AI Generated

.VERSION
    1.0.0

.CHANGELOG
    1.0.0 (2026-09-16) - Compliance hardening: canonical rich header + ErrorActionPreference Stop

.LASTUPDATE
    2026-09-16

.EXAMPLE
    & '.\New-DynamicGroupFromCsv.ps1'
    Creates dynamic groups using the CSV path defined in the script.

.EXAMPLE
    Get-Help .\New-DynamicGroupFromCsv.ps1 -Full
    Shows the required CSV schema (GroupName, Description, Rule) and auth setup.

.NOTES
    CSV must include GroupName, Description, and Rule columns.
    Exit codes: 0 = success, 1 = failure.
    Log path: C:\ProgramData\Microsoft365Scripts\Logs\
#>

#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION - solution identity for the embedded logging block.
# ============================================================================

$SolutionName = 'New-DynamicGroupFromCsv'
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

# Permission scopes for user-based authentication (optional if not using app-based auth)
$scopes = "https://graph.microsoft.com/.default"

# Path to the CSV file containing group information
$csvFilePath = "C:\Groups.csv"  # Update with your actual path

# ====================== Function Definitions ======================

# Function to connect to Microsoft Graph
function Connect-ToGraph {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)] [string]$Tenant,
        [Parameter(Mandatory = $false)] [string]$AppId,
        [Parameter(Mandatory = $false)] [string]$AppSecret,
        [Parameter(Mandatory = $false)] [string]$Scopes = "https://graph.microsoft.com/.default"
    )

    Process {
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

            $version = (Get-Module Microsoft.Graph.Authentication | Select-Object -ExpandProperty Version).Major

            if ($AppId) {
                # Ensure all required parameters for app-based authentication are provided
                if (-not $Tenant -or -not $AppId -or -not $AppSecret) {
                    throw "Tenant, AppId, and AppSecret are required for app-based authentication."
                }

                # App-based authentication
                $body = @{
                    grant_type    = "client_credentials"
                    client_id     = $AppId
                    client_secret = $AppSecret
                    scope         = "https://graph.microsoft.com/.default"
                }

                $response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Body $body
                $accessToken = $response.access_token

                if ($version -eq 2) {
                    Write-Log -Message "Version 2 module detected" -Level 'SUCCESS'
                    $accessTokenFinal = ConvertTo-SecureString -String $accessToken -AsPlainText -Force
                } else {
                    Write-Log -Message "Version 1 module detected" -Level 'WARNING'
                    Select-MgProfile -Name Beta
                    $accessTokenFinal = $accessToken
                }
                $graph = Connect-MgGraph -AccessToken $accessTokenFinal
                Write-Log -Message "Connected to Intune tenant $Tenant using app-based authentication" -Level 'SUCCESS'
            } else {
                # User-based authentication
                if ($version -eq 2) {
                    Write-Log -Message "Version 2 module detected" -Level 'SUCCESS'
                } else {
                    Write-Log -Message "Version 1 module detected" -Level 'WARNING'
                    Select-MgProfile -Name Beta
                }
                $graph = Connect-MgGraph -Scopes $Scopes
                Write-Log -Message "Connected to Intune tenant $($graph.TenantId)" -Level 'SUCCESS'
            }
        } catch {
            Write-Error "Failed to connect to Microsoft Graph: $_"
            exit
        }
    }
}

# Function to create a dynamic group in Azure AD
function Create-DynamicGroup {
    param (
        [string]$GroupName,
        [string]$GroupDescription,
        [string]$MembershipRule
    )
    try {
        # Validate that required parameters are not empty
        if ([string]::IsNullOrWhiteSpace($GroupName) -or [string]::IsNullOrWhiteSpace($MembershipRule)) {
            Write-Warning "Skipping group creation due to missing GroupName or MembershipRule."
            return
        }

        Write-Log -Message "Creating dynamic group '$GroupName'..." -Level 'INFO'

        # Define the group properties
        $group = @{
            DisplayName                      = $GroupName
            Description                      = $GroupDescription
            MailEnabled                      = $false
            MailNickname                     = $GroupName.Replace(" ", "").ToLower()
            SecurityEnabled                  = $true
            GroupTypes                       = @("DynamicMembership")
            MembershipRule                   = $MembershipRule
            MembershipRuleProcessingState    = "On"
        }

        # Create the dynamic group in Azure AD
        $newGroup = New-MgGroup @group

        # Output the result in a formatted manner
        Write-Log -Message "Dynamic group '$GroupName' created successfully!" -Level 'SUCCESS'
        Write-Log -Message "Group Name:        $($newGroup.DisplayName)" -Level 'SUCCESS'
        Write-Log -Message "Description:       $($newGroup.Description)" -Level 'INFO'
        Write-Log -Message "Object ID:         $($newGroup.Id)" -Level 'INFO'
        Write-Log -Message "Membership Rule:   $MembershipRule" -Level 'INFO'
    } catch {
        Write-Log -Message "Failed to create dynamic group '$GroupName': $_" -Level 'ERROR'
    }
}

# ====================== Script Execution ======================

# Step 1: Connect to Microsoft Graph using app-based or user-based authentication
Connect-ToGraph -Tenant $tenantID -AppId $appID -AppSecret $appSecretPlain -Scopes $scopes

# Step 2: Read the CSV file and create groups
if (Test-Path $csvFilePath) {
    $groups = Import-Csv -Path $csvFilePath

    foreach ($group in $groups) {
        $groupName = $group.GroupName
        $groupDescription = $group.GroupDescription
        $membershipRule = $group.MembershipRule

        # Create the dynamic group with the specified details
        Create-DynamicGroup -GroupName $groupName -GroupDescription $groupDescription -MembershipRule $membershipRule
    }
} else {
    Write-Error "CSV file not found at path: $csvFilePath"
}

# End of script
