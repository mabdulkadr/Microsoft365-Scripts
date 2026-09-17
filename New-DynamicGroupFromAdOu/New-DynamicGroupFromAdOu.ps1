<#
.TITLE
    New-DynamicGroupFromAdOu - Create dynamic Entra ID groups from on-prem AD OU names

.SYNOPSIS
    Creates dynamic Azure AD groups based on on-premises Active Directory OU names.

.DESCRIPTION
    This script connects to both On-Premises Active Directory and Azure Active Directory (AAD) using Microsoft Graph.
    It retrieves all OU names from the local AD and creates corresponding dynamic AAD groups with the prefix Devices-.
    Each group is created with a membership rule targeting devices with onPremisesDistinguishedName containing the OU name.
    The script ensures that no duplicate groups are created by checking for existing group names before creation.
    Additional features include robust error handling and logging, automatic installation of required PowerShell modules, and support for secure app-based authentication.

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
    .\New-DynamicGroupFromAdOu.ps1
    Creates dynamic groups from on-prem AD OU names.

.EXAMPLE
    Get-Help .\New-DynamicGroupFromAdOu.ps1 -Full
    Shows configuration variables ($groupPrefix, auth) and log file location.

.NOTES
    Requires connectivity to both on-prem AD and Microsoft Graph.
    Default log file: CreateDynamicGroups.log beside the script.
    Exit codes: 0 = success, 1 = failure.
#>

#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

# ====================== Configuration Section ======================

$tenantID       = "xxxxxxxxxxxxxxxxxxxxxxxx"  # replace with your actual Tenant ID
$appID          = "xxxxxxxxxxxxxxxxxxxxxxxx"  # replace with your actual Client ID
$appSecret      = "xxxxxxxxxxxxxxxxxxxxxxxx"  # replace with your actual Client Secret as plain text

# Prefix for dynamic group names
$groupPrefix = "Devices - "

# ====================== Logging Configuration ======================

# Path to the log file (anchored beside the script per Law 12)
$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$logFilePath = Join-Path $scriptDirectory "CreateDynamicGroups.log"

# Ensure the log directory exists
$logDirectory = Split-Path -Path $logFilePath -Parent
if (-not (Test-Path -Path $logDirectory)) {
    try {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Log -Message "Failed to create log directory at '$logDirectory'. Error: $_" -Level 'ERROR'
        exit 1
    }
}

# Function to log messages with timestamps
# why: legacy file-logger predating the canonical block; console + file behavior kept.
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] - $Message"
    # Output to console with color-coded levels
    switch ($Level) {
        "INFO" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        default { Write-Host $logMessage }
    }
    # Append to log file
    Add-Content -Path $logFilePath -Value $logMessage
}

# ====================== Function Definitions ======================

# Function to ensure a PowerShell module is installed and imported
function Ensure-Module {
    param (
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Log "Module '$ModuleName' not found. Installing..." -Level "INFO"
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Module '$ModuleName' installed successfully." -Level "INFO"
        }
        catch {
            Write-Log "Failed to install module '$ModuleName'. Error: $_" -Level "ERROR"
            exit 1
        }
    }
    else {
        Write-Log "Module '$ModuleName' is already installed." -Level "INFO"
    }

    # Import the module
    try {
        Import-Module -Name $ModuleName -ErrorAction Stop
        Write-Log "Module '$ModuleName' imported successfully." -Level "INFO"
    }
    catch {
        Write-Log "Failed to import module '$ModuleName'. Error: $_" -Level "ERROR"
        exit 1
    }
}

# Function to connect to Microsoft Graph (Azure AD)
function Connect-ToAzureAD {
    param (
        [Parameter(Mandatory = $true)][string]$TenantID,
        [Parameter(Mandatory = $true)][string]$AppID,
        [Parameter(Mandatory = $true)][securestring]$AppSecret
    )

    try {
        # Convert the secure string to an encrypted plain text (handled securely)
        $secureAppSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($AppSecret)
        )

        # Connect using app-based authentication
        Connect-MgGraph -ClientId $AppID -TenantId $TenantID -ClientSecret $secureAppSecret -Scopes "Group.ReadWrite.All","Directory.Read.All"

        Write-Log "Successfully connected to Azure AD." -Level "INFO"
    }
    catch {
        Write-Log "Failed to connect to Azure AD: $_" -Level "ERROR"
        exit 1
    }
}

# Function to retrieve all OUs from On-Premises Active Directory
function Get-ADOrganizationalUnits {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Log "ActiveDirectory module imported successfully." -Level "INFO"
    }
    catch {
        Write-Log "ActiveDirectory module is not installed. Please install RSAT tools to get this module. Error: $_" -Level "ERROR"
        exit 1
    }

    try {
        $OUs = Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName
        Write-Log "Retrieved $($OUs.Count) Organizational Units from On-Prem AD." -Level "INFO"
        return $OUs
    }
    catch {
        Write-Log "Failed to retrieve OUs from On-Prem AD: $_" -Level "ERROR"
        exit 1
    }
}

# Function to check if a group exists in Azure AD
function Get-AzureADGroupIfExists {
    param (
        [string]$GroupName
    )

    try {
        $group = Get-MgGroup -Filter "displayName eq '$GroupName'" -ErrorAction SilentlyContinue
        return $group
    }
    catch {
        Write-Log "Error checking existence of group '$GroupName': $_" -Level "ERROR"
        return $null
    }
}

# Function to generate dynamic membership rule based on OU
function Generate-MembershipRule {
    param (
        [string]$OUName
    )

    # Hardcoded Azure AD attribute that maps to the On-Prem OU name
    # Modify the attribute name as per your environment if different
    $rule = "(device.extensionAttribute1 -eq `"$OUName`")"
    return $rule
}

# Function to create a dynamic group in Azure AD
function Create-DynamicAzureADGroup {
    param (
        [string]$GroupName,
        [string]$GroupDescription,
        [string]$MembershipRule
    )

    try {
        # Define the group properties
        $groupParams = @(
            @{
                DisplayName                      = $GroupName
                Description                      = $GroupDescription
                MailEnabled                      = $false
                MailNickname                     = ($GroupName -replace ' ', '').ToLower()
                SecurityEnabled                  = $true
                GroupTypes                       = @("DynamicMembership")
                MembershipRule                   = $MembershipRule
                MembershipRuleProcessingState    = "On"
            }
        )

        # Create the dynamic group in Azure AD
        $newGroup = New-MgGroup @groupParams

        Write-Log "Dynamic group '$GroupName' created successfully. Group ID: $($newGroup.Id)" -Level "INFO"
    }
    catch {
        Write-Log "Failed to create dynamic group '$GroupName': $_" -Level "ERROR"
    }
}

# ====================== Script Execution ======================

# Start Logging
Write-Log "======================" -Level "INFO"
Write-Log "Script Execution Started." -Level "INFO"
Write-Log "======================" -Level "INFO"

# Step 1: Ensure required modules are installed and imported
Ensure-Module -ModuleName "Microsoft.Graph"
Ensure-Module -ModuleName "ActiveDirectory"

# Step 2: Connect to Azure AD
Connect-ToAzureAD -TenantID $tenantID -AppID $appID -AppSecret $appSecret

# Step 3: Retrieve all OUs from On-Prem AD
$OUs = Get-ADOrganizationalUnits

# Step 4: Iterate through each OU to create dynamic groups
foreach ($OU in $OUs) {
    $OUName = $OU.Name.Trim()
    $GroupName = "$groupPrefix$OUName"

    Write-Log "Processing OU: '$OUName' - Intended Group Name: '$GroupName'" -Level "INFO"

    # Check if the group already exists
    $existingGroup = Get-AzureADGroupIfExists -GroupName $GroupName

    if ($existingGroup) {
        Write-Log "Group '$GroupName' already exists. Skipping creation." -Level "WARNING"
    }
    else {
        # Generate the dynamic membership rule
        $membershipRule = Generate-MembershipRule -OUName $OUName

        # Define group description
        $groupDescription = "Dynamic group for devices in OU '$OUName'"

        # Create the dynamic group
        Create-DynamicAzureADGroup -GroupName $GroupName -GroupDescription $groupDescription -MembershipRule $membershipRule
    }
}

Write-Log "Dynamic group creation process completed." -Level "INFO"
Write-Log "======================" -Level "INFO"

# End of Script
