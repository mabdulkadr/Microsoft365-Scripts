# ========================================================================================
# Enhanced Script to Create Dynamic Azure AD Groups Based on On-Premises OU Names
# ========================================================================================
# Description:
# This script connects to both On-Premises Active Directory and Azure Active Directory (AAD),
# retrieves all OU names from On-Prem AD, and creates corresponding dynamic groups in AAD
# with a prefix "Devices-". It ensures that groups are only created if they do not already exist.
# The script includes robust error handling, logging, and automatic module installation.
# ========================================================================================

# Azure AD Tenant ID (replace with your actual Tenant ID)
$tenantID = "c2b04da6-8487-41cc-8803-90321048a772"

# Azure AD App Registration Client ID (replace with your actual Client ID)
$appID = "6c70c0c3-e3a6-489c-973e-51e8138540f9"

# Azure AD App Registration Client Secret (replace with your actual Client Secret as plain text)
$appSecret = "Uoj8Q~1_acd.7WU4Ol3vOczrfeYQbdHR_mzhTb6n"

# Prefix for dynamic group names
$groupPrefix = "Devices - "

# ======================
# Logging Configuration
# ======================

# Path to the log file
$logFilePath = "C:\Scripts\Logs\CreateDynamicGroups.log"

# Ensure the log directory exists
$logDirectory = Split-Path -Path $logFilePath -Parent
if (-not (Test-Path -Path $logDirectory)) {
    try {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Host "Failed to create log directory at '$logDirectory'. Error: $_" -ForegroundColor Red
        exit 1
    }
}

# Function to log messages with timestamps
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

# ======================
# Function Definitions
# ======================

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

# ======================
# Script Execution
# ======================

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
