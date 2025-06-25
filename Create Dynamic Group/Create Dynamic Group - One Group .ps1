# ========================================================================================
# Script to Create Dynamic Azure AD Group Based on Device Names
# ========================================================================================
# Description:
# This script connects to Microsoft Graph using a service principal, creates a dynamic
# Azure Active Directory (AAD) group where membership is based on the device display name
# containing a specific string (e.g., 'it-op'). The script includes error handling,
# secure credential management, and formatted output.
# ========================================================================================

# =====================
# Module Installation Check
# =====================

# Define the submodules required for the script
$requiredModules = @("Microsoft.Graph.Groups", "Microsoft.Graph.Authentication")

# Check if the submodules are installed
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "The required module '$module' is not installed. Installing it now..." -ForegroundColor Yellow
        
        try {
            # Install the submodule
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
            Write-Host "Module '$module' installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install the required module '$module'. Please check your internet connection or permissions and try again."
            exit
        }
    } else {
        Write-Host "Module '$module' is already installed." -ForegroundColor Green
    }
}

# Import only the necessary Microsoft Graph submodules
Import-Module Microsoft.Graph.Groups -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# =====================
# Configuration Section
# =====================

# Azure AD Tenant ID (replace with your actual Tenant ID)
$tenantID = "xxxxxxxx"

# Azure AD App Registration Client ID (replace with your actual Client ID)
$appID = "xxxxxxxxxxx"

# Azure AD App Registration Client Secret (replace with your actual Client Secret)
$appSecretPlain = "xxxxxxxxx"
$appSecret = ConvertTo-SecureString $appSecretPlain -AsPlainText -Force

# Group Properties
$groupName = "IT-Operations Devices"                            # Name of the new dynamic group
$groupDescription = "Dynamic group for devices with 'it-op' in the display name" # Description of the group
$mailNickname = $groupName.Replace(" ", "").ToLower()           # Mail nickname for the group

# Dynamic Membership Rule for devices whose displayName contains 'it-op'
$membershipRule = "(device.displayName -contains 'it-op')"      # Adjusted rule for devices

# ======================
# Function Definitions
# ======================

# Function to connect to Microsoft Graph
function Connect-ToGraph {
    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

        # Connect to Microsoft Graph using the service principal credentials
        Connect-MgGraph -ClientId $appID -TenantId $tenantID -ClientSecret (ConvertFrom-SecureString $appSecret) -ErrorAction Stop

        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
    }
    catch {
        # Catch and report any errors during connection
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit
    }
}

# Function to create a dynamic group in Azure AD
function Create-DynamicGroup {
    try {
        Write-Host "Creating dynamic group '$groupName'..." -ForegroundColor Cyan

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
        Write-Host "Dynamic group created successfully!" -ForegroundColor Green
        Write-Host "-----------------------------------"
        Write-Host "Group Name:        $($newGroup.DisplayName)"
        Write-Host "Description:       $($newGroup.Description)"
        Write-Host "Object ID:         $($newGroup.Id)"
        Write-Host "Membership Rule:   $membershipRule"
        Write-Host "-----------------------------------"
    }
    catch {
        # Catch and report any errors during group creation
        Write-Error "Failed to create dynamic group: $_"
    }
}

# ======================
# Script Execution
# ======================

# Step 1: Connect to Microsoft Graph
Connect-ToGraph

# Step 2: Create the dynamic group with the specified rules
Create-DynamicGroup

# End of script
