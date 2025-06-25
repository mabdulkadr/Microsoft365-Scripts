# ========================================================================================
# Script to Create Multiple Dynamic Azure AD Groups Based on Names from a CSV File
# ========================================================================================
# Description:
# This script connects to Microsoft Graph using either app-based or user-based authentication,
# reads group details from a CSV file, and creates dynamic Azure Active Directory (AAD) groups
# based on the device display name and other criteria specified in the CSV.
# ========================================================================================

# =====================
# Configuration Section
# =====================

# Azure AD Tenant ID (replace with your actual Tenant ID)
$tenantID = "c2b04da6-8487-41cc-8803-90321048a772"

# Azure AD App Registration Client ID (replace with your actual Client ID)
$appID = "6c70c0c3-e3a6-489c-973e-51e8138540f9"

# Azure AD App Registration Client Secret (replace with your actual Client Secret as plain text)
$appSecretPlain = "Uoj8Q~1_acd.7WU4Ol3vOczrfeYQbdHR_mzhTb6n"

# Permission scopes for user-based authentication (optional if not using app-based auth)
$scopes = "https://graph.microsoft.com/.default"

# Path to the CSV file containing group information
$csvFilePath = "C:\Groups.csv"  # Update with your actual path

# ======================
# Function Definitions
# ======================

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
                    Write-Host "Version 2 module detected" -ForegroundColor Green
                    $accessTokenFinal = ConvertTo-SecureString -String $accessToken -AsPlainText -Force
                } else {
                    Write-Host "Version 1 module detected" -ForegroundColor Yellow
                    Select-MgProfile -Name Beta
                    $accessTokenFinal = $accessToken
                }
                $graph = Connect-MgGraph -AccessToken $accessTokenFinal
                Write-Host "Connected to Intune tenant $Tenant using app-based authentication" -ForegroundColor Green
            } else {
                # User-based authentication
                if ($version -eq 2) {
                    Write-Host "Version 2 module detected" -ForegroundColor Green
                } else {
                    Write-Host "Version 1 module detected" -ForegroundColor Yellow
                    Select-MgProfile -Name Beta
                }
                $graph = Connect-MgGraph -Scopes $Scopes
                Write-Host "Connected to Intune tenant $($graph.TenantId)" -ForegroundColor Green
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

        Write-Host "Creating dynamic group '$GroupName'..." -ForegroundColor Cyan

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
        Write-Host ""
        Write-Host "Dynamic group '$GroupName' created successfully!" -ForegroundColor Green
        Write-Host "-----------------------------------"
        Write-Host "Group Name:        $($newGroup.DisplayName)"
        Write-Host "Description:       $($newGroup.Description)"
        Write-Host "Object ID:         $($newGroup.Id)"
        Write-Host "Membership Rule:   $MembershipRule"
        Write-Host "-----------------------------------"
        Write-Host ""
    } catch {
        Write-Error "Failed to create dynamic group '$GroupName': $_"
    }
}

# ======================
# Script Execution
# ======================

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
