<#
.SYNOPSIS
    Synchronizes members from an on-premises Active Directory group with a Microsoft Entra ID (Azure AD) group.

.DESCRIPTION
    - Retrieves all members of an on-premises AD group.
    - Retrieves all members of an Entra ID (Azure AD) group.
    - Compares both groups based on UserPrincipalName (UPN) case-insensitively.
    - Adds missing users from on-prem group to the Entra ID group.
    - Removes users from the Entra ID group if they are no longer present in the on-prem group.
    - Uses app-based authentication (Client Credentials Grant) with Microsoft Graph.

.REQUIREMENTS
    - RSAT Active Directory PowerShell module installed on the server/machine.
    - Microsoft.Graph PowerShell modules:
        - Microsoft.Graph.Users
        - Microsoft.Graph.Groups
        - Microsoft.Graph.Authentication
    - An Azure AD App Registration with the following permissions:
        - GroupMember.ReadWrite.All
        - User.Read.All
        - Group.Read.All

.AUTHOR
    Mohammed Omar
    Date: 2025-04-27
#>

# =======================
# Configuration Variables
# =======================

$OnPremGroupName    = "operations"                              # Name of the on-premises AD distribution/security group
$EntraGroupId       = "6cea5bde-f244-4bf5-b94b-0fa7f4cca633"     # Object ID of the Entra ID (Azure AD) group


$TenantId           = "c2b04da6-8487-41cc-8803-90321048a772"     # Tenant ID (Directory ID) of your Microsoft 365 tenant
$ClientId           = "6c70c0c3-e3a6-489c-973e-51e8138540f9"     # App registration Client ID
$ClientSecret       = "Uoj8Q~1_acd.7WU4Ol3vOczrfeYQbdHR_mzhTb6n" # App registration Client Secret

# =======================
# Load Required Modules
# =======================

# List of modules required by the script
$Modules = @(
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Authentication",
    "ActiveDirectory"
)

# Install and Import modules if necessary
foreach ($Module in $Modules) {
    if (-not (Get-Module -ListAvailable $Module)) {
        Write-Host "Installing $Module..." -ForegroundColor Cyan
        Install-Module $Module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Module -Force
}

# =======================
# Authenticate to Microsoft Graph
# =======================

Write-Host "`nAuthenticating to Microsoft Graph..." -ForegroundColor Cyan

# Create an OAuth 2.0 Token using client credentials
$TokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $TokenBody
    Connect-MgGraph -AccessToken (ConvertTo-SecureString $TokenResponse.access_token -AsPlainText -Force)
    Write-Host "Authenticated to Microsoft Graph successfully." -ForegroundColor Green
} catch {
    Write-Host "Graph authentication failed: $_" -ForegroundColor Red
    exit
}

# =======================
# Retrieve On-Prem AD Users
# =======================

Write-Host "`nFetching users from on-prem AD group '$OnPremGroupName'..." -ForegroundColor Cyan

try {
    # Get all members from on-premises AD group and store their UPNs
    $OnPremUPNs = Get-ADGroupMember -Identity $OnPremGroupName -Recursive |
        Where-Object { $_.ObjectClass -eq "user" } |
        ForEach-Object { (Get-ADUser $_.SamAccountName -Properties UserPrincipalName).UserPrincipalName.Trim().ToLower() }

    Write-Host "Retrieved $($OnPremUPNs.Count) on-prem AD users." -ForegroundColor Green
} catch {
    Write-Host "Error retrieving on-prem users: $_" -ForegroundColor Red
    exit
}

# =======================
# Retrieve Entra ID Users
# =======================

Write-Host "`nFetching users from Entra ID group..." -ForegroundColor Cyan

$EntraMembers = @()

try {
    # Get all members of Entra group and store their UPNs
    $CloudMembers = Get-MgGroupMember -GroupId $EntraGroupId -All
    foreach ($member in $CloudMembers) {
        if ($member.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') {
            $user = Get-MgUser -UserId $member.Id
            $EntraMembers += [PSCustomObject]@{
                UPN = $user.UserPrincipalName.ToLower()
                Id  = $user.Id
            }
        }
    }
    Write-Host "Retrieved $($EntraMembers.Count) Entra ID users." -ForegroundColor Green
} catch {
    Write-Host "Error retrieving Entra ID users: $_" -ForegroundColor Red
    exit
}

# =======================
# Add Users Missing in Entra ID
# =======================

Write-Host "`nAdding missing users to Entra ID group..." -ForegroundColor Cyan

# Compare: Find users present in On-Prem but missing in Entra
$EntraUPNs = $EntraMembers.UPN
$UsersToAdd = $OnPremUPNs | Where-Object { $EntraUPNs -notcontains $_ }

foreach ($upn in $UsersToAdd) {
    try {
        # Find user in Entra
        $user = Get-MgUser -Filter "userPrincipalName eq '$upn'"
        if ($user) {
            # Add user to the group
            New-MgGroupMember -GroupId $EntraGroupId -DirectoryObjectId $user.Id
            Write-Host "Added: $upn" -ForegroundColor Green
        } else {
            Write-Host "User not found in Entra ID: $upn" -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "Error adding user $upn : $_"
    }
}

if ($UsersToAdd.Count -eq 0) {
    Write-Host "No users to add. All users already exist in Entra ID group." -ForegroundColor Green
}

# =======================
# Remove Users Not in On-Prem AD
# =======================

Write-Host "`nRemoving users not in on-prem AD group from Entra ID group..." -ForegroundColor Cyan

Import-Module Microsoft.Graph.Groups -Force

# Compare: Find users present in Entra but missing in On-Prem
$UsersToRemove = $EntraMembers | Where-Object { $OnPremUPNs -notcontains $_.UPN }

foreach ($user in $UsersToRemove) {
    try {
        # Remove the user from the Entra group
        Remove-MgGroupMemberByRef -GroupId $EntraGroupId -DirectoryObjectId $user.Id -Confirm:$false
        Write-Host "Removed: $($user.UPN)" -ForegroundColor Yellow
    } catch {
        Write-Warning "Error removing user $($user.UPN): $_"
    }
}

if ($UsersToRemove.Count -eq 0) {
    Write-Host "No users to remove. Groups are fully synchronized." -ForegroundColor Green
}

# =======================
# Disconnect from Microsoft Graph
# =======================

Disconnect-MgGraph | Out-Null
Write-Host "`nSynchronization completed successfully." -ForegroundColor Cyan
