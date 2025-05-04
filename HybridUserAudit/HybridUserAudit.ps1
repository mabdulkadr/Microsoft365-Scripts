<#
.SYNOPSIS
    Generates a full user report from Active Directory and Entra ID, with Arabic-safe CSV export and timestamped filename.

.DESCRIPTION
    This script fetches user details from on-premises Active Directory and Microsoft Entra ID (formerly Azure AD), merges them by username,
    and exports a unified CSV report to the user's Desktop with the current timestamp. It includes attributes like creation date,
    last logon date, password last set, and more.

    The output is formatted in UTF-8 with BOM to support Arabic characters. All date fields are formatted as yyyy-MM-dd.

.NOTES
    Author: Your Name
    Date: 2025-05-04
#>

# Load required modules
Import-Module ActiveDirectory -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

# Connect to Microsoft Graph
Write-Progress -Activity "Connecting to Microsoft Graph" -Status "Authenticating..." -PercentComplete 5
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome

# Get all AD users with additional properties
Write-Progress -Activity "Getting AD Users" -Status "Fetching from on-prem AD..." -PercentComplete 25
$adUsers = Get-ADUser -Filter * -Properties * | ForEach-Object {
    [PSCustomObject]@{
        Username             = $_.SamAccountName
        DisplayName          = $_.DisplayName
        Department           = $_.Department
        Title                = $_.Title
        Email                = $_.Mail
        AD_Enabled           = if ($_.Enabled) { 'Enabled' } else { 'Disabled' }
        AD_Created           = $_.WhenCreated
        AD_LastLogon         = $_.LastLogonDate
        AD_WhenChanged       = $_.whenChanged
        AD_PwdLastSet        = ([datetime]::FromFileTime($_.pwdLastSet))
        AD_Description       = $_.Description
        AD_DistinguishedName = $_.DistinguishedName
    }
}

# Get all Entra ID users with additional properties
Write-Progress -Activity "Getting Entra ID Users" -Status "Fetching from Microsoft Entra ID..." -PercentComplete 45
$entraUsers = Get-MgUser -All -Property DisplayName, UserPrincipalName, Department, JobTitle, Mail, AccountEnabled, CreatedDateTime, SignInActivity | ForEach-Object {
    [PSCustomObject]@{
        Username                        = $_.UserPrincipalName.Split("@")[0]
        DisplayName                     = $_.DisplayName
        Department                      = $_.Department
        Title                           = $_.JobTitle
        Email                           = $_.Mail
        Entra_Enabled                   = if ($_.AccountEnabled) { 'Enabled' } else { 'Disabled' }
        Entra_Created                   = $_.CreatedDateTime
        Entra_LastInteractiveSignIn     = $_.SignInActivity.LastSignInDateTime
        Entra_LastNonInteractiveSignIn  = $_.SignInActivity.LastNonInteractiveSignInDateTime
    }
}

# Merge users by username
Write-Progress -Activity "Merging Users" -Status "Combining AD + Entra ID data..." -PercentComplete 70
$merged = @()
$allUsernames = ($adUsers.Username + $entraUsers.Username) | Sort-Object -Unique
$count = 0
$total = $allUsernames.Count

foreach ($username in $allUsernames) {
    $count++
    Write-Progress -Activity "Merging User Data" -Status "$count of $total users" -PercentComplete ([math]::Round(($count / $total) * 25) + 70)

    $ad = $adUsers | Where-Object { $_.Username -eq $username }
    $entra = $entraUsers | Where-Object { $_.Username -eq $username }

    $merged += [PSCustomObject]@{
        Username                        = $username
        InAD                            = if ($ad) { "Yes" } else { "No" }
        InEntraID                       = if ($entra) { "Yes" } else { "No" }
        DisplayName                     = $ad.DisplayName ?? $entra.DisplayName
        Department                      = $ad.Department ?? $entra.Department
        Title                           = $ad.Title ?? $entra.Title
        Email                           = $ad.Email ?? $entra.Email
        AD_Enabled                      = $ad.AD_Enabled
        Entra_Enabled                   = $entra.Entra_Enabled
        AD_Created                      = $ad.AD_Created?.ToString("yyyy-MM-dd")
        Entra_Created                   = $entra.Entra_Created?.ToString("yyyy-MM-dd")
        AD_LastLogon                    = $ad.AD_LastLogon?.ToString("yyyy-MM-dd")
        Entra_LastInteractiveSignIn     = $entra.Entra_LastInteractiveSignIn?.ToString("yyyy-MM-dd")
        Entra_LastNonInteractiveSignIn  = $entra.Entra_LastNonInteractiveSignIn?.ToString("yyyy-MM-dd")
        AD_WhenChanged                  = $ad.AD_WhenChanged?.ToString("yyyy-MM-dd")
        AD_PwdLastSet                   = $ad.AD_PwdLastSet?.ToString("yyyy-MM-dd")
        AD_Description                  = $ad.AD_Description
        AD_DistinguishedName            = $ad.AD_DistinguishedName
    }
}

# Generate timestamped filename on Desktop
Write-Progress -Activity "Exporting Report" -Status "Saving to Desktop..." -PercentComplete 98
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
$desktopPath = [Environment]::GetFolderPath('Desktop')
$csvPath = Join-Path $desktopPath "FullUserReport-$timestamp.csv"
$utf8Bom = New-Object System.Text.UTF8Encoding $true

# Rearranged property order: AD fields first, then Entra fields
$orderedProps = @(
    'Username','DisplayName','Department','Title','Email',
    'InAD','AD_Enabled','AD_Created','AD_LastLogon','AD_WhenChanged','AD_PwdLastSet','AD_Description','AD_DistinguishedName',
    'InEntraID','Entra_Enabled','Entra_Created','Entra_LastInteractiveSignIn','Entra_LastNonInteractiveSignIn'
)

try {
    [System.IO.File]::WriteAllLines($csvPath, ($merged | Select-Object $orderedProps | ConvertTo-Csv -NoTypeInformation), $utf8Bom)
    Write-Host "`nReport saved to:`n$csvPath" -ForegroundColor Green
    # Step 6: Open folder
    Start-Process "explorer.exe" $desktopPath
} catch {
    Write-Host "Failed to save report. Error: $($_.Exception.Message)" -ForegroundColor Red
}