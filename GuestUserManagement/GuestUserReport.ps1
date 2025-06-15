<#

.SYNOPSIS
    Export Office 365 Guest Users and Their Group Memberships using Microsoft Graph

.DESCRIPTION
    This PowerShell script connects to Microsoft Graph using the Microsoft.Graph module to retrieve
    all guest users in the Microsoft 365 tenant. It then collects key attributes such as display name,
    UPN, email, creation date, creation type, company (guessed from domain if not provided), and
    invitation status. Additionally, it fetches the group memberships for each guest user.

    Optional parameters allow filtering guests based on account age, such as listing only recently
    added or long-standing guests.

    The result is exported to a CSV file in the C:\Temp directory, with a timestamp in the filename.
    After exporting, the script prompts the user to open the CSV.

.PARAMETER StaleGuests
    Optional. Only include guests whose accounts are older than this many days.

.PARAMETER RecentlyCreatedGuests
    Optional. Only include guests whose accounts are newer than this many days.

.OUTPUTS
    CSV file saved to C:\Temp\GuestUserReport_<timestamp>.csv

.EXAMPLE
    .\GuestUserReport.ps1
    Runs the script and exports all guest users.

    .\GuestUserReport.ps1 -StaleGuests 180
    Exports only guest users older than 180 days.

    .\GuestUserReport.ps1 -RecentlyCreatedGuests 30
    Exports only guest users created in the last 30 days.


.NOTES
    Author  : Mohammad Abdulkader Omar
    Version : 3.3
    Website : https://momar.tech
    Date    : 2025-06-15

#>

Param (
    [Parameter(Mandatory = $false)]
    [int]$StaleGuests,                # Only include users older than this many days
    [int]$RecentlyCreatedGuests       # Only include users newer than this many days
)

# Ensure C:\Temp exists
if (-not (Test-Path "C:\Temp")) {
    New-Item -Path "C:\" -Name "Temp" -ItemType Directory | Out-Null
}

# Ensure Microsoft.Graph is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft.Graph module not found. Installing..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
    Import-Module Microsoft.Graph
}

# Connect to Graph

Connect-MgGraph
Write-Host "✅ Connected to Microsoft Graph" -ForegroundColor Green

# Output file path
$Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ExportCSV = "C:\Temp\GuestUserReport_$Timestamp.csv"
$Results = @()
$Counter = 0

# Fetch guest users
$Guests = Get-MgUser -All -Filter "UserType eq 'Guest'" `
    -ExpandProperty MemberOf `
    -Property DisplayName,UserPrincipalName,Mail,CompanyName,CreatedDateTime,CreationType,ExternalUserState

foreach ($User in $Guests) {
    $Counter++
    Write-Progress -Activity "Exporting Guest Users" -Status "Processing $($User.DisplayName)" -PercentComplete (($Counter / $Guests.Count) * 100)

    # Basic info
    $DisplayName = $User.DisplayName
    $UserPrincipalName = $User.UserPrincipalName
    $Email = $User.Mail

    # Company: prefer real, else extract from domain
    if ($User.CompanyName) {
        $Company = $User.CompanyName
    } elseif ($User.Mail -and $User.Mail -match "@(.+?)\.") {
        $Domain = $Matches[1].ToLower()
        switch ($Domain) {
            "gmail"   { $Company = "Gmail" }
            "hotmail" { $Company = "Hotmail" }
            "outlook" { $Company = "Outlook" }
            "yahoo"   { $Company = "Yahoo" }
            "icloud"  { $Company = "iCloud" }
            default   { $Company = ($Domain.Substring(0,1).ToUpper() + $Domain.Substring(1)) }
        }
    } else {
        $Company = "-"
    }

    # Account age and creation
    if ($User.CreatedDateTime) {
        $CreationTime = $User.CreatedDateTime
        $AccountAge = (New-TimeSpan -Start $User.CreatedDateTime).Days
    } else {
        $CreationTime = "Unknown"
        $AccountAge = "Unknown"
    }

    # Skip filtered guests
    if ($AccountAge -ne "Unknown") {
        if ($StaleGuests -and ($AccountAge -lt $StaleGuests)) { continue }
        if ($RecentlyCreatedGuests -and ($AccountAge -gt $RecentlyCreatedGuests)) { continue }
    }

    # Creation type and invitation status
    $CreationType = if ($User.CreationType) { $User.CreationType } else { "-" }
    $InvitationAccepted = if ($User.ExternalUserState) { $User.ExternalUserState } else { "-" }

    # Group memberships
    $GroupMemberships = "-"
    if ($User.MemberOf) {
        $GroupNames = @()
        foreach ($Group in $User.MemberOf) {
            if ($Group.AdditionalProperties["displayName"]) {
                $GroupNames += $Group.AdditionalProperties["displayName"]
            }
        }
        if ($GroupNames.Count -gt 0) {
            $GroupMemberships = $GroupNames -join ", "
        }
    }

    # Add result row
    $Results += [PSCustomObject]@{
        DisplayName         = $DisplayName
        UserPrincipalName   = $UserPrincipalName
        Company             = $Company
        EmailAddress        = $Email
        CreationTime        = $CreationTime
        "AccountAge(days)"  = $AccountAge
        CreationType        = $CreationType
        InvitationAccepted  = $InvitationAccepted
        GroupMembership     = $GroupMemberships
    }
}

# Export to CSV
if ($Results.Count -gt 0) {
    $Results | Export-Csv -Path $ExportCSV -NoTypeInformation -Encoding UTF8
    Write-Host "`n✅ Export completed successfully!" -ForegroundColor Green
    Write-Host "📄 File saved to: $ExportCSV"
    Write-Host "👥 Guest users exported: $($Results.Count)"

    $Prompt = New-Object -ComObject wscript.shell
    if ($Prompt.popup("Open the output file?", 0, "Export Complete", 4) -eq 6) {
        Invoke-Item $ExportCSV
    }
} else {
    Write-Host "❌ No guest users found matching filters." -ForegroundColor Yellow
}

# Disconnect
Disconnect-MgGraph | Out-Null
