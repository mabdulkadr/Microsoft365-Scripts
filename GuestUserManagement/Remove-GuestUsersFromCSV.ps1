<#
.SYNOPSIS
    Bulk deletes guest users from Microsoft Entra ID based on a provided CSV list.

.DESCRIPTION
    This script connects to Microsoft Graph and processes a CSV file containing user identifiers
    (such as UserPrincipalName or Email). For each entry, the script attempts to find the corresponding
    account in Microsoft Entra ID, fetches its UserType property, and deletes the user only if their 
    UserType is "Guest". The script prints the detected UserType for each user, logs actions, and
    provides a summary report of deleted, skipped, not found, and failed entries.

.PARAMETER CSV File
    The input CSV file must include a column named "UserPrincipalName" or "Email". The script prompts
    you to select the file interactively.

.EXAMPLE
    # Run the script and select the CSV file when prompted.
    PS> .\Remove-GuestUsersFromCSV.ps1

.NOTES
    Author  : Mohammad Abdulkader Omar
    Website : https://momar.tech
    Date    : 2025-06-15
    Version : 2.0

#>

# Connect to Microsoft Graph
Connect-MgGraph

# Prompt for CSV file
Add-Type -AssemblyName System.Windows.Forms
$OpenDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenDialog.Filter = "CSV Files (*.csv)|*.csv"
$OpenDialog.Title  = "Select CSV with UserPrincipalName column"
if ($OpenDialog.ShowDialog() -ne 'OK') {
    Write-Host "❌ No file selected. Exiting." -ForegroundColor Red
    exit
}
$InputCsv = $OpenDialog.FileName

# Import CSV
$List = Import-Csv $InputCsv

# Detect column
if ($List[0].PSObject.Properties.Name -contains 'UserPrincipalName') {
    $Col = 'UserPrincipalName'
} elseif ($List[0].PSObject.Properties.Name -contains 'Email') {
    $Col = 'Email'
} else {
    Write-Host "❌ No UserPrincipalName or Email column found." -ForegroundColor Red
    exit
}

# Start processing
$Total = 0; $Deleted = 0; $Skipped = 0; $NotFound = 0; $Failed = 0
foreach ($row in $List) {
    $upn = $row.$Col
    $Total++
    try {
        # Always request UserType property!
        $user = Get-MgUser -Filter "UserPrincipalName eq '$upn'" -Property Id,UserType,DisplayName,UserPrincipalName

        if ($user) {
            $realType = $user.UserType
            if ($realType) { $realType = $realType.Trim().ToLower() } else { $realType = "<empty>" }
            Write-Host "[$upn] UserType (from Graph): $realType"

            if ($realType -eq "guest") {
                Remove-MgUser -UserId $user.Id -Confirm:$false
                Write-Host "✅ Deleted guest: $upn" -ForegroundColor Green
                $Deleted++
            } else {
                Write-Host "⏩ Skipped (not guest): $upn (UserType returned: $realType)" -ForegroundColor Cyan
                $Skipped++
            }
        } else {
            Write-Host "⚠️ Not found: $upn" -ForegroundColor Yellow
            $NotFound++
        }
    } catch {
        Write-Host "❌ Failed to process: $upn | $_" -ForegroundColor Red
        $Failed++
    }
}

Write-Host "`nSummary:"
Write-Host "Total processed : $Total"
Write-Host "Deleted guests  : $Deleted"
Write-Host "Skipped (not guest): $Skipped"
Write-Host "Not found       : $NotFound"
Write-Host "Failed          : $Failed"
