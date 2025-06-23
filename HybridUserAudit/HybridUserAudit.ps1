<#
.SYNOPSIS
    Generates a full hybrid user audit report from Active Directory and Microsoft Entra ID.

.DESCRIPTION
    This PowerShell script collects all user accounts from both on-premises Active Directory and Microsoft Entra ID (formerly Azure AD),
    matches them by username, and generates a merged audit report. For each user, it extracts attributes such as display name, email,
    department, job title, creation date, last logon (from AD and Entra), password last set, and account status.

    The script:
    - Shows real-time progress while processing each user.
    - Exports the results to a UTF-8 encoded CSV file under C:\Reports.
    - Logs all output to a .txt file for auditing.
    - Automatically creates the C:\Reports folder if it does not exist.
    - Is optimized to run in PowerShell ISE 5.1.

.EXAMPLE
    Run from PowerShell ISE:
    PS> .\HybridUserAudit.ps1

.NOTES
    Author  : Mohammad Abdulkader Omar
    Website : https://momar.tech
    Date    : 2025-06-23
#>

# ===============================
# Load Required PowerShell Modules
# ===============================
Import-Module ActiveDirectory -ErrorAction Stop               # For querying local AD users
Import-Module Microsoft.Graph.Users -ErrorAction Stop        # For querying Microsoft Entra ID users

# ===============================
# Connect to Microsoft Graph
# ===============================
Write-Host "🔄 Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome
Write-Host "✅ Connected to Microsoft Graph.`n" -ForegroundColor Green

# ===============================
# Setup Export Paths
# ===============================
$timestamp  = Get-Date -Format "yyyy-MM-dd_HH-mm"
$reportPath = "C:\Reports"
if (-not (Test-Path $reportPath)) {
    New-Item -Path $reportPath -ItemType Directory -Force | Out-Null
}

$csvPath = "$reportPath\FullUserReport_$timestamp.csv"
$logPath = "$reportPath\HybridUserAuditLog_$timestamp.txt"

# ===============================
# Define CSV Column Order
# ===============================
$columns = @(
    'Username','DisplayName','Department','Title','Email',
    'InAD','AD_Enabled','AD_Created','AD_LastLogon','AD_WhenChanged','AD_PwdLastSet','AD_Description','AD_DistinguishedName',
    'InEntraID','Entra_Enabled','Entra_Created','Entra_LastInteractiveSignIn','Entra_LastNonInteractiveSignIn'
)

# ===============================
# Initialize CSV with Headers
# ===============================
@() | Select-Object $columns | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# ===============================
# Start Logging Output
# ===============================
Start-Transcript -Path $logPath -Append

# ===============================
# Fetch Users from Microsoft Entra ID
# ===============================
Write-Host "🔎 Fetching Entra ID users..." -ForegroundColor Yellow
$entraUsers = @{}
$j = 0
try {
    Get-MgUser -All -Property DisplayName, UserPrincipalName, Department, JobTitle, Mail, AccountEnabled, CreatedDateTime, SignInActivity |
    ForEach-Object {
        $username = ($_.UserPrincipalName -split "@")[0].ToLower()
        $entraUsers[$username] = $_
        $j++
        Write-Host "[EntraID] $j - $username : $($_.DisplayName)" -ForegroundColor DarkYellow
    }
    Write-Host "`n✅ Total Entra ID users: $j" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to load Entra ID users: $($_.Exception.Message)" -ForegroundColor Red
}

# ===============================
# Fetch AD Users and Generate Merged Report
# ===============================
Write-Host "`n🔎 Fetching AD users..." -ForegroundColor Yellow
$i = 0
$prefixes = @('a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','0','1','2','3','4','5','6','7','8','9')
foreach ($prefix in $prefixes) {
    Write-Host "🔍 Fetching AD users starting with '$prefix'..." -ForegroundColor Yellow
    Get-ADUser -LDAPFilter "(sAMAccountName=$prefix*)" -Properties * | ForEach-Object {
        $i++
        $adUser = $_
        $username = $adUser.SamAccountName.ToLower()
        $entraUser = $entraLookup[$username]

        $record = [PSCustomObject]@{
            Username                        = $username
            InAD                            = "Yes"
            InEntraID                       = if ($entraUser) { "Yes" } else { "No" }
            DisplayName                     = if ($adUser.DisplayName) { $adUser.DisplayName } elseif ($entraUser) { $entraUser.DisplayName } else { "" }
            Department                      = if ($adUser.Department) { $adUser.Department } elseif ($entraUser) { $entraUser.Department } else { "" }
            Title                           = if ($adUser.Title) { $adUser.Title } elseif ($entraUser) { $entraUser.JobTitle } else { "" }
            Email                           = if ($adUser.Mail) { $adUser.Mail } elseif ($entraUser) { $entraUser.Mail } else { "" }
            AD_Enabled                      = if ($adUser.Enabled) { 'Enabled' } else { 'Disabled' }
            Entra_Enabled                   = if ($entraUser.AccountEnabled) { 'Enabled' } else { 'Disabled' }
            AD_Created                      = $adUser.WhenCreated.ToString("yyyy-MM-dd")
            Entra_Created                   = if ($entraUser.CreatedDateTime) { $entraUser.CreatedDateTime.ToString("yyyy-MM-dd") } else { "" }
            AD_LastLogon                    = if ($adUser.LastLogonDate) { $adUser.LastLogonDate.ToString("yyyy-MM-dd") } else { "" }
            Entra_LastInteractiveSignIn     = if ($entraUser.SignInActivity.LastSignInDateTime) { $entraUser.SignInActivity.LastSignInDateTime.ToString("yyyy-MM-dd") } else { "" }
            Entra_LastNonInteractiveSignIn  = if ($entraUser.SignInActivity.LastNonInteractiveSignInDateTime) { $entraUser.SignInActivity.LastNonInteractiveSignInDateTime.ToString("yyyy-MM-dd") } else { "" }
            AD_WhenChanged                  = if ($adUser.whenChanged) { $adUser.whenChanged.ToString("yyyy-MM-dd") } else { "" }
            AD_PwdLastSet                   = if ($adUser.pwdLastSet) { ([datetime]::FromFileTime($adUser.pwdLastSet)).ToString("yyyy-MM-dd") } else { "" }
            AD_Description                  = $adUser.Description
            AD_DistinguishedName            = $adUser.DistinguishedName
        }

        $record | Select-Object $columns | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Append
        Write-Host "[✔] $i - $username : $($record.DisplayName)" -ForegroundColor Magenta
    }
}

# ===============================
# Wrap-Up
# ===============================
Stop-Transcript
Write-Host "`n✅ Report saved to: $csvPath" -ForegroundColor Green
Start-Process "explorer.exe" -ArgumentList (Split-Path $csvPath)
