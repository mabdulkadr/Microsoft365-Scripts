<#
.SYNOPSIS
    Unified Active Directory Reports Script with Enhanced Console Menu & HTML Styling

.DESCRIPTION
    Combines six AD reporting scripts with a stylish, robust menu.
    Features centralized variables, reusable CSS, error handling, timestamped outputs,
    and the option to open generated reports in your default browser.

.NOTES
    Author: Combined and Enhanced by ChatGPT (2025-06-16)
    Original scripts by: Mohammed Omar
#>

# ========== GLOBAL SETTINGS ==========
$ReportPath = "C:\ADReports"
$NowString = Get-Date -Format "yyyy-MM-dd_HH-mm"
$CSS = @"
<style>
body { font-family: Arial, sans-serif; margin: 20px; background-color: #f4f4f4; color: #333; }
h1, h2, h3 { color: #0056b3; }
table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 20px;
    box-shadow: 0 2px 3px rgba(0,0,0,0.1);
    background-color: #fff;
}
th, td {
    border: 1px solid #ddd;
    padding: 8px;
    text-align: left;
}
th {
    background-color: #007bff;
    color: white;
    text-transform: uppercase;
}
tr:nth-child(even) { background-color: #f2f2f2; }
tr:hover { background-color: #e0e0e0; }
.header {
    background-color: #007bff;
    color: white;
    padding: 10px;
    text-align: center;
    border-radius: 5px;
    margin-bottom: 20px;
}
.report-section {
    margin-bottom: 40px;
    padding: 20px;
    border: 1px solid #ddd;
    border-radius: 5px;
    background-color: #fff;
}
</style>
"@

# Ensure AD module and output folder
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Host "ActiveDirectory module is required. Please install RSAT and try again." -ForegroundColor Red
    exit
}
If (-not (Test-Path $ReportPath)) { New-Item -Path $ReportPath -ItemType Directory | Out-Null }

# ========== UTILITIES ==========
function Write-Header($text) {
    Write-Host ""
    Write-Host ("=" * 65) -ForegroundColor Cyan
    Write-Host (" " + $text) -ForegroundColor Yellow
    Write-Host ("=" * 65) -ForegroundColor Cyan
    Write-Host ""
}
function Build-HtmlReport($title, $fragment, $precontent) {
    return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>$title</title>
    $CSS
</head>
<body>
    <div class="header">
        <h1>$title</h1>
        <p>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>
    <div class="report-section">
        $precontent
        $fragment
    </div>
</body>
</html>
"@
}
function Prompt-OpenReport($path) {
    Write-Host ""
    Write-Host "Press [O] to open the last report, any other key to return..." -ForegroundColor DarkGray
    $k = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if ($k.Character -eq 'O' -or $k.Character -eq 'o') {
        Start-Process $path
    }
}

# ========== REPORT FUNCTIONS ==========

function Show-CompleteComputerObjectReport {
    Write-Header "Complete Computer Object Report"
    $ReportFile = "$ReportPath\AD_CompleteComputerReport_$NowString.html"
    try {
        $Computers = Get-ADComputer -Filter * -Properties * | Sort-Object Name
        $ComputerData = $Computers | Select-Object Name, DNSHostName, IPv4Address, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, LastLogonDate, Enabled, Description, @{Name='OU';Expression={$_.DistinguishedName -replace '^CN=.*?,(OU=.*?|DC=.*?)$','$1'}}
        $frag = $ComputerData | ConvertTo-Html -Fragment
        $html = Build-HtmlReport "Complete AD Computer Report" $frag "<h2>All Computer Objects</h2>"
        $html | Out-File $ReportFile -Encoding UTF8
        Write-Host "Complete Computer Report generated: $ReportFile" -ForegroundColor Green
        Prompt-OpenReport $ReportFile
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-DomainControllersReport {
    Write-Header "Domain Controllers Report"
    $ReportFile = "$ReportPath\AD_DomainControllersReport_$NowString.html"
    try {
        $DCs = Get-ADDomainController -Filter * | Select-Object Name, HostName, IPv4Address, Site, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, IsGlobalCatalog, IsReadOnlyDC, Enabled | Sort-Object Name
        $frag = $DCs | ConvertTo-Html -Fragment
        $html = Build-HtmlReport "AD Domain Controllers Report" $frag "<h2>Domain Controllers</h2>"
        $html | Out-File $ReportFile -Encoding UTF8
        Write-Host "Domain Controllers Report generated: $ReportFile" -ForegroundColor Green
        Prompt-OpenReport $ReportFile
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-WorkstationsReport {
    Write-Header "Workstations Report"
    $ReportFile = "$ReportPath\AD_WorkstationsReport_$NowString.html"
    try {
        $Workstations = Get-ADComputer -Filter {OperatingSystem -NotLike "*Server*"} -Properties * | Sort-Object Name
        $WorkstationData = $Workstations | Select-Object Name, DNSHostName, IPv4Address, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, LastLogonDate, Enabled, Description, @{Name='OU';Expression={$_.DistinguishedName -replace '^CN=.*?,(OU=.*?|DC=.*?)$','$1'}}
        $frag = $WorkstationData | ConvertTo-Html -Fragment
        $html = Build-HtmlReport "AD Workstations Report" $frag "<h2>Workstations</h2>"
        $html | Out-File $ReportFile -Encoding UTF8
        Write-Host "Workstations Report generated: $ReportFile" -ForegroundColor Green
        Prompt-OpenReport $ReportFile
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-ServersReport {
    Write-Header "Servers Report"
    $ReportFile = "$ReportPath\AD_ServersReport_$NowString.html"
    try {
        $Servers = Get-ADComputer -Filter {OperatingSystem -Like "*Server*"} -Properties * | Sort-Object Name
        $ServerData = $Servers | Select-Object Name, DNSHostName, IPv4Address, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, LastLogonDate, Enabled, Description, @{Name='OU';Expression={$_.DistinguishedName -replace '^CN=.*?,(OU=.*?|DC=.*?)$','$1'}}
        $frag = $ServerData | ConvertTo-Html -Fragment
        $html = Build-HtmlReport "AD Servers Report" $frag "<h2>Servers</h2>"
        $html | Out-File $ReportFile -Encoding UTF8
        Write-Host "Servers Report generated: $ReportFile" -ForegroundColor Green
        Prompt-OpenReport $ReportFile
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-ComputerAccountStatusReport {
    Write-Header "Computer Account Status Report"
    $ReportFile = "$ReportPath\AD_ComputerAccountStatusReport_$NowString.html"
    try {
        $Computers = Get-ADComputer -Filter * -Properties Name, Enabled, LastLogonDate, DNSHostName | Sort-Object Name
        $AccountStatusData = $Computers | Select-Object Name, DNSHostName, Enabled, LastLogonDate, @{Name='AccountStatus'; Expression={if ($_.Enabled) {'Enabled'} else {'Disabled'}}}
        $frag = $AccountStatusData | ConvertTo-Html -Fragment
        $html = Build-HtmlReport "AD Computer Account Status Report" $frag "<h2>Computer Account Status</h2>"
        $html | Out-File $ReportFile -Encoding UTF8
        Write-Host "Computer Account Status Report generated: $ReportFile" -ForegroundColor Green
        Prompt-OpenReport $ReportFile
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-OSBasedReports {
    Write-Header "OS-Based Computer Report"
    $ReportFile = "$ReportPath\AD_OSBasedReport_$NowString.html"
    try {
        $Computers = Get-ADComputer -Filter * -Properties Name, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, LastLogonDate, DNSHostName, Enabled, Description | Sort-Object OperatingSystem, Name
        $OSGroups = $Computers | Group-Object OperatingSystem | Sort-Object Name

        $OSBasedHTML = ""
        foreach ($OSGroup in $OSGroups) {
            $OSName = $OSGroup.Name
            $OSCount = $OSGroup.Count
            $OSBasedHTML += "<h3>$OSName ($OSCount computers)</h3>"
            $OSGroup.Group | Select-Object Name, DNSHostName, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, LastLogonDate, Enabled, Description | ConvertTo-Html -Fragment | Out-String | ForEach-Object { $OSBasedHTML += $_ }
        }
        $html = Build-HtmlReport "AD OS-Based Computer Report" $OSBasedHTML "<h2>Computers by Operating System</h2>"
        $html | Out-File $ReportFile -Encoding UTF8
        Write-Host "OS-Based Computer Report generated: $ReportFile" -ForegroundColor Green
        Prompt-OpenReport $ReportFile
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Run-AllReports {
    Write-Header "Running ALL REPORTS"
    Show-CompleteComputerObjectReport
    Show-DomainControllersReport
    Show-WorkstationsReport
    Show-ServersReport
    Show-ComputerAccountStatusReport
    Show-OSBasedReports
    Write-Host "`nAll reports have been generated." -ForegroundColor Green
}

# ========== MAIN MENU ==========
function Show-Menu {
    Clear-Host
    Write-Host ""
    Write-Host ("=" * 65) -ForegroundColor Cyan
    Write-Host "                  Active Directory Reports - Main Menu                 " -ForegroundColor Yellow
    Write-Host ("=" * 65) -ForegroundColor Cyan
    Write-Host ""
    Write-Host " 1. Complete Computer Object Report"
    Write-Host " 2. Domain Controllers Report"
    Write-Host " 3. Workstations Report"
    Write-Host " 4. Servers Report"
    Write-Host " 5. Computer Account Status Report"
    Write-Host " 6. OS Based Reports"
    Write-Host " 7. Run ALL Reports"
    Write-Host " 0. Exit"
    Write-Host ""
}

do {
    Show-Menu
    $choice = Read-Host "Enter your choice (0-7)"
    switch ($choice) {
        '1' { Show-CompleteComputerObjectReport }
        '2' { Show-DomainControllersReport }
        '3' { Show-WorkstationsReport }
        '4' { Show-ServersReport }
        '5' { Show-ComputerAccountStatusReport }
        '6' { Show-OSBasedReports }
        '7' { Run-AllReports }
        '0' { Write-Host "Exiting. Goodbye!" -ForegroundColor Cyan }
        default { Write-Host "Invalid choice. Please enter a number from 0 to 7." -ForegroundColor Red }
    }
    if ($choice -ne '0') {
        Write-Host ""
        Write-Host "Press any key to return to the menu..." -ForegroundColor DarkGray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
} while ($choice -ne '0')
