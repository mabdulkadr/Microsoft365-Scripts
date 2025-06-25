<#PSScriptInfo
 
.VERSION 0.2
 
.GUID c65d2f63-9958-427d-bf2c-0188c4739681
 
.AUTHOR Daniel Bradley
 
.COMPANYNAME Ourcloudnetwork.co.uk
 
.COPYRIGHT
 
.TAGS
    ourcloudnetwork
    Microsoft 365
    Microsoft Graph
 
.LICENSEURI
 
.PROJECTURI
 
.ICONURI
 
.EXTERNALMODULEDEPENDENCIES
    Microsoft.Graph.Authentication
 
.RELEASENOTES
    v0.1 - Initial release
    v0.2 - Remove dependency on modules. Add more subscriptions to the promo/free filter
#>

<#
.DESCRIPTION
 This script generates a license report for Microsoft 365
#> 

#Params
param(
     [Parameter(Mandatory)]
     [ValidateNotNullOrEmpty()]
     [string]$outpath
 )

# Check Microsoft Graph connection
$state = Get-MgContext

# Define required permissions properly as an array of strings
$requiredPerms = @(
    "User.Read.All",
    "AuditLog.Read.All",
    "Organization.Read.All",
    "RoleManagement.Read.Directory"
)

# Check if we're connected and have all required permissions
$hasAllPerms = $false
if ($state) {
    $missingPerms = @()
    foreach ($perm in $requiredPerms) {
        if ($state.Scopes -notcontains $perm) {
            $missingPerms += $perm
        }
    }
    
    if ($missingPerms.Count -eq 0) {
        $hasAllPerms = $true
        Write-Host "Connected to Microsoft Graph with all required permissions" -ForegroundColor Green
    } else {
        Write-Host "Missing required permissions: $($missingPerms -join ', ')" -ForegroundColor Yellow
        Write-Host "Reconnecting with all required permissions..." -ForegroundColor Yellow
    }
} else {
    Write-Host "Not connected to Microsoft Graph. Connecting now..." -ForegroundColor Yellow
}

# Connect if we need to
if (-not $hasAllPerms) {
    try {
        Connect-MgGraph -Scopes $requiredPerms -ErrorAction Stop -NoWelcome
        Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit
    }
}

# Get organization information
$orgname = Invoke-MgGraphRequest -Uri "beta/organization" -OutputType PSObject | Select-Object -ExpandProperty Value | Select-Object -ExpandProperty DisplayName

# Download the translation table
$translationTable = Invoke-RestMethod -Method Get -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" | ConvertFrom-Csv

#Get all users including sign-in and assigned license information
$uri = "beta/users?`$select=Id,accountenabled,DisplayName,UserPrincipalName,signInActivity,AssignedLicenses&`$top=999"
$Result = Invoke-MgGraphRequest -Uri $Uri -OutputType PSObject
$AllUsers = $Result.value
$NextLink = $Result."@odata.nextLink"
while ($null -ne $NextLink) {
    $Result = Invoke-MgGraphRequest -Method GET -Uri $NextLink -OutputType PSObject
    $AllUsers += $Result.value
    $NextLink = $Result."@odata.nextLink"
}

##Get tenant license usage information
$Report = [System.Collections.Generic.List[Object]]::new()
#Get all enabled licenses
$licenses = Invoke-MgGraphRequest -Uri "Beta/subscribedSkus" -OutputType PSObject | Select-Object -Expand Value | Where-Object {$_.CapabilityStatus -eq 'Enabled'}
#Get all directory subscription information
$directorySubscription = Invoke-MgGraphRequest -Uri "beta/directory/subscriptions" -OutputType PSObject | Select-Object -Expand Value
#Loop throuugh all licenses
Foreach ($license in $licenses){
    #Translaste the license name
    $licensename = $skuNamePretty = ($translationTable | Where-Object {$_.GUID -eq $license.skuId} | Sort-Object Product_Display_Name -Unique).Product_Display_Name
    If (($licensename -eq "") -or ($null -eq $licensename)){
        $licensename = $license.skuPartNumber
    }
    #Create a custom object with the license information
    $obj = [PSCustomObject][ordered]@{
        "License SKU" = $licensename
        "Type" = If (($directorySubscription | Where-Object {$_.skuId -eq $license.SkuId}).IsTrial -eq $true) {"Trial"} else {"Paid"}
        "Total Licenses" = $license.PrepaidUnits.Enabled
        "Used Licenses" = $license.ConsumedUnits
        "Unused licenses" = $license.PrepaidUnits.Enabled - $license.ConsumedUnits
        "Renewal/Expiratrion Date" = ($directorySubscription | Where-Object {$_.skuId -eq $license.SkuId}).NextLifecycleDateTime
    }
    $report.Add($obj)
}

#Obtain licenses users who have not successfully signed in, in the last 90 days.
#Unable to combine the needed filters
$90daysago = (Get-Date).AddDays(-90).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
$inactiveUsers = $AllUsers | Where-Object {$_.SignInActivity.LastSuccessfulSignInDateTime -lt $90daysago}
$licensedInactiveUsers = $inactiveUsers | Where-Object {$_.assignedLicenses -ne $null}
$licensedInactiveUsersReport = [System.Collections.Generic.List[Object]]::new()
#Loop through all inactive users
Foreach ($user in $licensedInactiveUsers){
    $skuNamePretty = @()
    foreach ($individuallicense in $user.AssignedLicenses.SkuId){
        $skuNamePretty += ($translationTable | Where-Object {$_.GUID -eq $individuallicense} | Sort-Object Product_Display_Name -Unique)."Product_Display_Name"
    }
    $obj = [pscustomobject][ordered]@{
        Name = $user.UserPrincipalName
        Licenses = ($skuNamePretty -join [environment]::NewLine) ##Updated -join "---"
        AccountEnabled = $user.accountenabled
        lastSignInAttempt = $user.SignInActivity.LastSignInDateTime
        lastSuccessfulSignIn = $user.SignInActivity.LastSuccessfulSignInDateTime
    }
    $licensedInactiveUsersReport.Add($obj)
}

##Obtain over-licensed privileged users
#Check tenant Entra level
$items = @("AAD_PREMIUM_P2", "AAD_PREMIUM", "AAD_BASIC")
$Skus = Invoke-MgGraphRequest -Uri "Beta/subscribedSkus" -OutputType PSObject | Select-Object -Expand Value
foreach ($item in $items) {
    $Search = $skus | Where-Object {$_.ServicePlans.servicePlanName -contains "$item"}
    if ($Search) {
        $licenseplan = $item
        break 
    } ElseIf ((!$Search) -and ($item -eq "AAD_BASIC")){
        $licenseplan = $item
        break
    }
}
#Get all users assigned roles
If ($licenseplan -eq "AAD_PREMIUM_P2") {
    $EligiblePIMRoles = Invoke-MgGraphRequest -Uri "beta/roleManagement/directory/roleEligibilitySchedules?`$expand=*" -OutputType PSObject | Select-Object -Expand Value
    $AssignedPIMRoles = Invoke-MgGraphRequest -Uri "beta/roleManagement/directory/roleAssignmentSchedules?`$expand=*" -OutputType PSObject | Select-Object -Expand Value
    $DirectoryRoles = $EligiblePIMRoles + $AssignedPIMRoles
    $PrivilegedRoles = $DirectoryRoles | Where-Object {
        ($_.RoleDefinition.DisplayName  -like "*Administrator*") -or ($_.RoleDefinition.DisplayName -like "*Writer*") -or ($_.RoleDefinition.DisplayName -eq "Global Reader")
    }
    $PrivilegedRoleUsers = $PrivilegedRoles | Where-Object {$_.Principal.'@odata.type' -eq "#microsoft.graph.user"}
    $RoleMembers = $PrivilegedRoleUsers.Principal.userPrincipalName | Select-Object -Unique
    $PrivilegedUsers = $RoleMembers | ForEach-Object { Invoke-MgGraphRequest -uri "/beta/users/$($_)?`$select=displayName,UserPrincipalName,AssignedLicenses" -OutputType PSobject }
}else{
    $DirectoryRoles = Invoke-MgGraphRequest -Uri "/beta/directoryRoles?" -OutputType PSObject | Select-Object -Expand Value
    $PrivilegedRoles = $DirectoryRoles | Where-Object {
        ($_.DisplayName  -like "*Administrator*") -or ($_.DisplayName -like "*Writer*") -or ($_.DisplayName -eq "Global Reader")
    }
    $RoleMembers = $PrivilegedRoles | ForEach-Object { Invoke-MgGraphRequest -uri "/beta/directoryRoles/$($_.id)/members" -OutputType PSObject | Select-Object -Expand Value} | Select-Object Id -Unique
    $PrivilegedUsers = $RoleMembers | ForEach-Object { Invoke-MgGraphRequest -uri "/beta/users/$($_.id)?`$select=displayName,UserPrincipalName,AssignedLicenses" -OutputType PSobject }
}
#Generate report
$overLicensedPrivUsers = [System.Collections.Generic.List[Object]]::new()
$translationTable = Invoke-RestMethod -Method Get -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" | ConvertFrom-Csv
Foreach ($user in $PrivilegedUsers){
    $licenses = @()
    If (($user.assignedLicenses.skuid -notin "41781fb2-bc02-4b7c-bd55-b576c07bb09d,eec0eb4f-6444-4f95-aba0-50c24d67f998") -and ($user.assignedLicenses.skuid.count -gt 0)){
        foreach($guid in $user.assignedLicenses.skuid) {
            $temp = ($translationTable | Where-Object {$_.GUID -eq $guid} | Sort-Object Product_Display_Name -Unique).Product_Display_Name
            $licenses += $temp
        }  
        $Obj2 = [pscustomobject][ordered]@{
            DisplayName = $user.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            License = $licenses -join [environment]::NewLine
        }
        $overLicensedPrivUsers.Add($Obj2)   
    }
}

###This is in progress
##Get users with duplicate licenses
#First obtain a list of all users and the friendly name of the license they have
$AllLicensedUsersReport = [System.Collections.Generic.List[Object]]::new()
$AllLicensedUsers = $AllUsers | Where-Object {$_.AssignedLicenses.count -ne 0}
Foreach ($user in $AllLicensedUsers){
    $licenses = @()
    foreach($guid in $user.assignedLicenses.skuid) {
        $temp = ($translationTable | Where-Object {$_.GUID -eq $guid} | Sort-Object Product_Display_Name -Unique).Product_Display_Name
        $licenses += $temp
    }  
    $Obj2 = [pscustomobject][ordered]@{
        DisplayName = $user.DisplayName
        UserPrincipalName = $User.UserPrincipalName
        License = $licenses -join [environment]::NewLine
    }
    $AllLicensedUsersReport.Add($Obj2)  
}


# Calculate totals for summary
$totalLicenses = ($report | Measure-Object "Total Licenses" -Sum).Sum
$totalUsed = ($report | Measure-Object "Used Licenses" -Sum).Sum
$totalUnused = ($report | Measure-Object "Unused licenses" -Sum).Sum
$unusedPercentage = [math]::Round(($totalUnused / $totalLicenses) * 100, 2)

# Get the count of inactive licensed users
$inactiveUsersCount = $licensedInactiveUsersReport.Count

# Get the count of over-licensed privileged users
$overLicensedPrivUsersCount = $overLicensedPrivUsers.Count

# Get the total number of licensed users
$totalLicensedUsersCount = $AllLicensedUsers.Count

# Generate HTML report
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 License Usage Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f5f5f5;
        }
        .header-container {
            background: linear-gradient(135deg, #0078D4 0%, #106EBE 100%);
            color: white;
            padding: 25px 40px;
            margin-bottom: 30px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }
        .header-content {
            max-width: 1200px;
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        h1 {
            font-size: 28px;
            font-weight: 600;
            margin: 0;
            letter-spacing: -0.5px;
        }
        .header-subtitle {
            font-size: 14px;
            font-weight: 400;
            margin-top: 0px;
            margin-bottom: 10px;
            opacity: 0.9;
        }
        .author-info {
            margin-top: 12px;
            border-top: 1px solid rgba(255, 255, 255, 0.3);
            padding-top: 10px;
            display: flex;
            align-items: center;
            font-size: 13px;
        }
        .author-label {
            opacity: 0.8;
            margin-right: 6px;
        }
        .author-links {
            display: flex;
            align-items: center;
        }
        .author-link {
            color: white;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            border: 1px solid rgba(255, 255, 255, 0.5);
            padding: 4px 10px;
            border-radius: 4px;
            margin-right: 10px;
            transition: all 0.2s ease;
            background-color: rgba(255, 255, 255, 0.1);
        }
        .author-link:hover {
            background-color: rgba(255, 255, 255, 0.2);
            border-color: rgba(255, 255, 255, 0.7);
        }
        .author-link svg {
            margin-right: 5px;
        }
        .report-info {
            text-align: right;
            font-size: 14px;
        }
        .report-date {
            font-weight: 500;
            margin-top: 5px;
        }
        .content-container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 40px 40px;
        }
        .summary-cards {
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
            margin-bottom: 40px;
            gap: 20px;
        }
        .summary-card {
            background-color: white;
            border-radius: 8px;
            padding: 20px;
            width: calc(33.333% - 14px);
            box-sizing: border-box;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            text-align: center;
        }
        @media (max-width: 768px) {
            .summary-card {
                width: 100%;
                margin-bottom: 15px;
            }
        }
        .card-title {
            font-size: 14px;
            color: #666;
            text-transform: uppercase;
            margin: 0 0 10px 0;
        }
        .card-value {
            font-size: 36px;
            font-weight: bold;
            margin: 0;
            color: #333;
        }
        .card-percentage {
            font-size: 14px;
            color: #666;
            margin-top: 5px;
        }
        table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            margin-top: 20px;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
            background-color: white;
        }
        th, td {
            padding: 15px;
            text-align: left;
        }
        th {
            background-color: #B5D8EB;
            color: #333;
            font-weight: bold;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            border-bottom: 2px solid #ddd;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f1f1f1;
        }
        td {
            border-bottom: 1px solid #ddd;
        }
        tr:last-child td {
            border-bottom: none;
        }
        .footer {
            text-align: center;
            margin-top: 30px;
            font-size: 12px;
            color: #666;
        }
        .high-unused {
            background-color: #ffe0e0 !important;
        }
        tr.high-unused:nth-child(even) {
            background-color: #ffdddd !important;
        }
        .export-btn {
            display: inline-block;
            padding: 10px 20px;
            background-color: #0078D4;
            color: white;
            text-decoration: none;
            border-radius: 4px;
            margin-top: 20px;
            font-weight: 500;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        }
        .export-btn:hover {
            background-color: #106EBE;
        }
        .switch {
            position: relative;
            display: inline-block;
            width: 60px;
            height: 34px;
            margin-right: 10px;
        }
        .switch input {
            opacity: 0;
            width: 0;
            height: 0;
        }
        .slider {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: #ccc;
            transition: .4s;
            border-radius: 34px;
        }
        .slider:before {
            position: absolute;
            content: "";
            height: 26px;
            width: 26px;
            left: 4px;
            bottom: 4px;
            background-color: white;
            transition: .4s;
            border-radius: 50%;
        }
        input:checked + .slider {
            background-color: #0078D4;
        }
        input:checked + .slider:before {
            transform: translateX(26px);
        }
        .filter-container {
            display: flex;
            align-items: center;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }
        .filter-label {
            margin-left: 10px;
            font-weight: 500;
        }
        .filter-group {
            display: flex;
            align-items: center;
            margin-right: 30px;
            margin-bottom: 10px;
        }
    </style>
</head>
<body>
    <div class="header-container">
        <div class="header-content">
            <div>
                <h1>Microsoft 365 License Usage Report</h1>
                <div class="header-subtitle">Overview of license allocation and usage across your tenant</div>
                <div class="author-info">
                    <span class="author-label">Created by:</span>
                    <div class="author-links">
                        <a href="https://www.linkedin.com/in/danielbradley2/" class="author-link" target="_blank">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="white">
                                <path d="M19 0h-14c-2.761 0-5 2.239-5 5v14c0 2.761 2.239 5 5 5h14c2.762 0 5-2.239 5-5v-14c0-2.761-2.238-5-5-5zm-11 19h-3v-11h3v11zm-1.5-12.268c-.966 0-1.75-.79-1.75-1.764s.784-1.764 1.75-1.764 1.75.79 1.75 1.764-.783 1.764-1.75 1.764zm13.5 12.268h-3v-5.604c0-3.368-4-3.113-4 0v5.604h-3v-11h3v1.765c1.396-2.586 7-2.777 7 2.476v6.759z"/>
                            </svg>
                            Daniel Bradley
                        </a>
                        <a href="https://ourcloudnetwork.com" class="author-link" target="_blank">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="white">
                                <path d="M21 13v10h-21v-19h12v2h-10v15h17v-8h2zm3-12h-10.988l4.035 4-6.977 7.07 2.828 2.828 6.977-7.07 4.125 4.172v-11z"/>
                            </svg>
                            ourcloudnetwork.com
                        </a>
                    </div>
                </div>
            </div>
            <div class="report-info">
                <div class="report-date">Generated: $(Get-Date -Format "MMMM d, yyyy")</div>
                <div class="tenant">Org: $orgname</div>
            </div>
        </div>
    </div>
     
    <div class="content-container">
        <div class="summary-cards">
            <div class="summary-card">
                <h3 class="card-title">Total Licenses</h3>
                <p class="card-value" id="totalLicensesValue">$totalLicenses</p>
                <p class="card-percentage" id="totalSubscriptions">Across $($report.Count) subscriptions</p>
            </div>
            <div class="summary-card">
                <h3 class="card-title">Used Licenses</h3>
                <p class="card-value" id="usedLicensesValue">$totalUsed</p>
                <p class="card-percentage" id="usedPercentage">$([math]::Round(($totalUsed / $totalLicenses) * 100, 2))% of total</p>
            </div>
            <div class="summary-card">
                <h3 class="card-title">Unused Licenses</h3>
                <p class="card-value" id="unusedLicensesValue">$totalUnused</p>
                <p class="card-percentage" id="unusedPercentage">$unusedPercentage% of total</p>
            </div>
            <div class="summary-card">
                <h3 class="card-title">Licensed Users</h3>
                <p class="card-value">$totalLicensedUsersCount</p>
                <p class="card-percentage">Users with assigned licenses</p>
            </div>
            <div class="summary-card">
                <h3 class="card-title">Inactive Users</h3>
                <p class="card-value">$inactiveUsersCount</p>
                <p class="card-percentage">90+ days without sign-in & licensed</p>
            </div>
            <div class="summary-card">
                <h3 class="card-title">Over-licensed Admins</h3>
                <p class="card-value">$overLicensedPrivUsersCount</p>
                <p class="card-percentage">Privileged users with unnecessary licenses</p>
            </div>
        </div>
         
        <div class="filter-container">
            <div class="filter-group">
                <label class="switch">
                    <input type="checkbox" id="toggleUnused" onchange="applyFilters()">
                    <span class="slider"></span>
                </label>
                <span class="filter-label">Show only subscriptions with unused licenses</span>
            </div>
             
            <div class="filter-group">
                <label class="switch">
                    <input type="checkbox" id="toggleHideFree" onchange="applyFilters()">
                    <span class="slider"></span>
                </label>
                <span class="filter-label">Hide free/promotional licenses</span>
            </div>
             
            <div class="filter-group">
                <label class="switch">
                    <input type="checkbox" id="toggleHideTrial" onchange="applyFilters()">
                    <span class="slider"></span>
                </label>
                <span class="filter-label">Hide trial licenses</span>
            </div>
        </div>
         
        <h2>License Details</h2>
        <table>
            <thead>
                <tr>
"@

# Add table headers
foreach ($header in $report[0].PSObject.Properties.Name) {
    $html += " <th>$header</th>`n"
}

$html += @"
                </tr>
            </thead>
            <tbody>
"@

# Add table rows with highlighting for rows with unused licenses
foreach ($row in $report) {
    # Only highlight rows where "Unused licenses" is greater than 0
    $rowClass = if ($row."Unused licenses" -gt 0) { 
        ' class="high-unused"' 
    } else { 
        '' 
    }
    
    $html += " <tr$rowClass>`n"
    foreach ($header in $row.PSObject.Properties.Name) {
        $value = $row.$header
        # Format date if it's a DateTime object
        if ($header -eq "Renewal/Expiratrion Date" -and $value -is [DateTime]) {
            $value = $value.ToString("yyyy-MM-dd")
        }
        $html += " <td>$value</td>`n"
    }
    $html += " </tr>`n"
}

$html += @"
            </tbody>
        </table>
         
        <h2 style="margin-top: 40px;">Inactive Licensed Users</h2>
        <p>The following users have licenses assigned but haven't successfully signed in for 90+ days.</p>
         
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Licenses</th>
                    <th>Account Enabled</th>
                    <th>Last Sign-In Attempt</th>
                    <th>Last Successful Sign-In</th>
                </tr>
            </thead>
            <tbody>
"@

# Add rows for inactive licensed users
foreach ($user in $licensedInactiveUsersReport) {
    $html += @"
                <tr>
                    <td>$($user.Name)</td>
                    <td style="white-space: pre-line;">$($user.Licenses)</td>
                    <td>$($user.AccountEnabled)</td>
                    <td>$($user.lastSignInAttempt)</td>
                    <td>$($user.lastSuccessfulSignIn)</td>
                </tr>
"@
}

$html += @"
            </tbody>
        </table>
         
        <h2 style="margin-top: 40px;">Over-Licensed Privileged Users</h2>
        <p>The following users have administrative roles but may have unnecessary licenses assigned.</p>
         
        <table>
            <thead>
                <tr>
                    <th>Display Name</th>
                    <th>User Principal Name</th>
                    <th>Licenses</th>
                </tr>
            </thead>
            <tbody>
"@

# Add rows for over-licensed privileged users
foreach ($user in $overLicensedPrivUsers) {
    $html += @"
                <tr>
                    <td>$($user.DisplayName)</td>
                    <td>$($user.UserPrincipalName)</td>
                    <td style="white-space: pre-line;">$($user.License)</td>
                </tr>
"@
}

$html += @"
            </tbody>
        </table>
         
        <div class="footer">
            <p>License usage report generated via Microsoft Graph API. Subscriptions with unused licenses are highlighted in red.</p>
        </div>
    </div>
     
    <script>
        // Wait for document to be fully loaded
        document.addEventListener('DOMContentLoaded', function() {
            // Attach event handlers
            document.getElementById('toggleUnused').onclick = function() { filterData(); };
            document.getElementById('toggleHideFree').onclick = function() { filterData(); };
            document.getElementById('toggleHideTrial').onclick = function() { filterData(); };
        });
 
        // The main filter function
        function filterData() {
            console.log('Filtering data...');
             
            // Get filter states
            let showOnlyUnused = document.getElementById('toggleUnused').checked;
            let hideFree = document.getElementById('toggleHideFree').checked;
            let hideTrial = document.getElementById('toggleHideTrial').checked;
             
            // Free licenses to filter
            let freeLicenses = [
                'windows store for business',
                'microsoft power automate free',
                'power virtual agents viral trial',
                'rights management adhoc',
                'power pages vtrial for makers',
                'microsoft power apps for developer'
            ];
             
            // Get the table and all data rows
            let table = document.getElementsByTagName('table')[0];
            if (!table) {
                console.error('Table not found!');
                return;
            }
             
            let tbody = table.getElementsByTagName('tbody')[0];
            if (!tbody) {
                console.error('Table body not found!');
                return;
            }
             
            let rows = tbody.getElementsByTagName('tr');
            if (rows.length === 0) {
                console.error('No rows found in table!');
                return;
            }
             
            // Tracking totals
            let totalSum = 0;
            let usedSum = 0;
            let unusedSum = 0;
            let visibleCount = 0;
             
            // Process each row
            for (let i = 0; i < rows.length; i++) {
                let row = rows[i];
                let cells = row.getElementsByTagName('td');
                 
                // Skip if we don't have enough cells
                if (cells.length < 5) continue;
                 
                // Extract data from cells
                let licenseName = cells[0].textContent.trim().toLowerCase();
                let licenseType = cells[1].textContent.trim().toLowerCase();
                let totalLicenses = parseInt(cells[2].textContent) || 0;
                let usedLicenses = parseInt(cells[3].textContent) || 0;
                let unusedLicenses = parseInt(cells[4].textContent) || 0;
                 
                // Apply filters
                let showRow = true;
                 
                // Filter: Only show rows with unused licenses
                if (showOnlyUnused && unusedLicenses <= 0) {
                    showRow = false;
                }
                 
                // Filter: Hide free licenses
                if (hideFree) {
                    for (let j = 0; j < freeLicenses.length; j++) {
                        if (licenseName.includes(freeLicenses[j])) {
                            showRow = false;
                            break;
                        }
                    }
                }
                 
                // Filter: Hide trial licenses
                if (hideTrial && licenseType === 'trial') {
                    showRow = false;
                }
                 
                // Apply visibility
                row.style.display = showRow ? '' : 'none';
                 
                // Add to totals if row is visible
                if (showRow) {
                    totalSum += totalLicenses;
                    usedSum += usedLicenses;
                    unusedSum += unusedLicenses;
                    visibleCount++;
                }
            }
             
            // Update summary cards
            let totalLicensesValue = document.getElementById('totalLicensesValue');
            let totalSubscriptions = document.getElementById('totalSubscriptions');
            let usedLicensesValue = document.getElementById('usedLicensesValue');
            let usedPercentageElem = document.getElementById('usedPercentage');
            let unusedLicensesValue = document.getElementById('unusedLicensesValue');
            let unusedPercentageElem = document.getElementById('unusedPercentage');
             
            if (totalLicensesValue) totalLicensesValue.textContent = totalSum;
            if (totalSubscriptions) totalSubscriptions.textContent = 'Across ' + visibleCount + ' subscriptions';
             
            if (usedLicensesValue) usedLicensesValue.textContent = usedSum;
            let usedPercent = totalSum > 0 ? Math.round((usedSum / totalSum) * 100) : 0;
            if (usedPercentageElem) usedPercentageElem.textContent = usedPercent + '% of total';
             
            if (unusedLicensesValue) unusedLicensesValue.textContent = unusedSum;
            let unusedPercent = totalSum > 0 ? Math.round((unusedSum / totalSum) * 100) : 0;
            if (unusedPercentageElem) unusedPercentageElem.textContent = unusedPercent + '% of total';
             
            console.log('Filtering complete. Visible rows:', visibleCount);
        }
         
        // Define legacy functions to maintain backward compatibility
        function applyFilters() {
            filterData();
        }
 
        function filterUnusedLicenses() {
            filterData();
        }
         
        function exportToCSV() {
            // Headers for the CSV
            const headers = [
                $(foreach($header in $report[0].PSObject.Properties.Name) { "'$header'," })
            ];
             
            // Data rows
            const rows = [
                headers.join(','),
                $(foreach($row in $report) {
                    $rowData = foreach($header in $row.PSObject.Properties.Name) {
                        if ($header -eq "Renewal/Expiratrion Date" -and $row.$header -is [DateTime]) {
                            "'$($row.$header.ToString("yyyy-MM-dd"))',"
                        } else {
                            "'$($row.$header)',"
                        }
                    }
                    "`n `"$($rowData -join ',')`","
                })
            ];
             
            // Create the CSV content
            const csvContent = "data:text/csv;charset=utf-8," + rows.join('\\r\\n');
             
            // Create a download link
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", "license_usage_report_$(Get-Date -Format "yyyy-MM-dd").csv");
            document.body.appendChild(link);
             
            // Trigger download
            link.click();
             
            document.body.removeChild(link);
        }
    </script>
</body>
</html>
"@

# Output the HTML to a file
$outputPathFile = "$outpath\M365_License_Usage_Report.html"
$html | Out-File -FilePath $outputPathFile -Encoding utf8

Write-Host "HTML report generated at $outputPathFile" -ForegroundColor Green