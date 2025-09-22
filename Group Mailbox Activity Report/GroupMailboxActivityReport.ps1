<#! 
.SYNOPSIS
    Group Activity (Mailbox usage) report with parallel collection.
    Generates per-group mailbox activity counts for 7/30/90 days, plus CSV and HTML (cards + email summary).

.DESCRIPTION
    This script enumerates a provided list of Microsoft 365 groups (M365/Unified, mail-enabled Security, 
    Distribution, or Entra ID security groups) and:
      1) Resolves each group to get ObjectId, type, owners/members, and metadata (EXO + AzureAD).
      2) Builds a unique list of *user-like* members and detects which have an Exchange Online mailbox.
      3) Uses Get-MailboxStatistics to measure LastLogonTime for each mailbox and aggregates counts:
         - Active in last 7 days
         - Active in last 30 days
         - Active in last 90 days
      4) Renders results to CSV and a detailed HTML dashboard (per-group “cards”), and
         composes an email summary (compact HTML) with both files attached via SMTP AUTH.

    Performance
      - Uses a lightweight runspace pattern with a simple throttle loop (MaxConcurrency).
      - Each worker creates its own EXO/AzureAD app-only connection to avoid cross-runspace reuse issues.

    Behavior Notes & Limits
      - “Users” = directory user objects in the group.
      - “Mailbox owners” = subset of users that actually have an Exchange Online mailbox.
      - Percentages use MailboxOwners as the denominator.
      - LastLogonTime semantics depend on mailbox activity that EXO records (background/service access 
        may not always update; shared/resource mailboxes can behave differently).
      - AzureAD module (AAD Graph) is used here for group membership lookup. It is deprecated by Microsoft; 
        consider migrating to Microsoft Graph PowerShell for future durability.

.PARAMETER TenantId
    Entra tenant ID (GUID). Used by AzureAD (app-only) and for tenant scoping.

.PARAMETER AppId
    Application (client) ID of the registered app used for app-only authentication.

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate (in LocalMachine/CurrentUser\My) used by EXO and AzureAD app-only auth.

.PARAMETER Organization
    Exchange Online organization (usually your primary SMTP domain or *.onmicrosoft.com).

.PARAMETER GroupsToReport
    Array of group identities (DisplayName or ObjectId GUID) to include in the report.

.PARAMETER OutDir
    Folder where CSV/HTML outputs are saved.

.PARAMETER Smtp settings (MailFrom, MailTo, SmtpServer, etc.)
    Credentials and server info for sending the email summary with attachments.

.OUTPUTS
    - CSV:  Group_Activity_Report.csv
    - HTML (Dashboard): Group_Activity_Report.html
    - Email: Summary table embedded; attaches CSV + full HTML.

.EXAMPLE
    # 1) Ensure certificate with private key is installed (LocalMachine\My), and the App Registration has
    #    Exchange.ManageAsApp + Directory.Read.All (application) consented by an admin.
    # 2) Update the CONFIG block (TenantId, AppId, CertificateThumbprint, Organization, SMTP, OutDir).
    # 3) List your groups in $GroupsToReport (display names or GUIDs).
    # 4) Run from an elevated, 64-bit Windows PowerShell 5.1 session:
    PS> .\GroupMailboxActivityReport.ps1

.NOTES
    Author  : Mohammad Abdelkader
    Website : https://momar.tech
    Date    : 2025-09-22
    Version : 1.0
#>

# ================== Configuration ==================
$TenantId              = "<your-tenant-guid>"
$AppId                 = "<your-app-id>"
$CertificateThumbprint = "<your-certificate-thumbprint>"
$Organization          = "<yourdomain.onmicrosoft.com>"   # or custom primary SMTP domain

# ================== SMTP settings ==================
$MailFrom    = "sender@domain.com"
$MailTo      = "recipient1@domain.com, recipient2@domain.com"
$MailCc      = ""
$MailBcc     = ""
$SmtpServer  = "smtp.office365.com"
$SmtpPort    = 587
$User        = $MailFrom
$Password    = ConvertTo-SecureString "<smtp-password-or-app-password>" -AsPlainText -Force
$Credential  = New-Object System.Management.Automation.PSCredential ($User, $Password)

# ================== Output ==================
$OutDir            = "C:\Reports"
$CsvPath           = Join-Path $OutDir "Group_Activity_Report.csv"
$HtmlPath          = Join-Path $OutDir "Group_Activity_Report.html"
$ReportTitle       = "Group Mailbox Activity Report"
$ReportDescription = "Mailbox activity summary (7/30/90 days) for the selected groups."

# ================== Groups ==================
$GroupsToReport = @(
    # Replace with group display names or ObjectId GUIDs
    "00000000-0000-0000-0000-000000000001"
    "00000000-0000-0000-0000-000000000002"
)


# ================== Prep ==================
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
if (-not (Test-Path $OutDir)) { New-Item -Path $OutDir -ItemType Directory -Force | Out-Null }
$w = { param($m,$lvl='INFO') $c=@{INFO='Cyan';SUCCESS='Green';WARN='Yellow';ERROR='Red'}[$lvl]; Write-Host $m -ForegroundColor $c }

# ================== Parallel Collect ==================
$MaxConcurrency = 10

# Thread-safe target for results
$rowsSync = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))

# Define the worker that handles ONE group (self-contained; no external functions)
$Worker = {
    param(
        [string]$GroupItem,
        [string]$TenantId,
        [string]$AppId,
        [string]$CertificateThumbprint,
        [string]$Organization
    )

    # ----- Minimal helpers in runspace -----
    $w = { param($m,$lvl='INFO') $c=@{INFO='Cyan';SUCCESS='Green';WARN='Yellow';ERROR='Red'}[$lvl]; Write-Host $m -ForegroundColor $c }
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    # Modules & connections are per-runspace
    if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop | Out-Null

    if (-not (Get-Module -ListAvailable AzureAD)) {
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
        }
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        Install-Module AzureAD -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    }
    Import-Module AzureAD -ErrorAction Stop | Out-Null

    Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false -ErrorAction Stop | Out-Null
    Connect-AzureAD        -TenantId $TenantId -ApplicationId $AppId -CertificateThumbprint $CertificateThumbprint | Out-Null

    # ----- Local resolver (no external deps) -----
    $ResolveGroupLocal = {
      param($InputNameOrId)

      $info = [ordered]@{
        Name=$InputNameOrId; ObjectId=$null; Email=$null; Type=$null
        CreatedDateTime=$null; OwnersCount=$null; MembersCount=$null; Source=$null
      }
      $enrichFromAAD = {
        param([string]$oid)
        if ([string]::IsNullOrWhiteSpace($oid)) { return }
        try {
          $aad = Get-AzureADGroup -ObjectId $oid -ErrorAction Stop
          if ($aad) {
            $info.Name  = $aad.DisplayName
            if (-not $info.Email) { $info.Email = $aad.Mail }
          }
        } catch {}
      }

      if ($InputNameOrId -match '^[0-9a-fA-F-]{36}$') {
        try {
          $g = Get-AzureADGroup -ObjectId $InputNameOrId -ErrorAction Stop
          if ($g) {
            $info.Name=$g.DisplayName; $info.ObjectId=$g.ObjectId; $info.Email=$g.Mail
            if ($g.MailEnabled -and -not $g.SecurityEnabled) { $info.Type = 'GroupMailbox' } else { $info.Type = 'Security (AAD)' }
            $info.CreatedDateTime=$g.CreationDateTime; $info.Source='AzureAD'
            $info.MembersCount = @(Get-AzureADGroupMember -ObjectId $g.ObjectId -All $true -ErrorAction SilentlyContinue).Count
            $info.OwnersCount  = @(Get-AzureADGroupOwner  -ObjectId $g.ObjectId -All $true -ErrorAction SilentlyContinue).Count
            return [pscustomobject]$info
          }
        } catch {}
      }

      try {
        $r = Get-Recipient -Identity $InputNameOrId -ErrorAction Stop
        if ($r) {
          $info.Email  = $r.PrimarySmtpAddress
          $info.Type   = $r.RecipientTypeDetails
          $info.Source = "Exchange"

          switch ($r.RecipientTypeDetails) {
            'GroupMailbox' {
              $g = Get-UnifiedGroup -Identity $InputNameOrId -ErrorAction SilentlyContinue
              if ($g) {
                $info.ObjectId=$g.ExternalDirectoryObjectId; $info.CreatedDateTime=$g.WhenCreatedUTC
                $info.OwnersCount=@(Get-UnifiedGroupLinks -Identity $g.Identity -LinkType Owners  -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
                $info.MembersCount=@(Get-UnifiedGroupLinks -Identity $g.Identity -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
                $info.Name=$g.DisplayName; & $enrichFromAAD $info.ObjectId
                return [pscustomobject]$info
              }
            }
            'MailUniversalDistributionGroup' {
              $dg = Get-DistributionGroup -Identity $r.Identity -ErrorAction SilentlyContinue
              if ($dg) {
                $info.ObjectId=$dg.ExternalDirectoryObjectId; $info.CreatedDateTime=$dg.WhenCreatedUTC
                $info.MembersCount=@(Get-DistributionGroupMember -Identity $dg.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
                $info.OwnersCount=@($dg.ManagedBy).Count; $info.Name=$dg.DisplayName
                & $enrichFromAAD $info.ObjectId
                return [pscustomobject]$info
              }
            }
            'MailUniversalSecurityGroup' {
              $sg = Get-DistributionGroup -Identity $r.Identity -ErrorAction SilentlyContinue
              if ($sg) {
                $info.ObjectId=$sg.ExternalDirectoryObjectId; $info.CreatedDateTime=$sg.WhenCreatedUTC
                $info.MembersCount=@(Get-DistributionGroupMember -Identity $sg.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
                $info.OwnersCount=@($sg.ManagedBy).Count; $info.Name=$sg.DisplayName
                & $enrichFromAAD $info.ObjectId
                return [pscustomobject]$info
              }
            }
          }
        }
      } catch {}

      try {
        $g = Get-AzureADGroup -Filter "DisplayName eq '$InputNameOrId'" -ErrorAction Stop
        if ($g) {
          $info.Name=$g.DisplayName; $info.ObjectId=$g.ObjectId; $info.Email=$g.Mail
          if ($g.MailEnabled -and -not $g.SecurityEnabled) { $info.Type = 'GroupMailbox' } else { $info.Type = 'Security (AAD)' }
          $info.CreatedDateTime=$g.CreationDateTime; $info.Source='AzureAD'
          $info.MembersCount=@(Get-AzureADGroupMember -ObjectId $g.ObjectId -All $true -ErrorAction SilentlyContinue).Count
          $info.OwnersCount=@(Get-AzureADGroupOwner  -ObjectId $g.ObjectId -All $true -ErrorAction SilentlyContinue).Count
          return [pscustomobject]$info
        }
      } catch {}

      $info.Type="Unknown"
      [pscustomobject]$info
    }

    $GetGroupUserMembersLocal = {
      param($GroupInfo, $OrgDomain)

      $users = @(); $mailboxOwners = @()
      $userLike = 'UserMailbox','User','MailUser','RemoteUserMailbox','LinkedMailbox','SharedMailbox','GuestMailUser'

      $buildId = {
        param($m)
        $id=$null
        foreach ($p in 'PrimarySmtpAddress','WindowsEmailAddress','ExternalEmailAddress','UserPrincipalName','WindowsLiveID','Alias') {
          if ($m.PSObject.Properties.Name -contains $p -and $m.$p) { $id = $m.$p; break }
        }
        if (-not $id -and ($m.PSObject.Properties.Name -contains 'ExternalDirectoryObjectId') -and $m.ExternalDirectoryObjectId) { $id = $m.ExternalDirectoryObjectId }
        $id
      }
      $hasMailbox = {
        param($identity)
        if (-not $identity) { return $false }
        $ok=$false
        foreach($candidate in @($identity)){
          try { $null = Get-Mailbox -Identity $candidate -ErrorAction Stop; $ok=$true; break } catch {}
        }
        if (-not $ok -and ($identity -notmatch '@')) {
          try {
            $domain = $OrgDomain
            $domain = ($domain -replace '.*?@','')   # keep backward-compat; but usually Organization is the domain
            $null = Get-Mailbox -Identity ("SMTP:{0}@{1}" -f $identity, $domain) -ErrorAction Stop
            $ok=$true
          } catch {}
        }
        $ok
      }

      switch ($GroupInfo.Type) {
        'GroupMailbox' {
          $members = Get-UnifiedGroupLinks -Identity $GroupInfo.Name -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue
          foreach ($m in $members) {
            $t = if ($m.PSObject.Properties.Name -contains 'RecipientTypeDetails') { $m.RecipientTypeDetails } else { $m.RecipientType }
            if ($t -in $userLike) {
              $id = & $buildId $m
              $users += [pscustomobject]@{ DisplayName=$m.DisplayName; Id=$id; Type=$t }
              if (& $hasMailbox $id) { $mailboxOwners += [pscustomobject]@{ DisplayName=$m.DisplayName; Id=$id; Type=$t } }
            }
          }
        }
        'MailUniversalDistributionGroup' {
          $members = Get-DistributionGroupMember -Identity $GroupInfo.Name -ResultSize Unlimited -ErrorAction SilentlyContinue
          foreach ($m in $members) {
            $t = if ($m.PSObject.Properties.Name -contains 'RecipientTypeDetails') { $m.RecipientTypeDetails } else { $m.RecipientType }
            if ($t -in $userLike) {
              $id = & $buildId $m
              $users += [pscustomobject]@{ DisplayName=$m.DisplayName; Id=$id; Type=$t }
              if (& $hasMailbox $id) { $mailboxOwners += [pscustomobject]@{ DisplayName=$m.DisplayName; Id=$id; Type=$t } }
            }
          }
        }
        'MailUniversalSecurityGroup' {
          $members = Get-DistributionGroupMember -Identity $GroupInfo.Name -ResultSize Unlimited -ErrorAction SilentlyContinue
          foreach ($m in $members) {
            $t = if ($m.PSObject.Properties.Name -contains 'RecipientTypeDetails') { $m.RecipientTypeDetails } else { $m.RecipientType }
            if ($t -in $userLike) {
              $id = & $buildId $m
              $users += [pscustomobject]@{ DisplayName=$m.DisplayName; Id=$id; Type=$t }
              if (& $hasMailbox $id) { $mailboxOwners += [pscustomobject]@{ DisplayName=$m.DisplayName; Id=$id; Type=$t } }
            }
          }
        }
        default {
          if ($GroupInfo.ObjectId) {
            $mems = Get-AzureADGroupMember -ObjectId $GroupInfo.ObjectId -All $true -ErrorAction SilentlyContinue |
                    Where-Object { $_.ObjectType -eq 'User' }
            foreach ($u in $mems) {
              $id = if ($u.Mail) { $u.Mail } else { $u.UserPrincipalName }
              $users += [pscustomobject]@{ DisplayName=$u.DisplayName; Id=$id; Type='User' }
              if (& $hasMailbox $id) { $mailboxOwners += [pscustomobject]@{ DisplayName=$u.DisplayName; Id=$id; Type='User' } }
            }
          }
        }
      }

      $users         = $users         | Group-Object { $_.Id } | ForEach-Object { $_.Group[0] }
      $mailboxOwners = $mailboxOwners | Group-Object { $_.Id } | ForEach-Object { $_.Group[0] }
      [pscustomobject]@{ Users=$users; MailboxOwners=$mailboxOwners }
    }

    $MeasureMailboxActivityLocal = {
      param($UsersWithMailboxes)
      $total=0;$a7=0;$a30=0;$a90=0
      $now=Get-Date;$t7=$now.AddDays(-7);$t30=$now.AddDays(-30);$t90=$now.AddDays(-90)
      foreach ($u in $UsersWithMailboxes) {
        try {
          $s = Get-MailboxStatistics -Identity $u.Id -ErrorAction Stop
          if ($s) {
            $total++
            $lad = $s.LastLogonTime
            if ($lad) {
              if ($lad -gt $t7)  { $a7++  }
              if ($lad -gt $t30) { $a30++ }
              if ($lad -gt $t90) { $a90++ }
            }
          }
        } catch {}
      }
      [pscustomobject]@{ TotalMailboxes=$total; ActiveLast7Days=$a7; ActiveLast30Days=$a30; ActiveLast90Days=$a90 }
    }

    # ----- Do one group -----
    &$w "Processing (worker): $GroupItem" 'INFO'
    $ginfo = & $ResolveGroupLocal $GroupItem
    if (-not $ginfo.ObjectId -and $ginfo.Type -eq 'Unknown') {
      return [pscustomobject]@{
        GroupName=$GroupItem; GroupType='Unknown'; ObjectId=$null; Email=$null; Created=$null
        Users=0; Owners=0; MailboxOwners=0; Active7='Error'; Active30='Error'; Active90='Error'
      }
    }

    $mm   = & $GetGroupUserMembersLocal $ginfo $Organization
    $usersList     = $mm.Users
    $mailboxOwners = $mm.MailboxOwners

    if ($ginfo.Type -eq 'Security (AAD)' -and $mailboxOwners.Count -eq 0 -and $usersList.Count -gt 0) {
      $mailboxOwners = @()
      foreach ($u in $usersList) { try { $null = Get-Mailbox -Identity $u.Id -ErrorAction Stop; $mailboxOwners += $u } catch {} }
    }

    $stats = & $MeasureMailboxActivityLocal $mailboxOwners

    return [pscustomobject]@{
      GroupName      = $ginfo.Name
      GroupType      = $ginfo.Type
      ObjectId       = $ginfo.ObjectId
      Email          = $ginfo.Email
      Created        = $ginfo.CreatedDateTime
      Users          = @($usersList).Count
      Owners         = $ginfo.OwnersCount
      MailboxOwners  = @($mailboxOwners).Count
      Active7        = $stats.ActiveLast7Days
      Active30       = $stats.ActiveLast30Days
      Active90       = $stats.ActiveLast90Days
    }
}

# Launch jobs with a simple throttle loop (PowerShell 5.1 friendly)
$pending = New-Object System.Collections.Queue
$GroupsToReport | ForEach-Object { $pending.Enqueue($_) }
$running = @()

&$w ("Starting parallel collection (max={0})..." -f $MaxConcurrency) 'INFO'

while ($pending.Count -gt 0 -or $running.Count -gt 0) {
    while (($running.Count -lt $MaxConcurrency) -and ($pending.Count -gt 0)) {
        $item = $pending.Dequeue()
        &$w ("Queue→Start: {0}" -f $item) 'INFO'
        $ps = [PowerShell]::Create()
        [void]$ps.AddScript($Worker).AddArgument($item).AddArgument($TenantId).AddArgument($AppId).AddArgument($CertificateThumbprint).AddArgument($Organization)
        $handle = $ps.BeginInvoke()
        $running += [pscustomobject]@{ PS=$ps; Handle=$handle; Item=$item }
    }

    foreach ($r in @($running)) {
        if ($r.Handle.IsCompleted) {
            try {
                $result = $r.PS.EndInvoke($r.Handle)
                foreach ($row in $result) { [void]$rowsSync.Add($row) }
                &$w ("Done: {0}" -f $r.Item) 'SUCCESS'
            } catch {
                &$w ("[ERROR] {0}: {1}" -f $r.Item, $_.Exception.Message) 'ERROR'
            } finally {
                $r.PS.Dispose()
                $running = $running | Where-Object { $_ -ne $r }
            }
        }
    }

    Start-Sleep -Milliseconds 200
}

# Normalize result type to your original $rows list
$rows = New-Object System.Collections.Generic.List[object]
foreach ($x in $rowsSync) { [void]$rows.Add($x) }

# ================== CSV ==================
$rows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
&$w "CSV saved: $CsvPath" 'SUCCESS'

# ================== HTML ==================
$Percent = { param([int]$p,[int]$w) if ($w -le 0) { "0%" } else { "{0}%" -f ([math]::Round(($p/$w)*100,2)) } }
$ClassFor = { param($a,$t) if ($t -le 0){'err'} elseif(($a/$t) -ge 0.7){'ok'} elseif($a -gt 0){'warn'} else {'err'} }

$css = @"
<style>
body{font-family:Segoe UI,Arial;margin:24px;background:#0b1220;color:#e9eef7}
h1{font-size:22px;margin:0 0 16px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:16px}
.card{background:#111a2e;border:1px solid #1e2b4a;border-radius:14px;padding:16px;box-shadow:0 4px 20px rgba(0,0,0,.25)}
.kv{display:grid;grid-template-columns:200px 1fr;gap:6px 12px;font-size:13px}
.kv .key{color:#a7b3cf}
.table{width:100%;border-collapse:collapse;margin-top:8px;font-size:13px}
.table th,.table td{border-bottom:1px solid #223357;padding:8px 10px;text-align:left}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;background:#1c2d52;border:1px solid #2f4475;font-size:12px}
.bar{height:10px;background:#1a2746;border:1px solid #2c3f6d;border-radius:999px;overflow:hidden}
.bar>span{display:block;height:100%}
.ok{background:#2aa198}.warn{background:#b58900}.err{background:#dc322f}
.small{font-size:12px;color:#a7b3cf}
.footer{margin-top:24px;color:#8ea2c9;font-size:12px}
.desc{margin:6px 0 4px;color:#a7b3cf}
.meta{display:grid;grid-template-columns:180px 1fr;gap:4px 12px;font-size:12px;color:#a7b3cf}
</style>
"@

$sections = foreach ($r in $rows) {
  $p7  = & $Percent $r.Active7  $r.MailboxOwners
  $p30 = & $Percent $r.Active30 $r.MailboxOwners
  $p90 = & $Percent $r.Active90 $r.MailboxOwners
  $c7  = & $ClassFor $r.Active7  $r.MailboxOwners
  $c30 = & $ClassFor $r.Active30 $r.MailboxOwners
  $c90 = & $ClassFor $r.Active90 $r.MailboxOwners
@"
<div class='card'>
  <h1>$($r.GroupName) <span class='badge'>$($r.GroupType)</span></h1>
  <div class='kv'>
    <div class='key'>Group name</div><div>$($r.GroupName)</div>
    <div class='key'>Object ID</div><div>$($r.ObjectId)</div>
    <div class='key'>Email</div><div>$($r.Email)</div>
    <div class='key'>Created</div><div>$($r.Created)</div>
    <div class='key'>Owners</div><div>$($r.Owners)</div>
    <div class='key'>Users (directory)</div><div>$($r.Users)</div>
    <div class='key'>Mailbox owners</div><div>$($r.MailboxOwners)</div>
  </div>
  <table class='table'>
    <thead><tr><th>Window</th><th>Active</th><th>% Active</th><th>Bar</th></tr></thead>
    <tbody>
      <tr><td>Last 7 days</td><td>$($r.Active7)</td><td>$p7</td>
          <td><div class='bar'><span class='$c7' style='width:$p7'></span></div></td></tr>
      <tr><td>Last 30 days</td><td>$($r.Active30)</td><td>$p30</td>
          <td><div class='bar'><span class='$c30' style='width:$p30'></span></div></td></tr>
      <tr><td>Last 90 days</td><td>$($r.Active90)</td><td>$p90</td>
          <td><div class='bar'><span class='$c90' style='width:$p90'></span></div></td></tr>
    </tbody>
  </table>
  <div class='small'>Notes:
   (1) “Users” mirrors Entra portal (user objects only). 
   (2) “Mailbox owners” are the subset with Exchange Online mailboxes; percentages use this number.</div>
</div>
"@
}

$totalGroups = $rows.Count
$totalUsers  = ($rows | Measure-Object Users -Sum).Sum
$totalMbx    = ($rows | Measure-Object MailboxOwners -Sum).Sum
$totalA7     = ($rows | Measure-Object Active7 -Sum).Sum
$totalA30    = ($rows | Measure-Object Active30 -Sum).Sum
$totalA90    = ($rows | Measure-Object Active90 -Sum).Sum
$pg7  = & $Percent $totalA7  $totalMbx
$pg30 = & $Percent $totalA30 $totalMbx
$pg90 = & $Percent $totalA90 $totalMbx

$html = @"
<html>
<head>
    <meta charset='utf-8'>
    <title>Group Activity Report</title>$css
</head>
<body>
    <div class='grid'>
        <div class='card'>
            <h1>$ReportTitle <span class='small'
                    style='display:block;font-weight:400;margin-top:6px'>$ReportDescription</span></h1>
            <div class='kv'>
                <div class='key'>Groups</div>
                <div>$totalGroups</div>
                <div class='key'>Users (sum)</div>
                <div>$totalUsers</div>
                <div class='key'>Mailbox owners (sum)</div>
                <div>$totalMbx</div>
                <div class='key'>Active (7d)</div>
                <div>$totalA7 ($pg7)</div>
                <div class='key'>Active (30d)</div>
                <div>$totalA30 ($pg30)</div>
                <div class='key'>Active (90d)</div>
                <div>$totalA90 ($pg90)</div>
            </div>
        </div>
    </div>
    <div class='grid'>
        $($sections -join "`n")
    </div>
    <div class='footer'>Generated on $(Get-Date).
        Data sources: Exchange Online PowerShell (Get-MailboxStatistics, Get-UnifiedGroup*, Get-DistributionGroup*),
        AzureAD module (Get-AzureADGroup*, app-only). Microsoft Docs links are in the script header.</div>
</body>
</html>
"@

$html | Out-File -FilePath $HtmlPath -Encoding UTF8
&$w "HTML saved: $HtmlPath" 'SUCCESS'
&$w "Done collecting and rendering." 'SUCCESS'

# ================== Email Summary (compact) ==================
# Recompute totals (safety)
if (-not $totalGroups) { $totalGroups = @($rows).Count }
if (-not $totalUsers)  { $totalUsers  = ($rows | Measure-Object Users -Sum).Sum }
if (-not $totalMbx)    { $totalMbx    = ($rows | Measure-Object MailboxOwners -Sum).Sum }
if (-not $totalA7)     { $totalA7     = ($rows | Measure-Object Active7 -Sum).Sum }
if (-not $totalA30)    { $totalA30    = ($rows | Measure-Object Active30 -Sum).Sum }
if (-not $totalA90)    { $totalA90    = ($rows | Measure-Object Active90 -Sum).Sum }

# Build numbered table rows with friendly type names
$sb = New-Object System.Text.StringBuilder
[int]$i = 0
foreach ($r in $rows) {
    $i++
    $friendlyType = switch ($r.GroupType) {
        'GroupMailbox'                  { 'M365 Group' }
        'MailUniversalDistributionGroup'{ 'Distribution' }
        'MailUniversalSecurityGroup'    { 'Security (mail-enabled)' }
        'Security (AAD)'                { 'Security' }
        default                         { [string]$r.GroupType }
    }
    $null = $sb.AppendLine(@"
<tr>
  <td style='text-align:right'>$i</td>
  <td style='word-break:break-word'>$([System.Web.HttpUtility]::HtmlEncode($r.GroupName))</td>
  <td>$([System.Web.HttpUtility]::HtmlEncode($friendlyType))</td>
  <td style='text-align:right'>$($r.Users)</td>
  <td style='text-align:right'>$($r.MailboxOwners)</td>
  <td style='text-align:right'>$($r.Active7)</td>
  <td style='text-align:right'>$($r.Active30)</td>
  <td style='text-align:right'>$($r.Active90)</td>
</tr>
"@)
}

$EmailSummaryHtml = @"
<html>
<head>
<meta charset='utf-8'>
<title>$ReportTitle – Summary</title>
<style>
  body{margin:0;padding:0;background:#ffffff;color:#0f172a;font-family:Segoe UI,Arial,Helvetica,sans-serif}
  .container{max-width:900px;margin:16px auto;padding:0 12px}
  .card{background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;padding:16px}
  h1{font-size:22px;margin:0 0 6px 0;color:#0f172a}
  .sub{font-size:13px;color:#475569;margin:0 0 14px 0}
  .summary{display:block;border:1px solid #e2e8f0;background:#f1f5f9;border-radius:10px;padding:10px 12px;margin:0 0 14px 0}
  .chips{display:flex;flex-wrap:wrap;gap:10px}
  .chip{display:flex;gap:6px;align-items:center;padding:4px 8px;background:#ffffff;border:1px solid #e2e8f0;border-radius:999px;font-size:12px}
  .k{color:#475569}
  .v{font-weight:600;color:#0f172a}
  table{width:100%;border-collapse:collapse;font-size:13px;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden}
  thead th{background:#f8fafc;color:#0f172a;font-weight:600;border-bottom:1px solid #e2e8f0;padding:8px}
  tbody td{border-bottom:1px solid #e2e8f0;padding:8px}
  .meta{font-size:12px;color:#475569;margin-top:10px}
  @media (prefers-color-scheme: dark) {
    body{background:#ffffff}
  }
</style>
</head>
<body>
  <div class="container">
    <div class="card">
      <h1>$ReportTitle</h1>
      <div class="sub">Mailbox activity summary (7/30/90 days) for the selected groups.</div>

      <div class="summary">
        <div class="chips">
          <div class="chip"><span class="k">Groups:</span><span class="v">$totalGroups</span></div>
          <div class="chip"><span class="k">Users (sum):</span><span class="v">$totalUsers</span></div>
          <div class="chip"><span class="k">Mailboxes (sum):</span><span class="v">$totalMbx</span></div>
          <div class="chip"><span class="k">Active 7d:</span><span class="v">$totalA7</span></div>
          <div class="chip"><span class="k">Active 30d:</span><span class="v">$totalA30</span></div>
          <div class="chip"><span class="k">Active 90d:</span><span class="v">$totalA90</span></div>
        </div>
      </div>

      <table role="table" aria-label="Group Activity Summary">
        <thead>
          <tr>
            <th style="width:40px;text-align:right">#</th>
            <th>Group</th>
            <th>Type</th>
            <th style="width:90px;text-align:right">Users</th>
            <th style="width:120px;text-align:right">Mailboxes</th>
            <th style="width:80px;text-align:right">7d</th>
            <th style="width:80px;text-align:right">30d</th>
            <th style="width:80px;text-align:right">90d</th>
          </tr>
        </thead>
        <tbody>
          $($sb.ToString())
        </tbody>
      </table>

      <div class="meta">Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss').</div>
    </div>
  </div>
</body>
</html>
"@

# Enforce TLS 1.2 (older hosts)
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# --- Send email (embed summary HTML; attach CSV + full HTML) ---
try {
    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From = $MailFrom
    ($MailTo -split '[;, ]+' | Where-Object { $_ -and $_.Contains('@') })  | ForEach-Object { [void]$mail.To.Add($_) }
    ($MailCc -split  '[;, ]+' | Where-Object { $_ -and $_.Contains('@') }) | ForEach-Object { [void]$mail.CC.Add($_) }
    ($MailBcc -split '[;, ]+' | Where-Object { $_ -and $_.Contains('@') }) | ForEach-Object { [void]$mail.Bcc.Add($_) }

    $mail.Subject    = "$ReportTitle – Summary ($(Get-Date -Format 'yyyy-MM-dd HH:mm'))"
    $mail.Body       = $EmailSummaryHtml
    $mail.IsBodyHtml = $true

    if (Test-Path $CsvPath)  { $mail.Attachments.Add([System.Net.Mail.Attachment]::new($CsvPath))  | Out-Null }
    if (Test-Path $HtmlPath) { $mail.Attachments.Add([System.Net.Mail.Attachment]::new($HtmlPath)) | Out-Null }

    $smtp = [System.Net.Mail.SmtpClient]::new($SmtpServer, $SmtpPort)
    $smtp.EnableSsl      = $true
    $smtp.Credentials    = $Credential
    $smtp.Timeout        = 120000
    $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network

    Write-Host "Sending email summary..." -ForegroundColor Cyan
    $smtp.Send($mail)
    Write-Host "Email summary sent to: $MailTo" -ForegroundColor Green
}
catch {
    Write-Host ("[ERROR] SMTP send failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
}
finally {
    if ($mail) { $mail.Dispose() }
}
