<#
.TITLE
    Get-GroupMailboxActivityReport - Group mailbox activity counts for 7-30-90 days

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
      4) Renders results to CSV and a detailed HTML dashboard (per-group cards), and
         composes an email summary (compact HTML) with both files attached via SMTP AUTH.

    Performance
      - Uses a lightweight runspace pattern with a simple throttle loop (MaxConcurrency).
      - Each worker creates its own EXO/AzureAD app-only connection to avoid cross-runspace reuse issues.

    Behavior Notes & Limits
      - Users = directory user objects in the group.
      - Mailbox owners = subset of users that actually have an Exchange Online mailbox.
      - Percentages use MailboxOwners as the denominator.
      - LastLogonTime semantics depend on mailbox activity that EXO records (background/service access
        may not always update; shared/resource mailboxes can behave differently).
      - AzureAD module (AAD Graph) is used here for group membership lookup. It is deprecated by Microsoft;
        consider migrating to Microsoft Graph PowerShell for future durability.

.TAGS
    Identity,M365,Exchange

.PLATFORM
    Windows 10/11/Server 2019+

.PERMISSIONS
    Exchange.ManageAsApp, Directory.Read.All

.AUTHOR
    AI Generated

.VERSION
    1.0.0

.CHANGELOG
    1.0.0 (2026-09-16) - Compliance hardening: canonical rich header, ErrorActionPreference Stop, alias and catch hygiene.

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Get-GroupMailboxActivityReport.ps1
    Runs the group mailbox activity report with default config.

.EXAMPLE
    Get-Help .\Get-GroupMailboxActivityReport.ps1 -Full
    Shows configuration variables ($GroupsToReport, SMTP, $OutDir) and behavior notes.

.NOTES
    Part of Microsoft365-Scripts toolkit - Identity,M365,Exchange
    Exit codes: 0 = success, 1 = failure, 2 = script error
    Elevation is detected at runtime via Test-IsElevated and degrades gracefully.
#>

#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

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

# ================== Output (anchored beside the script per Law 12) ==================
$scriptDirectory   = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$OutDir            = Join-Path $scriptDirectory "Reports"
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
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }
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
# why: legacy console dispatcher (also injected into runspaces); all call sites funnel here.
$w = { param($m,$lvl='INFO') $c=@{INFO='Cyan';SUCCESS='Green';WARN='Yellow';ERROR='Red'}[$lvl]; Write-Host $m -ForegroundColor $c }
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }

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
        } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }
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
        } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }
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
      } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }

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
      } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }

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
          try { $null = Get-Mailbox -Identity $candidate -ErrorAction Stop; $ok=$true; break } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }
        }
        if (-not $ok -and ($identity -notmatch '@')) {
          try {
            $domain = $OrgDomain
            $domain = ($domain -replace '.*?@','')   # keep backward-compat; but usually Organization is the domain
            $null = Get-Mailbox -Identity ("SMTP:{0}@{1}" -f $identity, $domain) -ErrorAction Stop
            $ok=$true
          } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }
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
        } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }
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
      foreach ($u in $usersList) { try { $null = Get-Mailbox -Identity $u.Id -ErrorAction Stop; $mailboxOwners += $u } catch { Write-Warning "Ignored error: $($_.Exception.Message)" } }
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

# ================== HTML (IBM Carbon Dark dashboard) ==================
$totalGroups = $rows.Count
$totalUsers  = ($rows | Measure-Object Users -Sum).Sum
$totalMbx    = ($rows | Measure-Object MailboxOwners -Sum).Sum
$totalA7     = ($rows | Measure-Object Active7 -Sum).Sum
$totalA30    = ($rows | Measure-Object Active30 -Sum).Sum
$totalA90    = ($rows | Measure-Object Active90 -Sum).Sum
$pctOf = { param($a,$t) if ($t -le 0) { '0%' } else { '{0}%' -f ([math]::Round(($a/$t)*100,2)) } }

$tableRows = foreach ($rowRef in ($rows | Select-Object -First 500)) {
    $p7 = & $pctOf $rowRef.Active7 $rowRef.MailboxOwners
    $p30 = & $pctOf $rowRef.Active30 $rowRef.MailboxOwners
    $p90 = & $pctOf $rowRef.Active90 $rowRef.MailboxOwners
    '<tr><td><code>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.GroupName)") + '</code></td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.GroupType)") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.Users)") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.MailboxOwners)") + '</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.Active7)") + ' (' + $p7 + ')</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.Active30)") + ' (' + $p30 + ')</td>' +
    '<td>' + [System.Net.WebUtility]::HtmlEncode("$($rowRef.Active90)") + ' (' + $p90 + ')</td></tr>'
}
$tableNote = if ($rows.Count -gt 500) { "<p>Showing 500 of $($rows.Count) groups; full data is in the CSV.</p>" } else { '' }
$tableHtml = '<div class="section-title">Group Activity Detail</div>' +
    '<div class="card"><h2>All Groups (' + $rows.Count + ')</h2>' +
    '<table><thead><tr><th>Group</th><th>Type</th><th>Users</th><th>Mailbox Owners</th><th>Active 7d</th><th>Active 30d</th><th>Active 90d</th></tr></thead><tbody>' +
    ($tableRows -join "`n") + '</tbody></table>' + $tableNote + '</div>'

$kpis = @(
    @{ value = "$totalGroups"; label = 'Groups'; color = '' },
    @{ value = "$totalMbx"; label = 'Mailbox owners'; color = '#0f62fe' },
    @{ value = "$totalA7 ($( & $pctOf $totalA7 $totalMbx ))"; label = 'Active (7 days)'; color = '#24a148' },
    @{ value = "$totalA30 ($( & $pctOf $totalA30 $totalMbx ))"; label = 'Active (30 days)'; color = '#24a148' },
    @{ value = "$totalA90 ($( & $pctOf $totalA90 $totalMbx ))"; label = 'Active (90 days)'; color = '#24a148' }
)
Export-StandardHtmlReport -OutputPath $HtmlPath -Title 'Group Mailbox Activity' -Subtitle $ReportDescription `
    -Body $tableHtml -Kpis $kpis -Version '1.0.0' -ReportName 'Group Mailbox Activity'
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
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { Write-Warning "Ignored error: $($_.Exception.Message)" }

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

    & $w "Sending email summary..." 'INFO'
    $smtp.Send($mail)
    & $w "Email summary sent to: $MailTo" 'SUCCESS'
}
catch {
    & $w ("[ERROR] SMTP send failed: {0}" -f $_.Exception.Message) 'ERROR'
}
finally {
    if ($mail) { $mail.Dispose() }
}

# ============================================================================
# HTML REPORT HELPERS (embedded canonical EnterpriseHtmlReport.template.ps1 v1.0.1).
# Six functions copied verbatim (Get-StandardHtmlHead/Open/Footer/Close/ChartScripts + Export-StandardHtmlReport);
# file-level header omitted. IBM Carbon Dark is the only approved HTML design system.
# ============================================================================

# ============================================================================
# Get-StandardHtmlHead - emits <head> with Carbon design tokens + base styles.
# ============================================================================
function Get-StandardHtmlHead {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter()][string]$Subtitle = ''
    )

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$Title $Subtitle</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

:root {
    --cds-background: #161616;
    --cds-layer-01: #262626;
    --cds-layer-02: #353535;
    --cds-border-strong-01: #4d4d4d;
    --cds-border-subtle-01: #393939;
    --cds-text-primary: #f4f4f4;
    --cds-text-secondary: #c6c6c6;
    --cds-text-helper: #8d8d8d;
    --cds-link: #78a9ff;
    --cds-blue: #0f62fe;
    --cds-purple: #8a3ffc;
    --cds-magenta: #d02670;
    --cds-support-success: #24a148;
    --cds-support-warning: #f1c21b;
    --cds-support-error: #da1e28;
    --cds-support-info: #0043ce;
}

* { margin: 0; padding: 0; box-sizing: border-box; }

body {
    font-family: 'IBM Plex Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background-color: var(--cds-background);
    color: var(--cds-text-primary);
    line-height: 1.4;
    padding: 32px;
    -webkit-font-smoothing: antialiased;
}

.header {
    margin-bottom: 40px;
    padding-bottom: 24px;
    border-bottom: 1px solid var(--cds-border-strong-01);
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    flex-wrap: wrap;
    gap: 16px;
}

.header-left h1 {
    font-size: 28px;
    font-weight: 300;
    letter-spacing: 0.5px;
    color: var(--cds-text-primary);
    margin-bottom: 4px;
}

.header-left h1 strong { font-weight: 600; }

.header .subtitle {
    color: var(--cds-text-secondary);
    font-size: 14px;
    font-family: 'IBM Plex Mono', monospace;
}

.header-right {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12px;
    color: var(--cds-text-helper);
    text-align: right;
}

.kpi-row {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 2px;
    background-color: var(--cds-border-subtle-01);
    border: 1px solid var(--cds-border-subtle-01);
    margin-bottom: 40px;
}

.kpi-card {
    background-color: var(--cds-layer-01);
    padding: 20px;
    display: flex;
    flex-direction: column-reverse;
    justify-content: space-between;
    min-height: 120px;
    transition: background-color 0.15s ease;
}

.kpi-card:hover { background-color: var(--cds-layer-02); }

.kpi-value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 38px;
    font-weight: 400;
    line-height: 1.1;
    color: var(--cds-text-primary);
}

.kpi-label {
    font-size: 12px;
    font-weight: 500;
    color: var(--cds-text-secondary);
    letter-spacing: 0.2px;
    margin-bottom: 12px;
}

.section-title {
    font-size: 14px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--cds-text-secondary);
    margin: 40px 0 16px 0;
    padding-bottom: 8px;
    border-bottom: 1px solid var(--cds-border-subtle-01);
    display: flex;
    align-items: center;
    gap: 8px;
}

.grid-2 {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(450px, 1fr));
    gap: 24px;
    margin-bottom: 24px;
}

.card {
    background-color: var(--cds-layer-01);
    border: 1px solid var(--cds-border-subtle-01);
    padding: 24px;
    display: flex;
    flex-direction: column;
}

.card h2 {
    font-size: 16px;
    font-weight: 400;
    margin-bottom: 24px;
    color: var(--cds-text-primary);
    border-left: 3px solid var(--cds-blue);
    padding-left: 12px;
}

.legend { list-style: none; flex: 1; min-width: 180px; }
.legend li {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 8px 0;
    font-size: 13px;
    border-bottom: 1px solid var(--cds-border-subtle-01);
}
.legend li:last-child { border-bottom: none; }
.legend .dot { width: 8px; height: 8px; flex-shrink: 0; }
.legend .count {
    margin-left: auto;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 500;
}

.bar-chart { display: flex; flex-direction: column; gap: 12px; }
.bar-row { display: flex; align-items: center; gap: 16px; font-size: 13px; }
.bar-label {
    min-width: 140px;
    text-align: right;
    color: var(--cds-text-secondary);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.bar-track { flex: 1; height: 20px; background-color: var(--cds-border-subtle-01); position: relative; }
.bar-fill {
    height: 100%;
    transition: width 0.8s cubic-bezier(0.16, 1, 0.3, 1);
    display: flex;
    align-items: center;
    padding-left: 8px;
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    color: #ffffff;
    min-width: 24px;
}

/* Donut chart layout (used by Get-StandardHtmlChartScripts) */
.donut-container {
    display: flex;
    align-items: center;
    justify-content: space-around;
    gap: 24px;
    flex-wrap: wrap;
}

.donut-wrap {
    position: relative;
    width: 160px;
    height: 160px;
}

.donut-center {
    position: absolute;
    top: 50%; left: 50%;
    transform: translate(-50%, -50%);
    text-align: center;
}

.donut-center .grade {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 36px;
    font-weight: 500;
    line-height: 1;
}

.donut-center .rate {
    font-size: 11px;
    color: var(--cds-text-helper);
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-top: 4px;
}

table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}

thead th {
    text-align: left;
    padding: 12px 16px;
    background-color: var(--cds-layer-02);
    border-bottom: 1px solid var(--cds-border-strong-01);
    color: var(--cds-text-primary);
    font-weight: 500;
    font-size: 12px;
}

tbody td {
    padding: 12px 16px;
    border-bottom: 1px solid var(--cds-border-subtle-01);
    color: var(--cds-text-secondary);
}

tbody tr {
    background-color: var(--cds-layer-01);
    transition: background-color 0.1s ease;
}

tbody tr:hover { background-color: var(--cds-layer-02); }

.badge {
    display: inline-block;
    padding: 2px 8px;
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.badge.critical { background-color: rgba(218, 30, 40, 0.15); color: #ff8389; border-left: 3px solid var(--cds-support-error); }
.badge.high     { background-color: rgba(219, 109, 40, 0.15); color: #ffb38a; border-left: 3px solid #db6d28; }
.badge.medium   { background-color: rgba(241, 194, 27, 0.15); color: #f1c21b; border-left: 3px solid var(--cds-support-warning); }
.badge.low      { background-color: rgba(36, 161, 72, 0.15); color: #8ee0a5; border-left: 3px solid var(--cds-support-success); }
.badge.info     { background-color: rgba(15, 98, 254, 0.15); color: #78a9ff; border-left: 3px solid var(--cds-support-info); }

.progress-bar {
    display: inline-block;
    width: 100px;
    height: 8px;
    background-color: var(--cds-border-subtle-01);
    vertical-align: middle;
}
.progress-fill { height: 100%; transition: width 0.5s ease; }
.progress-fill.low      { background-color: var(--cds-support-success); }
.progress-fill.medium   { background-color: var(--cds-support-warning); }
.progress-fill.high     { background-color: #db6d28; }
.progress-fill.critical { background-color: var(--cds-support-error); }
.pct-label {
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    margin-left: 8px;
    vertical-align: middle;
    color: var(--cds-text-secondary);
}

.empty-state {
    text-align: center;
    padding: 40px;
    color: var(--cds-text-helper);
    font-style: normal;
    border: 1px dashed var(--cds-border-strong-01);
    background-color: var(--cds-background);
}

canvas { max-width: 100%; }

.footer {
    margin-top: 80px;
    padding: 32px;
    background-color: var(--cds-layer-01);
    border: 1px solid var(--cds-border-subtle-01);
    display: grid;
    grid-template-columns: 1.2fr 0.8fr 1fr;
    gap: 32px;
    align-items: start;
}

.footer-col { display: flex; flex-direction: column; gap: 8px; }
.footer-label {
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: var(--cds-text-helper);
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 500;
    margin-bottom: 4px;
}
.footer-value {
    font-size: 13px;
    color: var(--cds-text-primary);
    font-family: 'IBM Plex Mono', monospace;
    line-height: 1.6;
    word-break: break-all;
}
.footer-value strong { color: var(--cds-text-primary); font-weight: 500; }

.grade-card {
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 16px;
    background-color: var(--cds-layer-02);
    border: 1px solid var(--cds-border-subtle-01);
    cursor: help;
    position: relative;
    transition: border-color 0.15s ease;
}
.grade-card:hover { border-color: var(--cds-blue); }
.grade-card .grade-letter {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 64px;
    font-weight: 300;
    line-height: 1;
    margin-bottom: 8px;
}
.grade-card .grade-rate { font-size: 12px; font-family: 'IBM Plex Mono', monospace; color: var(--cds-text-secondary); letter-spacing: 0.5px; }
.grade-card .grade-label {
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: var(--cds-text-helper);
    margin-top: 12px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 500;
}
.grade-card[data-tooltip]:hover::after {
    content: attr(data-tooltip);
    position: absolute;
    bottom: calc(100% + 8px);
    left: 50%;
    transform: translateX(-50%);
    background-color: var(--cds-layer-02);
    border: 1px solid var(--cds-blue);
    color: var(--cds-text-primary);
    padding: 12px 16px;
    font-size: 11px;
    font-family: 'IBM Plex Sans', sans-serif;
    text-transform: none;
    letter-spacing: normal;
    line-height: 1.5;
    width: 240px;
    text-align: left;
    z-index: 10;
    pointer-events: none;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4);
    white-space: pre-wrap;
}

.action-bar { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 4px; }

.btn {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 8px 14px;
    font-size: 12px;
    font-family: 'IBM Plex Sans', sans-serif;
    font-weight: 500;
    background-color: var(--cds-layer-02);
    color: var(--cds-text-primary);
    border: 1px solid var(--cds-border-strong-01);
    cursor: pointer;
    transition: background-color 0.15s ease, border-color 0.15s ease;
    text-decoration: none;
}
.btn:hover { background-color: var(--cds-layer-01); border-color: var(--cds-blue); }
.btn-primary { background-color: var(--cds-blue); color: #ffffff; border-color: var(--cds-blue); }
.btn-primary:hover { background-color: #0353e9; border-color: #0353e9; }
.btn-icon { width: 14px; height: 14px; flex-shrink: 0; }

.footer-meta {
    margin-top: 24px;
    padding-top: 20px;
    border-top: 1px solid var(--cds-border-subtle-01);
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
    font-size: 11px;
    color: var(--cds-text-helper);
    font-family: 'IBM Plex Mono', monospace;
    grid-column: 1 / -1;
}
.footer-meta a { color: var(--cds-link); text-decoration: none; }
.footer-meta a:hover { text-decoration: underline; }

.disclaimer-box {
    position: fixed;
    inset: 0;
    background-color: rgba(0, 0, 0, 0.7);
    display: none;
    align-items: center;
    justify-content: center;
    z-index: 100;
    padding: 32px;
}
.disclaimer-box.is-open { display: flex; }
.disclaimer-modal {
    background-color: var(--cds-layer-01);
    border: 1px solid var(--cds-border-strong-01);
    max-width: 560px;
    width: 100%;
    padding: 32px;
    max-height: 80vh;
    overflow-y: auto;
}
.disclaimer-modal h3 {
    font-size: 16px;
    font-weight: 600;
    margin-bottom: 16px;
    color: var(--cds-text-primary);
    text-transform: uppercase;
    letter-spacing: 1px;
}
.disclaimer-modal p {
    font-size: 13px;
    line-height: 1.6;
    color: var(--cds-text-secondary);
    margin-bottom: 16px;
}
.disclaimer-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 24px; }

@media (max-width: 768px) {
    .grid-2 { grid-template-columns: 1fr; }
    .kpi-row { grid-template-columns: repeat(2, 1fr); }
    .footer { grid-template-columns: 1fr; gap: 24px; }
    body { padding: 16px; }
}

@media print {
    body { background-color: #ffffff; color: #000000; padding: 12px; }
    .header, .footer, .kpi-card, .card { background-color: #ffffff !important; border-color: #d0d0d0 !important; break-inside: avoid; }
    .footer { display: none; }
    .section-title { color: #000000; border-bottom-color: #000000; }
    thead th { background-color: #f4f4f4; color: #000000; }
    .no-print { display: none; }
}
</style>
</head>
<body>
"@
}

# ============================================================================
# Get-StandardHtmlOpen - opens <body> with header bar + KPI tiles.
# Pass a hashtable of @{value=...; label=...; color=...} for each KPI.
# ============================================================================
function Get-StandardHtmlOpen {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter()][string]$Subtitle = '',
        [Parameter()][string]$GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        [Parameter()][string]$Operator = '',
        [Parameter()][hashtable[]]$Kpis = @()
    )

    $kpiHtml = ''
    foreach ($k in $Kpis) {
        $color = if ($k.ContainsKey('color') -and $k['color']) { "color:$($k['color'])" } else { '' }
        $kpiHtml += "<div class=`"kpi-card`"><div class=`"kpi-value`" style=`"$color`">$($k['value'])</div><div class=`"kpi-label`">$($k['label'])</div></div>`n"
    }

    $operatorRow = if ($Operator) { "<div style=`"margin-top: 4px;`">OPERATOR: $Operator</div>" } else { '' }

    return @"
<div class="header">
    <div class="header-left">
        <h1>$Title</h1>
        <div class="subtitle">$Subtitle</div>
    </div>
    <div class="header-right">
        <div>GENERATED: $GeneratedAt</div>
        $operatorRow
    </div>
</div>

<!-- KPI Cards -->
<div class="kpi-row">
$kpiHtml</div>
"@
}

# ============================================================================
# Get-StandardHtmlFooter - emits the canonical 3-column footer + modal.
# ============================================================================
function Get-StandardHtmlFooter {
    [CmdletBinding()]
    param(
        [Parameter()][string]$Tenant = '',
        [Parameter()][string]$Operator = '',
        [Parameter()][string]$GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        [Parameter()][string]$Timezone = [System.TimeZoneInfo]::Local.DisplayName,
        [Parameter()][string]$Utc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"),
        [Parameter()][string]$RunId = ([guid]::NewGuid().ToString().Substring(0, 8).ToUpper()),
        [Parameter()][string]$Version = '1.0.0',
        [Parameter()][string]$ReportName = 'Report',
        [Parameter()][string]$Grade = '',
        [Parameter()][string]$GradeRate = '',
        [Parameter()][string]$GradeColor = '',
        [Parameter()][string]$GradeTip = ''
    )

    $gradeBlock = if ($Grade) {
        $tipAttr = if ($GradeTip) { " data-tooltip=`"$GradeTip`"" } else { '' }
        @"
    <div class="footer-col">
        <div class="grade-card"$tipAttr>
            <div class="grade-letter" style="color:$GradeColor">$Grade</div>
            <div class="grade-rate">$GradeRate</div>
            <div class="grade-label">Compliance Grade</div>
        </div>
    </div>
"@
    } else {
        '<div class="footer-col"></div>'
    }

    return @"
<!-- Footer -->
<div class="footer">
    <div class="footer-col">
        <div class="footer-label">Tenant</div>
        <div class="footer-value"><strong>$Tenant</strong></div>
        <div class="footer-label" style="margin-top:12px;">Operator</div>
        <div class="footer-value">$Operator</div>
        <div class="footer-label" style="margin-top:12px;">Generated</div>
        <div class="footer-value">$GeneratedAt</div>
        <div class="footer-value" style="color:var(--cds-text-helper);font-size:11px;">$Timezone &middot; UTC $Utc</div>
    </div>
$gradeBlock
    <div class="footer-col">
        <div class="footer-label">Run</div>
        <div class="footer-value">$RunId</div>
        <div class="footer-label" style="margin-top:12px;">Quick Actions</div>
        <div class="action-bar">
            <button class="btn btn-primary" onclick="window.print()" title="Print or save as PDF">
                <svg class="btn-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M4 2h8v3H4V2zm0 5h8a2 2 0 0 1 2 2v3h-2v3H4v-3H2V9a2 2 0 0 1 2-2zm1 7v-3h6v3H5z"/></svg>
                Print
            </button>
            <button class="btn" onclick="navigator.clipboard.writeText(window.location.href)" title="Copy file path">
                <svg class="btn-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M5 2h7a1 1 0 0 1 1 1v9h-2V4H6v8H4V3a1 1 0 0 1 1-1zM2 5h8a1 1 0 0 1 1 1v8a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1zm1 2v6h6V7H3z"/></svg>
                Copy Path
            </button>
            <button class="btn" onclick="document.querySelector('.header').scrollIntoView({behavior:'smooth'})" title="Back to top">
                <svg class="btn-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M8 3.5l5 5h-3v4H6v-4H3l5-5z"/></svg>
                Top
            </button>
        </div>
        <div class="action-bar" style="margin-top:8px;">
            <button class="btn" onclick="document.getElementById('disclaimerModal').classList.add('is-open')" title="View full disclaimer">
                <svg class="btn-icon" viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 1 0 0 14A7 7 0 0 0 8 1zm0 3a1 1 0 0 1 1 1v4a1 1 0 0 1-2 0V5a1 1 0 0 1 1-1zm0 8a1 1 0 1 1 0-2 1 1 0 0 1 0 2z"/></svg>
                Disclaimer
            </button>
        </div>
    </div>
    <div class="footer-meta">
        <span>$ReportName &middot; v$Version &middot; Run $RunId</span>
        <span>Generated by <a href="https://github.com/mabdulkadr/powershell-enterprise-admin-skill" target="_blank" rel="noopener">PowerShell Enterprise Admin</a></span>
    </div>
</div>

<!-- Disclaimer modal -->
<div class="disclaimer-box" id="disclaimerModal" onclick="if(event.target===this)this.classList.remove('is-open')">
    <div class="disclaimer-modal">
        <h3>Disclaimer</h3>
        <p>This report is generated from read-only queries and is provided as-is with no warranty of any kind. The metrics and identifiers shown are a point-in-time snapshot and may not reflect the current state by the time this report is reviewed.</p>
        <p>Test generated tools in a staging environment before deploying to production. The authors assume no liability for any damage or data loss resulting from their use.</p>
        <p>This report may contain tenant identifiers and operator account names. Treat the file as confidential and follow your organization's data-handling policy when sharing.</p>
        <div class="disclaimer-actions">
            <button class="btn" onclick="document.getElementById('disclaimerModal').classList.remove('is-open')">Close</button>
        </div>
    </div>
</div>
"@
}

# ============================================================================
# Get-StandardHtmlClose - emits </body></html> + the standard JS helpers.
# ============================================================================
function Get-StandardHtmlClose {
    [CmdletBinding()]
    param()

    return @"
<script>
// Close disclaimer modal on Escape
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        const m = document.getElementById('disclaimerModal');
        if (m) m.classList.remove('is-open');
    }
});
</script>
</body>
</html>
"@
}

# ============================================================================
# Get-StandardHtmlChartScripts - optional helper for scripts that draw donuts/bars.
# ============================================================================
function Get-StandardHtmlChartScripts {
    [CmdletBinding()]
    param()

    return @"
<script>
// Mini donut chart renderer (no dependencies)
function drawDonut(canvasId, data, colors) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const cx = canvas.width / 2, cy = canvas.height / 2;
    const outerR = 76, innerR = 58;
    const total = data.reduce((a, b) => a + b, 0);
    if (total === 0) return;
    let startAngle = -Math.PI / 2;
    data.forEach((val, i) => {
        const sliceAngle = (val / total) * 2 * Math.PI;
        ctx.beginPath();
        ctx.arc(cx, cy, outerR, startAngle, startAngle + sliceAngle);
        ctx.arc(cx, cy, innerR, startAngle + sliceAngle, startAngle, true);
        ctx.closePath();
        ctx.fillStyle = colors[i];
        ctx.fill();
        startAngle += sliceAngle;
    });
}

// Bar chart renderer
function drawBars(containerId, data, colorFn) {
    const container = document.getElementById(containerId);
    if (!container || !data) return;
    const dataArray = Array.isArray(data) ? data : [data];
    if (!dataArray.length || (dataArray.length === 1 && !dataArray[0])) return;
    const maxVal = Math.max(...dataArray.map(d => d.value || 0));
    const colors = ['#0f62fe','#8a3ffc','#00b0ff','#008080','#da1e28','#ff832b','#8d8d8d','#e0e0e0'];
    dataArray.forEach((item, i) => {
        if (!item) return;
        const val = item.value || 0;
        const label = item.label || 'Unknown';
        const pct = maxVal > 0 ? (val / maxVal * 100) : 0;
        const color = colorFn ? colorFn(item, i) : colors[i % colors.length];
        const row = document.createElement('div');
        row.className = 'bar-row';
        row.innerHTML =
            '<div class="bar-label" title="' + label + '">' + label + '</div>' +
            '<div class="bar-track"><div class="bar-fill" style="width:' + pct + '%;background-color:' + color + '">' + val + '</div></div>';
        container.appendChild(row);
    });
}
</script>
"@
}

# ============================================================================
# Export-StandardHtmlReport - convenience: build + write a complete report in one call.
# Pass -Body as the HTML between header and footer.
# ============================================================================
function Export-StandardHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter()][string]$Subtitle = '',
        [Parameter()][string]$Tenant = '',
        [Parameter()][string]$Operator = '',
        [Parameter()][string]$Body = '',
        [Parameter()][hashtable[]]$Kpis = @(),
        [Parameter()][string]$Grade = '',
        [Parameter()][string]$GradeRate = '',
        [Parameter()][string]$GradeColor = '',
        [Parameter()][string]$GradeTip = '',
        [Parameter()][string]$Version = '1.0.0',
        [Parameter()][string]$ReportName = 'Report',
        [Parameter()][string]$ChartScripts = ''
    )

    $now = Get-Date
    $html = Get-StandardHtmlHead -Title $Title -Subtitle $Subtitle
    $html += Get-StandardHtmlOpen -Title $Title -Subtitle $Subtitle -GeneratedAt $now.ToString('yyyy-MM-dd HH:mm:ss') -Operator $Operator -Kpis $Kpis
    $html += $Body
    $html += Get-StandardHtmlFooter -Tenant $Tenant -Operator $Operator -GeneratedAt $now.ToString('yyyy-MM-dd HH:mm:ss') `
        -Timezone ([System.TimeZoneInfo]::Local.DisplayName) `
        -Utc $now.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") `
        -RunId ([guid]::NewGuid().ToString().Substring(0, 8).ToUpper()) `
        -Version $Version -ReportName $ReportName `
        -Grade $Grade -GradeRate $GradeRate -GradeColor $GradeColor -GradeTip $GradeTip
    if ($ChartScripts) { $html += "`n$ChartScripts" }
    $html += Get-StandardHtmlClose

    $html | Out-File -FilePath $OutputPath -Encoding utf8 -Force
}
