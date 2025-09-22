
# Group Mailbox Activity Report
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Version](https://img.shields.io/badge/version-1.0-green.svg)

## Overview
`GroupMailboxActivityReport.ps1` collects mailbox-usage activity for specific Microsoft 365 groups in parallel (no Graph SDK).  
It outputs a CSV, a styled HTML dashboard (per-group “cards”), and also sends a compact HTML email summary.  

The script is designed for **Windows PowerShell 5.1** (ISE-friendly), using **app-only authentication** with certificate-based connections for both Exchange Online and Azure AD.

**Reference Docs (Microsoft):**
- [Exchange Online PowerShell](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell)  
- [Get-MailboxStatistics](https://learn.microsoft.com/powershell/module/exchange/get-mailboxstatistics)  
- [Get-UnifiedGroup](https://learn.microsoft.com/powershell/module/exchange/get-unifiedgroup)  
- [Get-DistributionGroup](https://learn.microsoft.com/powershell/module/exchange/get-distributiongroup)  
- [AzureAD Module](https://learn.microsoft.com/powershell/azure/active-directory/install-adv2)  
- [SMTP AUTH (client submission)](https://learn.microsoft.com/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission)  

---

## Script Included
- **GroupMailboxActivityReport.ps1**  
  Collects mailbox activity for a list of groups, generates CSV + HTML report, and sends an email summary.

---

## Script Details

### Purpose
- Resolve each target group (M365/Unified, Distribution, Security, or Entra ID Security).  
- Enumerate user-like members and detect mailbox owners.  
- Query `Get-MailboxStatistics.LastLogonTime` for mailbox owners.  
- Aggregate activity counts for 7/30/90 days.  
- Export results to **CSV**, render a **styled HTML dashboard**, and send a **summary email**.

### Prerequisites
- **Windows PowerShell 5.1 (x64)**.
- **Modules**:
  - `ExchangeOnlineManagement`
  - `AzureAD` (legacy; used here instead of Microsoft Graph SDK)  
- **App registration** with **certificate-based authentication**:
  - **Exchange Online (Application permission)**: `Exchange.ManageAsApp`
  - **Azure AD Directory (Application permission)**: `Directory.Read.All`
  - Admin consent required.
- **SMTP AUTH** enabled for the sending account (port 587, TLS).

---

### Configuration (edit inside the script)

```powershell
# ================== Configuration ==================
$TenantId              = "<tenant-guid>"
$AppId                 = "<app-id>"
$CertificateThumbprint = "<certificate-thumbprint>"
$Organization          = "yourdomain.onmicrosoft.com"   # or custom domain

# List of groups (display names or ObjectId GUIDs)
$GroupsToReport = @(
  "group-guid-or-displayname-1",
  "group-guid-or-displayname-2"
)

# SMTP + Output
$MailFrom   = "sender@domain.com"
$MailTo     = "recipient1@domain.com,recipient2@domain.com"
$SmtpServer = "smtp.office365.com"
$SmtpPort   = 587
$OutDir     = "C:\Reports"
````

---

### How to Run

```powershell
# Run from 64-bit, elevated Windows PowerShell 5.1
PS> .\GroupMailboxActivityReport.ps1
```

---

### Outputs

* **CSV**: `Group_Activity_Report.csv`
  Columns: GroupName, GroupType, ObjectId, Email, Created, Users, Owners, MailboxOwners, Active7, Active30, Active90
* **HTML Dashboard**: `Group_Activity_Report.html`

  * Dark “card” layout with per-group activity + overall totals.
* **Email Summary**:

  * Compact HTML table embedded in the email.
  * CSV and HTML report attached.

---

## Performance Notes

* Uses a **runspace throttle** (`$MaxConcurrency` default = 10).
* Each worker opens its own Exchange Online + AzureAD app-only connections.
* For large groups, lower concurrency or run during off-hours to avoid throttling.

---

## Troubleshooting

* **Modules not found**:

  * Install NuGet provider

    ```powershell
    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    ```
  * Install required modules

    ```powershell
    Install-Module ExchangeOnlineManagement -Scope CurrentUser
    Install-Module AzureAD -Scope CurrentUser
    ```
* **SMTP send fails**: Check SMTP AUTH, port 587, and account credentials.
* **No activity counts**: Some mailbox types don’t update `LastLogonTime` reliably (shared/resource).
* **Certificate issues**: Ensure cert with private key is in `CurrentUser\My` or `LocalMachine\My`.

---

## Security Considerations

* Protect certificate private key and SMTP credentials.
* Rotate credentials regularly.
* Use secure storage (e.g., Key Vault, Secret Store).

---

## Notes

* AzureAD module is deprecated but included here for compatibility.
  Microsoft Graph PowerShell is recommended for future-proofing.

---

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).
