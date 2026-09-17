<div align="center">

# 💻 Get-GroupMailboxActivityReport

**Per-group mailbox activity counts for 7/30/90 days (CSV + HTML + email).**

Parallel EXO collection with runspace throttling, styled HTML dashboard, and SMTP summary.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**Get-GroupMailboxActivityReport** is a PowerShell script that measures mailbox activity (`Get-MailboxStatistics.LastLogonTime`) for members of selected Microsoft 365 groups and aggregates active counts for 7/30/90 days. It emits a CSV, a dark card-style HTML dashboard, and a compact HTML email summary with both files attached.

---

# ✨ Features

* Handles M365/Unified, mail-enabled Security, Distribution, and Entra security groups
* Runspace-parallel collection with throttle (`$MaxConcurrency` default 10, own EXO connection per worker)
* CSV + Carbon Dark HTML dashboard (KPI tiles + per-group detail table) + SMTP email summary with attachments
* Mailbox-owner-aware percentages (denominator = users that actually have mailboxes)

---

# 📂 Project Structure

```text
Get-GroupMailboxActivityReport
│
├── Get-GroupMailboxActivityReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-GroupMailboxActivityReport.ps1
```

### Configuration (edit inside the script)
```powershell
$TenantId              = "<tenant-guid>"
$AppId                 = "<app-id>"
$CertificateThumbprint = "<certificate-thumbprint>"
$Organization          = "yourdomain.onmicrosoft.com"

$GroupsToReport = @("group-guid-or-displayname-1", "group-guid-or-displayname-2")

$MailFrom   = "sender@domain.com"
$MailTo     = "recipient1@domain.com,recipient2@domain.com"
$SmtpServer = "smtp.office365.com"
$SmtpPort   = 587
$OutDir     = "Reports"  # beside the script (Law 12); override per environment
```

---

# ⚙️ Parameters

This script takes no command-line parameters. Targets, credentials, SMTP, and output folder are set in the configuration block above.

### Outputs
* **CSV** `Group_Activity_Report.csv` — GroupName, GroupType, ObjectId, Email, Created, Users, Owners, MailboxOwners, Active7, Active30, Active90
* **HTML dashboard** `Group_Activity_Report.html` — IBM Carbon Dark (KPI tiles + detail table)
* **Email summary** — compact HTML table with CSV + HTML attached

### Exit Codes
| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |
| 2    | Script error |

---

# ⚙️ Requirements

### Operating System
* Windows 10 / Windows 11 (64-bit Windows PowerShell 5.1, elevated)

### PowerShell
* PowerShell **5.1 or later**
* `ExchangeOnlineManagement`, `AzureAD` (legacy AAD Graph — deprecated by Microsoft; Graph migration recommended)

### Permissions
* App registration (certificate): `Exchange.ManageAsApp`, `Directory.Read.All` + admin consent; SMTP AUTH enabled for the sender.

---

# 🛡 Operational Notes
* For large groups lower concurrency or run off-hours to avoid throttling.
* `LastLogonTime` semantics vary (shared/resource mailboxes, background access) — interpret counts accordingly.
* Protect the certificate private key and SMTP credentials (Key Vault / Secret Store).
* Module install if missing: `Install-Module ExchangeOnlineManagement, AzureAD -Scope CurrentUser` (plus NuGet provider / trusted PSGallery first).
* Docs: [Exchange Online PowerShell](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell) · [Get-MailboxStatistics](https://learn.microsoft.com/powershell/module/exchange/get-mailboxstatistics) · [SMTP AUTH](https://learn.microsoft.com/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission).

---

## 👤 Author

**Mohammad Abdelkader Omar**
GitHub: [@mabdulkadr](https://github.com/mabdulkadr)
Website: [momar.tech](https://momar.tech)

---

## 📜 License
This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## ⚠ Disclaimer

This skill and every script it generates are provided as-is with no warranty of any kind. Test generated tools in a staging environment before deploying to production. The authors assume no liability for any damage or data loss resulting from their use.

---
<div align="center">

⭐ **If this skill saves you time, star the repo — it helps others find it.**

[Report an Issue](../../issues) · [momar.tech](https://momar.tech)

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/mabdulkadrx)

Built with [**PowerShell Enterprise Admin**](https://github.com/mabdulkadr/powershell-enterprise-admin-skill)

</div>
