<div align="center">

# 💻 Get-M365UserLastActivityReport

**Real last-logon report for Office 365 users via mailbox statistics.**

Finds stale mailboxes by `LastLogonTime` with filters for mailbox type, license, and never-logged-in accounts.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**Get-M365UserLastActivityReport** is a PowerShell script that reports each Office 365 user's real last logon time from Exchange Online mailbox statistics and exports it to CSV + Carbon Dark HTML dashboard. It auto-installs the Exchange Online and MSOnline modules if missing and authenticates to Azure AD + Exchange Online.

---

# ✨ Features

* Per-mailbox `LastLogonTime` collection via Exchange Online
* Inactivity threshold filter (`-InactiveDays`)
* Never-logged-in-only mode for provisioning cleanup
* Mailbox-type, license, and single-user targeting switches

---

# 📂 Project Structure

```text
Get-M365UserLastActivityReport
│
├── Get-M365UserLastActivityReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-M365UserLastActivityReport.ps1 -InactiveDays 90
```

### With Parameters
```powershell
.\Get-M365UserLastActivityReport.ps1 -InactiveDays 90 -UserMailboxOnly -LicensedUserOnly
```

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `MBNamesFile` | String | No | (all) | File with mailbox names to scope the report. |
| `InactiveDays` | Int | No | (off) | Only mailboxes inactive longer than N days. |
| `UserMailboxOnly` | Switch | No | off | Only user mailboxes (skip shared/resource). |
| `LicensedUserOnly` | Switch | No | off | Only licensed users. |
| `ReturnNeverLoggedInMBOnly` | Switch | No | off | Only mailboxes that never logged in. |
| `UserName` | String | No | (prompt) | Admin account for authentication. |
| `Password` | String | No | (prompt) | Admin password (prefer interactive/MFA instead). |
| `FriendlyTime` | Switch | No | off | Human-friendly time formatting. |
| `NoMFA` | Switch | No | off | Basic-auth path for non-MFA accounts. |

### Exit Codes
| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |
| 2    | Script error |

---

# ⚙️ Requirements

### Operating System
* Windows 10 / Windows 11

### PowerShell
* PowerShell **5.1 or later**
* Exchange Online + MSOnline modules (auto-installed if missing)

### Permissions
* Exchange Online admin (mailbox statistics read).

---

# 🛡 Operational Notes
* Prefer MFA/interactive sign-in over `-UserName`/`-Password` + `-NoMFA`.
* `LastLogonTime` semantics vary by mailbox type — shared/resource mailboxes can mislead; combine with `-UserMailboxOnly`.
* Large tenants take a while; scope with `-MBNamesFile` for targeted runs.

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
