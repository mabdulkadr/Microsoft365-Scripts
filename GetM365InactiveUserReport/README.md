<div align="center">

# 💻 GetM365InactiveUserReport

**Inactive Microsoft 365 user report from Graph sign-in activity.**

Interactive + non-interactive sign-in aging with enabled/disabled/external/never-logged-in slices.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**GetM365InactiveUserReport** is a PowerShell script that reports inactive Microsoft 365 users from Microsoft Graph sign-in activity. It ages both interactive and non-interactive sign-ins, slices by account state, and exports UTF-8 CSV beside the script. Scheduler-friendly for automated hygiene runs.

---

# ✨ Features

* Dual aging: interactive and non-interactive sign-ins
* Slices: enabled / disabled / external (`#EXT#`) / never-logged-in
* Certificate app-auth support for unattended runs
* UTF-8 CSV + Carbon Dark HTML dashboard (shared run) with UPN, dates, inactive-day counts, status, department, title

---

# 📂 Project Structure

```text
GetM365InactiveUserReport
│
├── GetM365InactiveUserReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\GetM365InactiveUserReport.ps1 -InactiveDays 30 -EnabledUsersOnly
```

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `InactiveDays` | Int | No | (off) | Inactivity threshold on interactive sign-ins. |
| `InactiveDays_NonInteractive` | Int | No | (off) | Inactivity threshold on non-interactive sign-ins. |
| `ReturnNeverLoggedInUser` | Switch | No | off | Only users who never logged in. |
| `EnabledUsersOnly` | Switch | No | off | Only enabled users. |
| `DisabledUsersOnly` | Switch | No | off | Only disabled users. |
| `ExternalUsersOnly` | Switch | No | off | Only external (`#EXT#`) users. |
| `CreateSession` | Switch | No | off | Disconnect existing Graph sessions first. |
| `TenantId` | String | No | (interactive) | Tenant ID for certificate auth. |
| `ClientId` | String | No | (interactive) | App ID for certificate auth. |
| `CertificateThumbprint` | String | No | (interactive) | Certificate thumbprint for app auth. |

### Exit Codes
| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

---

# ⚙️ Requirements

### Operating System
* Windows 10 / Windows 11

### PowerShell
* PowerShell **5.1 or later**
* Microsoft Graph modules (prompted install if missing)

### Permissions
* `User.Read.All`, `AuditLog.Read.All` (Microsoft Graph).

---

# 🛡 Operational Notes
* CSV + HTML land beside the script with UPN, creation date, last sign-ins, inactive-day counts, status, department, employee ID/name, title.
* For unattended runs supply `TenantId` + `ClientId` + `CertificateThumbprint`.
* Reference: [o365reports.com — Inactive User Report](https://o365reports.com/2023/06/21/microsoft-365-inactive-user-report-ms-graph-powershell/).

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
