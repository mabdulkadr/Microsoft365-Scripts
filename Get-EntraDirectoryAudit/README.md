<div align="center">

# 📊 Get-EntraDirectoryAudit

**Reports recent Entra ID directory changes from audit logs.**

Pulls directory audit logs showing recent changes for users, groups, applications, role assignments, and Conditional Access policies with actor details and timestamps for governance review.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.1-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Structure](#-project-structure) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get-EntraDirectoryAudit** is a PowerShell reporting script that pulls directory audit logs showing recent changes for users, groups, applications, role assignments, and Conditional Access policies with actor details and timestamps for governance review.

It queries the Graph auditLogs/directoryAudits endpoint with a configurable lookback window and category filter, then expands each audit entry with actor, target, and result for governance and change-tracking.

---

# ✨ Features

* Pulls directory audit logs with Hours and Category filtering
* Shows actor, target, activity, and result for each change
* Covers users, groups, applications, roles, and Conditional Access
* Handles UTC time filtering and pagination via `Get-MgGraphAllPages`
* Exports CSV + Carbon Dark HTML dashboard (shared timestamp) beside the script for evidence retention

---

# 📂 Project Structure

```text
Get-EntraDirectoryAudit
│
├── Get-EntraDirectoryAudit.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-EntraDirectoryAudit.ps1
```

### Example 1
```powershell
.\Get-EntraDirectoryAudit.ps1
```
Reports directory changes from the last 24 hours.

### Example 2
```powershell
.\Get-EntraDirectoryAudit.ps1 -Hours 168 -Category Role
```
Reports role changes from the last 7 days.

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `Hours` | Int32 | No | 24 | Hours to look back. |
| `Category` | String | No | All | Filter by category: All, User, Group, Application, Role, Policy. |
| `ExportPath` | String | No | Beside script | Optional CSV export path. |

### Exit Codes
| Code | Status |
| ---- | ------ |
| 0 | Success |
| 1 | Failure |

---

# ⚙️ Requirements

### Operating System
* Windows 10 / Windows 11

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* `AuditLog.Read.All, Directory.Read.All`

### Logging
* `C:\ProgramData\Get-EntraDirectoryAudit\Logs\`

---

# 🛡 Operational Notes
* Read-only; never modifies directory objects.
* Audit log retention is tenant-dependent; large Hours windows may be throttled.
* Empty windows render as empty reports without error.

---

## 👤 Author
**Mohammad Abdelkader Omar**  
GitHub: [@mabdulkadr](https://github.com/mabdulkadr)  
Website: [momar.tech](https://momar.tech)
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