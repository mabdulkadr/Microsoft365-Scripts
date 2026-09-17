<div align="center">

# 📊 Get-EntraGuestUserAudit

**Audits external and guest users in Entra ID for stale access.**

Lists all guest users and flags never signed in, inactive for the configured window, group memberships, and invitation status to support external access review and cleanup.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.1-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Structure](#-project-structure) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get-EntraGuestUserAudit** is a PowerShell reporting script that lists all guest users and flags never signed in, inactive for the configured window, group memberships, and invitation status to support external access review and cleanup.

It enumerates guest users with sign-in activity, group memberships, and externalUserState to surface never-signed-in guests, 90+ day inactive guests, and pending invitations for access review and offboarding.

---

# ✨ Features

* Lists all guest users with sign-in and invitation state
* Flags never signed in and 90+ day inactive guests
* Resolves guest group memberships where applicable
* Surfaces pending acceptance and disabled guest accounts
* Exports CSV + Carbon Dark HTML dashboard (shared timestamp) beside the script for cleanup workflows

---

# 📂 Project Structure

```text
Get-EntraGuestUserAudit
│
├── Get-EntraGuestUserAudit.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-EntraGuestUserAudit.ps1
```

### Example 1
```powershell
.\Get-EntraGuestUserAudit.ps1
```
Audits guest users with default 90-day inactivity window.

### Example 2
```powershell
.\Get-EntraGuestUserAudit.ps1 -InactiveDays 60
```
Audits with a 60-day inactivity window.

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `InactiveDays` | Int32 | No | 90 | Flag guests inactive for more than this many days. |
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
* `User.Read.All, AuditLog.Read.All, Directory.Read.All, GroupMember.Read.All`

### Logging
* `C:\ProgramData\Get-EntraGuestUserAudit\Logs\`

---

# 🛡 Operational Notes
* Read-only; never modifies guest users.
* Missing sign-in data is treated as never signed in.
* Inactivity is calculated from signInActivity.lastSignInDateTime; verify with Entra logs.

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