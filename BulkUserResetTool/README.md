<div align="center">

# 💻 BulkUserResetTool

**Bulk user access control for Microsoft 365: reset passwords, revoke sessions, block sign-in, disable devices.**

One switch-driven script for incident response and offboarding at scale.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**BulkUserResetTool** is a PowerShell script that performs bulk access-control actions in Microsoft 365 via Microsoft Graph: password resets (with optional force-change), session revocation, sign-in blocking, and device disabling — for all users or a targeted list, with exclusions.

---

# ✨ Features

* Password reset with force-change-at-next-sign-in option
* Session + refresh-token revocation per user
* Sign-in blocking and device disabling
* `-All` fleet mode or `-UserPrincipalNames` targeting, minus `-Exclude`

---

# 📂 Project Structure

```text
BulkUserResetTool
│
├── BulkUserResetTool.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\BulkUserResetTool.ps1 -All -SignOut -BlockSignIn
```

### With Parameters
```powershell
.\BulkUserResetTool.ps1 -UserPrincipalNames user1@domain.com, user2@domain.com -ResetPassword
.\BulkUserResetTool.ps1 -All -DisableDevices -SignOut
.\BulkUserResetTool.ps1 -All -BlockSignIn -Exclude admin@domain.com, support@domain.com
```

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `All` | Switch | No | off | Target every user in the directory. |
| `ResetPassword` | Switch | No | off | Reset passwords for targeted users. |
| `DisableDevices` | Switch | No | off | Disable devices registered to targeted users. |
| `SignOut` | Switch | No | off | Revoke all sessions and refresh tokens. |
| `BlockSignIn` | Switch | No | off | Block sign-ins for targeted users. |
| `Exclude` | String[] | No | (none) | UPNs to skip in bulk runs. |
| `UserPrincipalNames` | String[] | No | (none) | Explicit UPN targets. |

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
* `Microsoft.Graph` module (`Install-Module Microsoft.Graph -Scope CurrentUser`), connected with `Directory.AccessAsUser.All`

### Permissions
* Entra ID rights to reset passwords, revoke sessions, block sign-ins, and disable devices.

---

# 🛡 Operational Notes
* High-impact tool: always test in a non-production tenant first.
* Outputs per-user success/error lines plus exclusion confirmations.
* Changelog: V1.00 (2023-06-18) initial release; V1.10 (2023-07-24) Graph PowerShell updates.
* Inspired by: [Force sign out users in Microsoft 365](https://www.alitajran.com/force-sign-out-users-microsoft-365/) (Ali Tajran).

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
