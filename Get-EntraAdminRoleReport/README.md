<div align="center">

# 📊 Get-EntraAdminRoleReport

**Reports Entra ID admin role assignments and privileged access posture.**

Lists all directory role assignments showing who has which admin roles, assignment permanence versus PIM eligibility, stale admin sign-ins, and users with multiple admin roles for privileged access review.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.1-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Structure](#-project-structure) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get-EntraAdminRoleReport** is a PowerShell reporting script that lists all directory role assignments showing who has which admin roles, assignment permanence versus PIM eligibility, stale admin sign-ins, and users with multiple admin roles for privileged access review.

It pulls role definitions, role assignments, and user sign-in activity to surface permanent versus eligible assignments, dormant admin accounts, and multi-role users. Use it for privileged access reviews and PIM hygiene.

---

# ✨ Features

* Lists all directory role assignments with role names resolved
* Distinguishes permanent versus PIM-eligible assignments where available
* Flags admin accounts with stale sign-ins and multiple roles
* Resolves user display names and groups for each assignment
* Exports CSV + Carbon Dark HTML dashboard (shared timestamp) beside the script with structured logging

---

# 📂 Project Structure

```text
Get-EntraAdminRoleReport
│
├── Get-EntraAdminRoleReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-EntraAdminRoleReport.ps1
```

### Example 1
```powershell
.\Get-EntraAdminRoleReport.ps1
```
Generates the admin role report for the tenant.

### Example 2
```powershell
.\Get-EntraAdminRoleReport.ps1 -ExportPath "C:\Reports\AdminRoles.csv"
```
Exports the report to a specific CSV path.

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
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
* `RoleManagement.Read.All, Directory.Read.All, User.Read.All, AuditLog.Read.All`

### Logging
* `C:\ProgramData\Get-EntraAdminRoleReport\Logs\`

---

# 🛡 Operational Notes
* Read-only; never modifies role assignments.
* PIM eligibility requires PIM configuration; otherwise assignments appear permanent.
* Stale sign-in thresholds are tenant-specific; verify with signInActivity data.

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