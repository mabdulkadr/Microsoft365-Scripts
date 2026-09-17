<div align="center">

# 📊 Get-EntraAppRegistrationAudit

**Audits Entra ID app registrations for expiring credentials and excessive permissions.**

Lists all app registrations and flags expiring or expired secrets and certificates, excessive API permissions, missing owners, stale sign-ins, and multi-tenant apps for security hygiene review.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.1-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Structure](#-project-structure) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get-EntraAppRegistrationAudit** is a PowerShell reporting script that lists all app registrations and flags expiring or expired secrets and certificates, excessive API permissions, missing owners, stale sign-ins, and multi-tenant apps for security hygiene review.

It inventories app registrations, service principals, and credential metadata to surface hygiene gaps — expired or soon-to-expire secrets, over-privileged apps, ownerless apps, and multi-tenant risks. Ideal for quarterly app hygiene reviews.

---

# ✨ Features

* Flags expiring and expired secrets and certificates within the chosen window
* Detects excessive API permissions and high-risk app roles
* Identifies apps with no owner and without recent sign-in activity
* Surfaces multi-tenant and stale apps for cleanup
* Exports CSV + Carbon Dark HTML dashboard (shared timestamp) beside the script via `Get-MgGraphAllPages`

---

# 📂 Project Structure

```text
Get-EntraAppRegistrationAudit
│
├── Get-EntraAppRegistrationAudit.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-EntraAppRegistrationAudit.ps1
```

### Example 1
```powershell
.\Get-EntraAppRegistrationAudit.ps1
```
Audits app registrations with default 30-day expiry window.

### Example 2
```powershell
.\Get-EntraAppRegistrationAudit.ps1 -DaysUntilExpiry 90
```
Audits with a 90-day expiry window.

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `DaysUntilExpiry` | Int32 | No | 30 | Flag credentials expiring within this many days. |
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
* `Application.Read.All, AuditLog.Read.All, Directory.Read.All`

### Logging
* `C:\ProgramData\Get-EntraAppRegistrationAudit\Logs\`

---

# 🛡 Operational Notes
* Read-only; never modifies app registrations.
* Credential expiry is evaluated in tenant local time; verify with Entra portal.
* Permission analysis is heuristic — review high-privilege apps manually.

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