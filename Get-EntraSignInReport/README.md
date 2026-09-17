<div align="center">

# 📊 Get Entra Sign-In Report

**Analyze Entra ID sign-in logs for security issues and authentication failures.**

This script pulls recent sign-in logs from Microsoft Graph and reports failed sign-ins, MFA failures, legacy authentication usage, conditional access failures, and risky sign-ins for security audits.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.1.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Structure](#-project-structure) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get Entra Sign-In Report** is a PowerShell reporting script that Pulls recent sign-in logs and reports failed sign-ins, MFA failures, legacy authentication usage, conditional access failures, sign-ins from unusual locations, top failing users, and top failing apps. Essential for security audits and incident investigation. Requires Entra ID P1/P2 for sign-in log access; data retained 30 days maximum. It runs from a workstation via Microsoft Graph and writes structured logs for every operation.

It is part of the **Intune Reporting Tools** category and runs from a workstation — no agent deployment required.

---

# ✨ Features

* Queries Microsoft Graph sign-in logs with paging and filtering
* Flags MFA failures, legacy auth, and CA policy failures
* Breaks down top failing users, apps, and error codes
* Highlights risky sign-ins and unusual locations
* Exports structured CSV + Carbon Dark HTML dashboard (shared timestamp) for security audits

---

# 📂 Project Structure

```text
Get-EntraSignInReport
│
├── Get-EntraSignInReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-EntraSignInReport.ps1
```

*Last 24 hours of failed sign-ins*

### Example 2
```powershell
.\Get-EntraSignInReport.ps1 -Hours 168 -IncludeSuccessful
```
Last 7 days including successes

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `Hours` | int | No | 24 | Number of hours to look back. Default: 24. Max: 720 (30 days). |
| `IncludeSuccessful` | switch | No | false | Include successful sign-ins (off by default to focus on failures). |
| `ExportPath` | string | No | - | Optional. Export to CSV at the specified path. |

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
* PowerShell **5.1 or later** (`#Requires -Version 5.1`)

### Permissions
* `AuditLog.Read.All,Directory.Read.All`

### Logging
* `C:\ProgramData\get-entra-sign-in-report\Logs`

---

# 🛡 Operational Notes
* Requires Entra ID P1/P2 for sign-in log access; retention is 30 days maximum.
* Filtering by Hours and IncludeSuccessful controls result volume; large windows may hit Graph throttling.
* Test in a staging tenant first; Graph permission errors surface as 403 — check Entra consent for the listed scopes.

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