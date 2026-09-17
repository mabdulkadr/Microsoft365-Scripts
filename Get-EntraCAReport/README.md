<div align="center">

# 📊 Get-EntraCAReport

**Generates a comprehensive Conditional Access policy report from Entra ID.**

Queries Microsoft Graph to retrieve all Conditional Access policies and expands them into a readable report with state, user and group targets, application targets, platform and location conditions, grant controls, and session controls for security auditing and change review.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.1-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Structure](#-project-structure) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get-EntraCAReport** is a PowerShell reporting script that queries Microsoft Graph to retrieve all Conditional Access policies and expands them into a readable report with state, user and group targets, application targets, platform and location conditions, grant controls, and session controls for security auditing and change review.

It resolves group and app IDs to display names, expands conditions and grant/session controls, and supports filtering by name and state. Use it for security audits, documentation, and change-review evidence.

---

# ✨ Features

* Expands policy state, inclusions, exclusions, and application targets
* Resolves user, group, and app IDs to display names
* Details platform, location, risk, and device state conditions
* Shows grant controls (MFA, compliant device) and session controls
* Supports filtering by name and state with CSV + Carbon Dark HTML export (shared timestamp)

---

# 📂 Project Structure

```text
Get-EntraCAReport
│
├── Get-EntraCAReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-EntraCAReport.ps1
```

### Example 1
```powershell
.\Get-EntraCAReport.ps1
```
Exports all enabled and report-only Conditional Access policies.

### Example 2
```powershell
.\Get-EntraCAReport.ps1 -PolicyName "MFA" -IncludeDisabled
```
Exports policies matching MFA including disabled ones.

### Example 3
```powershell
.\Get-EntraCAReport.ps1 -EnabledOnly -ExportPath "C:\temp\ca_policies.csv"
```
Exports active policies only to a specific CSV.

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `PolicyName` | String | No | - | Filter by display name (partial match). |
| `EnabledOnly` | Switch | No | False | Only show enabled policies. |
| `IncludeDisabled` | Switch | No | False | Include disabled policies. |
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
* `Policy.Read.All, Directory.Read.All, Application.Read.All, Group.Read.All`

### Logging
* `C:\ProgramData\Get-EntraCAReport\Logs\`

---

# 🛡 Operational Notes
* Read-only; never modifies Conditional Access policies.
* Unresolvable IDs render as raw IDs without failing the report.
* Report-only policies are included by default; use -EnabledOnly to narrow.

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