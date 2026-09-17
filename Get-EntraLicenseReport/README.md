<div align="center">

# 📊 Get-EntraLicenseReport

**Reports Entra ID license utilization and identifies waste.**

Lists all license SKUs with assigned versus available counts, identifies users with no licenses and disabled accounts still consuming licenses, and calculates utilization rates for optimization.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.1-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Structure](#-project-structure) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get-EntraLicenseReport** is a PowerShell reporting script that lists all license SKUs with assigned versus available counts, identifies users with no licenses and disabled accounts still consuming licenses, and calculates utilization rates for optimization.

It pulls subscribed SKUs and user license assignments to show total versus consumed versus available units, then drills into user-level waste (disabled accounts with licenses, unlicensed users, guests with licenses) for optimization and cost review.

---

# ✨ Features

* Lists SKU totals, consumed, available, and utilization percentage
* Identifies disabled accounts still consuming licenses (waste)
* Surfaces unlicensed users and guests with licenses
* Calculates per-SKU and per-user assignment counts
* Exports user-level CSV + Carbon Dark HTML dashboard (shared timestamp) beside the script for finance review

---

# 📂 Project Structure

```text
Get-EntraLicenseReport
│
├── Get-EntraLicenseReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-EntraLicenseReport.ps1
```

### Example 1
```powershell
.\Get-EntraLicenseReport.ps1
```
Generates the license utilization report for the tenant.

### Example 2
```powershell
.\Get-EntraLicenseReport.ps1 -ExportPath "C:\Reports\Licenses.csv"
```
Exports user license data to a specific CSV.

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `ExportPath` | String | No | Beside script | Optional CSV export path for user license data. |

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
* `Organization.Read.All, User.Read.All, Directory.Read.All`

### Logging
* `C:\ProgramData\Get-EntraLicenseReport\Logs\`

---

# 🛡 Operational Notes
* Read-only; never modifies licenses or users.
* Waste is defined as disabled accounts with assignedLicenses — review before removal.
* SKU display names use skuPartNumber; map to friendly names separately if needed.

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