<div align="center">

# 💻 New-DynamicGroupFromAdOu

**Mirror on-prem AD OUs as dynamic Entra ID device groups.**

Creates one `Devices-<OU>` dynamic group per on-premises OU, targeting devices by `onPremisesDistinguishedName`.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**New-DynamicGroupFromAdOu** is a PowerShell script for hybrid environments. It reads OU names from on-premises Active Directory and creates matching dynamic Entra ID groups (prefix `Devices-`), each with a membership rule on `onPremisesDistinguishedName`. Existing group names are checked first so reruns create no duplicates.

---

# ✨ Features

* Discovers all OU names from on-prem AD automatically
* Creates one dynamic group per OU with a `Devices-` prefix
* Pre-creation duplicate check — safe to rerun
* Auto-installs required PowerShell modules; secure app-based Graph auth

---

# 📂 Project Structure

```text
New-DynamicGroupFromAdOu
│
├── New-DynamicGroupFromAdOu.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\New-DynamicGroupFromAdOu.ps1
```

### How It Works
1. Edit the configuration block (`$tenantID`, `$appID`, `$appSecret`, `$groupPrefix`).
2. Run from a machine joined to (or able to reach) the on-prem AD; the script pulls OU names and mirrors them to Entra ID.

---

# ⚙️ Parameters

This script takes no command-line parameters. All settings live in the configuration block at the top of the script:

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `$groupPrefix` | String | `Devices - ` | Prefix for created group names. |
| `$tenantID` / `$appID` / `$appSecret` | String | `""` | App-based auth credentials. |
| `$logFilePath` | String | `C:\CreateDynamicGroups.log` | Local log file path. |

### Exit Codes
| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

---

# ⚙️ Requirements

### Operating System
* Windows 10 / Windows 11 / Windows Server 2019+ with line-of-sight to on-prem AD

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* On-prem AD read (OU enumeration); Graph `Group.ReadWrite.All`, `Directory.Read.All`.

### Logging
* Local log file (see `$logFilePath` above).

---

# 🛡 Operational Notes
* Requires connectivity to **both** on-prem AD and Microsoft Graph — it is a hybrid-only tool.
* Never commit real tenant IDs or secrets; replace the checked-in placeholders per environment.
* CSV-driven variant: [`New-DynamicGroupFromCsv`](../New-DynamicGroupFromCsv/README.md). Single-keyword variant: [`New-DynamicGroupByKeyword`](../New-DynamicGroupByKeyword/README.md).

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
