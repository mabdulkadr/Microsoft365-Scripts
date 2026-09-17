<div align="center">

# 💻 New-DynamicGroupByKeyword

**Create one dynamic Entra ID group for devices matching a name keyword.**

Single-group variant for device fleets that follow a naming convention (e.g. `it-op` in the display name).

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**New-DynamicGroupByKeyword** is a PowerShell script that creates a single dynamic Entra ID group whose membership rule targets devices whose display name contains a keyword (default `it-op`). Ideal for Intune/Entra ID environments where devices follow a naming convention.

---

# ✨ Features

* Creates one dynamic group with a keyword-based membership rule
* App-based (service principal) Graph authentication with secure credential handling
* Structured error management with a creation summary
* No duplicate handling needed — single-shot, single-group flow

---

# 📂 Project Structure

```text
New-DynamicGroupByKeyword
│
├── New-DynamicGroupByKeyword.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\New-DynamicGroupByKeyword.ps1
```

### How It Works
1. Edit the configuration block inside the script (`$tenantID`, `$appID`, `$appSecretPlain`, `$groupName`, `$membershipRule`).
2. Run the script; it connects to Microsoft Graph and creates the group.

---

# ⚙️ Parameters

This script takes no command-line parameters. All settings live in the configuration block at the top of the script:

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `$groupName` | String | `IT-Operations Devices` | Display name of the group to create. |
| `$membershipRule` | String | `(device.displayName -contains 'it-op')` | Dynamic membership rule. |
| `$tenantID` / `$appID` / `$appSecretPlain` | String | `""` | App-based auth credentials. |

### Exit Codes
| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |
| 2    | Script error |

---

# ⚙️ Requirements

### Operating System
* Windows 10 / Windows 11

### PowerShell
* PowerShell **5.1 or later**
* Modules `Microsoft.Graph.Groups`, `Microsoft.Graph.Authentication` (auto-installed if missing)

### Permissions
* `Group.ReadWrite.All`, `Directory.Read.All` (Microsoft Graph).

---

# 🛡 Operational Notes
* Never commit real tenant IDs or secrets; replace the checked-in placeholders per environment.
* Test the membership rule preview in Entra ID before relying on it for policy targeting.
* Bulk variant (CSV input): [`New-DynamicGroupFromCsv`](../New-DynamicGroupFromCsv/README.md). OU-driven variant: [`New-DynamicGroupFromAdOu`](../New-DynamicGroupFromAdOu/README.md).

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
