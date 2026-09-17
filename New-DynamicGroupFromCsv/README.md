<div align="center">

# 💻 New-DynamicGroupFromCsv

**Create dynamic Entra ID groups in bulk from a CSV file.**

Bulk-provisions dynamic Azure AD groups from CSV input with per-row membership rules, for Intune-managed device fleets.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**New-DynamicGroupFromCsv** is a PowerShell script that creates multiple dynamic Entra ID groups from a CSV file. Each row defines a group name, description, and membership rule criteria, so device fleets that follow a naming convention can be grouped automatically at scale.

---

# ✨ Features

* Reads group definitions from CSV (`GroupName`, `Description`, `Rule` columns)
* Generates a dynamic membership rule per CSV row (e.g. on device display names)
* App-based (client credentials) or user-based (delegated) Graph authentication
* Skips/commits per row with structured error handling and a creation summary

---

# 📂 Project Structure

```text
New-DynamicGroupFromCsv
│
├── New-DynamicGroupFromCsv.ps1
├── Groups.csv
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\New-DynamicGroupFromCsv.ps1
```

### How It Works
1. Edit the configuration block inside the script (`$tenantID`, `$appID`, `$appSecretPlain`, `$csvFilePath`).
2. Prepare `Groups.csv` with `GroupName`, `Description`, and `Rule` columns.
3. Run the script; it connects to Microsoft Graph and creates one dynamic group per row.

```csv
GroupName,Description,Rule
IT-Ops Devices,Dynamic group for IT-Ops devices,it-op
HR Devices,Dynamic group for HR devices,hr-
```

---

# ⚙️ Parameters

This script takes no command-line parameters. All settings live in the configuration block at the top of the script:

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `$tenantID` | String | `""` | Entra ID tenant ID for app-based auth. |
| `$appID` | String | `""` | App (client) ID of the service principal. |
| `$appSecretPlain` | String | `""` | Client secret (prefer a certificate or Key Vault in production). |
| `$csvFilePath` | String | `C:\Groups.csv` | Path to the input CSV file. |

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

### Permissions
* `Group.ReadWrite.All`, `Directory.Read.All` (Microsoft Graph).

### Logging
* `C:\ProgramData\Microsoft365Scripts\Logs\`

---

# 🛡 Operational Notes
* Never commit real tenant IDs or secrets; the checked-in placeholders must be replaced per environment.
* Prefer certificate-based auth or Azure Key Vault over plaintext secrets for scheduled runs.
* Validate membership rules against a single test group before bulk-creating dozens of groups.
* Related single-group variant: [`New-DynamicGroupByKeyword`](../New-DynamicGroupByKeyword/README.md). OU-driven variant: [`New-DynamicGroupFromAdOu`](../New-DynamicGroupFromAdOu/README.md).

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
