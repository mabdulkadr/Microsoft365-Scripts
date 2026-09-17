<div align="center">

# 💻 Add-UsersToAzureADGroup

**Bulk-add users to an Entra ID security group from CSV.**

Existence-checked, duplicate-safe, throttled against Graph API limits.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**Add-UsersToAzureADGroup** is a PowerShell script that adds users to an Entra ID security group from a CSV of UPNs. Each user is verified to exist, skipped if already a member, and added otherwise — with rate-limiting to avoid API throttling.

---

# ✨ Features

* CSV-driven bulk add by UPN
* Existence check before every add
* Duplicate-safe (existing members skipped)
* Rate-limited calls plus detailed per-user error messages

---

# 📂 Project Structure

```text
Add-UsersToAzureADGroup
│
├── Add-UsersToAzureADGroup.ps1
├── SampleUsersFile.csv
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Add-UsersToAzureADGroup.ps1
```

### CSV Format
```
UPN
user1@domain.com
user2@domain.com
user3@domain.com
```

### Customization
```powershell
$GroupID = "your-group-id"
$CSVFilePath = "C:\path\to\your\file.csv"
```

---

# ⚙️ Parameters

This script takes no command-line parameters. Target group and CSV path are set inside the script (see Customization above).

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
* `AzureAD` module (`Install-Module -Name AzureAD -Force`)

### Permissions
* Admin privileges to manage Entra ID groups.

---

# 🛡 Operational Notes
* Per-user outcomes: `User does not exist in Azure AD` / `already a member` / `Error adding UPN: <message>`.
* Sample input: `SampleUsersFile.csv` in this folder.
* Test in staging before production rollouts.

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
