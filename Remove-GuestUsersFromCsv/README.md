<div align="center">

# 💻 Remove-GuestUsersFromCsv

**Bulk-remove Entra ID guest users listed in a CSV file — with safety checks.**

Second half of the guest lifecycle toolkit: [`Get-GuestUserReport`](../Get-GuestUserReport/README.md) audits, this script removes.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**Remove-GuestUsersFromCsv** is a PowerShell script that bulk-deletes guest accounts from Entra ID using a CSV list. Each row is looked up by `UserPrincipalName`, verified to be a guest (`UserType = Guest`), and only then deleted — member accounts are never touched. A console summary reports deleted / skipped / not-found / failed counts.

---

# ✨ Features

* Interactive file dialog for CSV selection — no hardcoded paths
* Guest-type verification per row (members are skipped, never deleted)
* Per-action logging with an end-of-run console summary
* Sample CSVs included (`exportUsers_2025-7-3.csv`)

---

# 📂 Project Structure

```text
Remove-GuestUsersFromCsv
│
├── Remove-GuestUsersFromCSV.ps1
├── exportUsers_2025-7-3.csv
├── exportUsers_2025-7-3-IT-OP-030.csv
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Remove-GuestUsersFromCSV.ps1
```

### How It Works
1. Run the script and select your CSV file when prompted.
2. For each row: look up the account by `UserPrincipalName`, confirm `UserType = Guest`, delete it.
3. Review the console summary (deleted / skipped / not found / failed).

### CSV Format
```csv
DisplayName,UserPrincipalName
m.abdelkader,m.abdelkader_upm.xyz.com#EXT#@abc.onmicrosoft.com
372113519,372113519_cloud.sa#EXT#@abc.onmicrosoft.com
```

---

# ⚙️ Parameters

This script takes no command-line parameters; the input CSV is chosen via file dialog at runtime.

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
* `Microsoft.Graph` module (`Install-Module Microsoft.Graph -Scope CurrentUser`)

### Permissions
* Entra ID rights to delete guest users.

---

# 🛡 Operational Notes
* Destructive by design: test on a small CSV (or staging tenant) before bulk runs.
* Keep a backup of the input CSV and the pre-removal report as your audit record.
* CSV must contain a `UserPrincipalName` column; extra columns are ignored.

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
