<div align="center">

# 💻 HybridUserAudit

**One merged user report across on-prem AD + Entra ID (100K+ ready).**

Username-joined attributes, dual last-logons, CSV + log output for audit and compliance.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**HybridUserAudit** is a PowerShell script that merges on-premises Active Directory and Entra ID user data on username, tracks login activity from both sides, and writes a UTF-8 CSV for auditing, compliance, or inventory. Read-only — it never modifies users.

---

# ✨ Features

* Username-joined AD + Entra attributes (display name, email, creation, logons, …)
* Presence flags (`InAD` / `InEntraID`) plus per-side enabled/created/logon columns
* Live console progress, detailed log file, optimized for 100K+ directories

---

# 📂 Project Structure

```text
HybridUserAudit
│
├── HybridUserAudit.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\HybridUserAudit.ps1
```

---

# ⚙️ Parameters

This script takes no command-line parameters; run it from an elevated 5.1 session (ISE recommended).

### Outputs
* **CSV** `Reports\FullUserReport_yyyy-MM-dd_HH-mm.csv` (beside the script) — Username, DisplayName, Department, Title, Email, InAD/InEntraID, AD_/Entra_ Enabled/Created, AD_LastLogon, Entra_LastInteractiveSignIn, AD_WhenChanged, AD_PwdLastSet, AD_Description, AD_DistinguishedName
* **HTML** `Reports\FullUserReport_yyyy-MM-dd_HH-mm.html` — Carbon Dark dashboard (KPIs + detail table, same rows)
* **Log** `Reports\HybridUserAuditLog_yyyy-MM-dd_HH-mm.txt` (beside the script)

### Exit Codes
| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

---

# ⚙️ Requirements

### Operating System
* Windows 10 / Windows 11 (domain-joined for the AD side)

### PowerShell
* PowerShell **5.1 or later** (run elevated)
* `ActiveDirectory`, `Microsoft.Graph.Users` modules

### Permissions
* AD read + Graph user/sign-in read.

---

# 🛡 Operational Notes
* Read-only by design — safe to run in production, but allow several minutes on large directories.
* UTF-8 export preserves non-Latin attributes.

---

## 👤 Author

**Mohammad Abdelkader Omar** ([momar.tech](https://momar.tech))
📅 Last Updated: 2025-06-23

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
