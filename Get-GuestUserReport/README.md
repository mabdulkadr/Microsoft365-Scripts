<div align="center">

# 💻 Get-GuestUserReport

**Export all Entra ID guest users and their group memberships to CSV + Carbon HTML.**

Audit companion to [`Remove-GuestUsersFromCsv`](../Remove-GuestUsersFromCsv/README.md) — report first, remove second.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**Get-GuestUserReport** is a PowerShell script that exports every guest (`UserType = Guest`) in the Entra ID tenant to a timestamped CSV: display name, UPN, email, company, invitation/redemption status, creation date, and group memberships. Use it as the audit record before any bulk removal.

---

# ✨ Features

* Full guest inventory via Microsoft Graph (`UserType = Guest`)
* Age filters for stale-account reviews (`-StaleGuests`, `-RecentlyCreatedGuests`)
* Per-guest group membership collection
* Timestamped CSV + Carbon Dark HTML dashboard (shared timestamp) with an offer to open the CSV on completion

---

# 📂 Project Structure

```text
Get-GuestUserReport
│
├── Get-GuestUserReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-GuestUserReport.ps1
```

### With Parameters
```powershell
.\Get-GuestUserReport.ps1 -StaleGuests 180
```

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `StaleGuests` | Int | No | (all) | Only guests older than N days. |
| `RecentlyCreatedGuests` | Int | No | (all) | Only guests newer than N days. |

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
* `Microsoft.Graph` module (`Install-Module Microsoft.Graph -Scope CurrentUser`)

### Permissions
* `User.Read.All`, `Directory.Read.All` (Microsoft Graph).

---

# 🛡 Operational Notes
* Report columns: Display Name, UPN, Email, Company, Invitation/Redemption status, Creation Date, Group Memberships.
* Review the report to confirm identity and account status before deleting anything.
* Pair with [`Remove-GuestUsersFromCsv`](../Remove-GuestUsersFromCsv/README.md) for the full audit → remove lifecycle.

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
