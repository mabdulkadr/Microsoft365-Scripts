<div align="center">

# 💻 Get-PasswordExpiryReport

**Microsoft 365 password expiry reports via Microsoft Graph.**

One script, six report angles: full inventory, never-expires, expired, soon-to-expire, recent changers — sliced by license and sign-in status.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**Get-PasswordExpiryReport** is a PowerShell script that exports Office 365 users' last password change and expiry dates via Microsoft Graph. It auto-installs the Graph SDK on confirmation, supports certificate-based app auth and MFA accounts, and writes every report to CSV.

---

# ✨ Features

* Six report angles: all users, never-expires, expired, soon-to-expire, recent changers
* Scope slices: all vs licensed users, all vs sign-in-enabled users
* Certificate-based authentication + MFA-friendly interactive sign-in
* CSV + Carbon Dark HTML dashboard exports for every angle

---

# 📂 Project Structure

```text
Get-PasswordExpiryReport
│
├── Get-PasswordExpiryReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-PasswordExpiryReport.ps1
```

### With Parameters
```powershell
.\Get-PasswordExpiryReport.ps1 -PwdNeverExpires
.\Get-PasswordExpiryReport.ps1 -SoonToExpire 14 -LicensedUserOnly
```

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `PwdNeverExpires` | Switch | No | off | Only users with passwords set to never expire. |
| `PwdExpired` | Switch | No | off | Only users with expired passwords. |
| `LicensedUserOnly` | Switch | No | off | Only licensed users. |
| `SoonToExpire` | Int | No | (off) | Users expiring within N days. |
| `RecentPwdChanges` | Int | No | (off) | Users changed within the last N days. |
| `EnabledUsersOnly` | Switch | No | off | Only sign-in-enabled users. |
| `TenantId` | String | No | (interactive) | Tenant ID for certificate app auth. |
| `ClientId` | String | No | (interactive) | App (client) ID for certificate app auth. |
| `CertificateThumbprint` | String | No | (interactive) | Certificate thumbprint for app auth. |

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
* MS Graph PowerShell SDK (installed on confirmation if missing)

### Permissions
* `User.Read.All`, `Directory.Read.All` (Microsoft Graph).

---

# 🛡 Operational Notes
* Combine switches to narrow scope, e.g. `-PwdExpired -LicensedUserOnly -EnabledUsersOnly`.
* For unattended runs, supply `TenantId` + `ClientId` + `CertificateThumbprint`.
* Treat expiry exports as sensitive — they enumerate account hygiene tenant-wide.

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
