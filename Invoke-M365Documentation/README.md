<div align="center">

# 💻 Invoke-M365Documentation

**Generate Word documentation for Microsoft 365 components (Intune + Entra ID).**

Menu-driven section picker on top of the M365Documentation module, with timestamped `.docx` output.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**Invoke-M365Documentation** is a PowerShell script that generates detailed Word documentation for Microsoft 365 components, focusing on **Intune** and **Entra ID**. A menu lets you pick the component and include/exclude sections dynamically. Data collection and rendering are powered by the [M365Documentation module](https://github.com/ThomasKur/M365Documentation).

---

# ✨ Features

* Intune coverage: configuration/compliance policies, apps, autopilot, baselines, roles, and more
* Entra ID coverage: domains, CA policies, auth policies, rollout policies, org settings, SKUs, admin units
* Dynamic include/exclude section picker with input validation
* Timestamped Word output per component

---

# 📂 Project Structure

```text
Invoke-M365Documentation
│
├── Invoke-M365Documentation.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Invoke-M365Documentation.ps1
```

### How It Works
1. Fill in `$TenantId`, `$ClientId`, `$ClientSecret` inside the script.
2. Run it, pick **Intune** or **AzureAD**, optionally filter sections.
3. Collect `Reports\<timestamp>-<component>-Documentation.docx` (beside the script).

---

# ⚙️ Parameters

This script takes no command-line parameters. Credentials and `$OutputDirectory` (`Reports\` beside the script) are set inside the script.

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

### Modules
* `MSAL.PS`, `PSWriteOffice`, `M365Documentation` (auto-installed if missing)

### Permissions
* App registration with Graph application permissions (`Directory.Read.All`, `DeviceManagement*.Read.All`, `Domain.Read.All`, `Policy.Read.All`, `Organization.Read.All`, `User.Read`, …) + admin consent.

### Logging
* `C:\ProgramData\Microsoft365Scripts\Logs\`

---

# 🛡 Operational Notes
* If API permission errors occur: verify the app registration, re-grant admin consent, or `Connect-MgGraph` manually first.
* If module installs fail: `Install-Module -Name MSAL.PS, PSWriteOffice, M365Documentation -Force`.
* Never commit real Tenant IDs or secrets.
* Reference: [M365Documentation module](https://github.com/ThomasKur/M365Documentation).

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
