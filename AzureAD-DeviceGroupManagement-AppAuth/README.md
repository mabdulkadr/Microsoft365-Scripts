<div align="center">

# 💻 AzureAD-DeviceGroupManagement-AppAuth

**Group Intune devices into static Entra ID groups using app-only authentication.**

Fully unattended variant of the device group manager — authenticates with Tenant ID / App ID / App Secret, no interactive sign-in.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**AzureAD-DeviceGroupManagement-AppAuth** is a PowerShell script that manages static Entra ID groups for Intune devices with app-only authentication. It finds existing static groups by prefix, creates new ones as needed, and distributes devices so no group exceeds the member limit. Companion to [`AzureAD-DeviceGroupManagement`](../AzureAD-DeviceGroupManagement/README.md), which uses the interactive/manual connection flow.

---

# ✨ Features

* App-only Graph connection (Tenant ID / App ID / App Secret) — scheduler-friendly
* Prefix-based group discovery with automatic overflow group creation
* Configurable batch size and zero-padded group numbering
* Optional file logging with `[timestamp] [LEVEL]` lines

---

# 📂 Project Structure

```text
AzureAD-DeviceGroupManagement-AppAuth
│
├── AzureAD-DeviceGroupManagement-AppAuth.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\AzureAD-DeviceGroupManagement-AppAuth.ps1
```

### With Parameters
```powershell
.\AzureAD-DeviceGroupManagement-AppAuth.ps1 -BatchSize 300 -GroupNamePrefix "CorporateDevices-" -NamePadding 3 -EnableLogging -LogFilePath "D:\Logs\DeviceGroups.log"
```

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `BatchSize` | Int | No | `500` | Max devices per group before a new overflow group is created. |
| `GroupNamePrefix` | String | No | `Devices-group` | Prefix for discovered/created static groups. |
| `NamePadding` | Int | No | `2` | Zero-padded digits in group numbering (`01`, `02`, …). |
| `EnableLogging` | Switch | No | off | Writes operations to the log file. |
| `LogFilePath` | String | No | `GroupCreationLog.txt` (beside the script) | Log file path (used with `-EnableLogging`). |

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
* `AzureAD` module; Microsoft Graph modules (auto-installed if missing)

### Permissions
* Entra ID admin able to create groups and manage memberships; Intune device read.

---

# 🛡 Operational Notes
* Manages **static** memberships only — no dynamic rules.
* Store the App Secret securely (Key Vault / Secret Store); never commit real secrets.
* For the interactive variant see [`AzureAD-DeviceGroupManagement`](../AzureAD-DeviceGroupManagement/README.md).

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
