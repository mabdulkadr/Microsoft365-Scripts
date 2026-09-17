<div align="center">

# 💻 Add-DevicesToAzureADGroup

**Bulk-add devices to an Entra ID group by device name.**

Resolves names to Device Object IDs first — the step Azure AD's bulk import can't do.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**Add-DevicesToAzureADGroup** is a PowerShell script that bulk-adds devices to an Entra ID group from a CSV of device names. Azure AD bulk import needs Device Object IDs, so the script looks each name up dynamically, adds every match (including duplicate names), and skips devices that are missing or already members.

---

# ✨ Features

* Bulk add from a CSV of device names
* Automatic Device ID lookup per name
* Handles duplicate device names (adds all matches)
* Skips not-found / already-member devices with per-device output

---

# 📂 Project Structure

```text
Add-DevicesToAzureADGroup
│
├── Add-DevicesToAzureADGroup.ps1
├── SampleDevicesFile.csv
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Add-DevicesToAzureADGroup.ps1 -GroupName "Device Test Group" -InputFile "C:\Scripts\DevicesToAdd.csv"
```

### CSV Format
```csv
DeviceName
Device-01
Device-02
Device-03
```

Device names must match Entra ID exactly.

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `GroupName` | String | Yes | — | Target Entra ID group name. |
| `InputFile` | String | Yes | — | CSV path with a `DeviceName` column. |

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
* `AzureAD` module (`Install-Module AzureAD -Scope CurrentUser`)

### Permissions
* Group membership management; app permissions `Group.ReadWrite.All`, `Device.Read.All`, `Directory.Read.All` for service runs.

---

# 🛡 Operational Notes
* Connects to Azure AD automatically before operating.
* Members already in the group are skipped, not errored.
* Sample input: `SampleDevicesFile.csv` in this folder.
* Reference: [Andrew IT Dev Lab — Add Devices to Group in Azure AD](https://github.com/andrewitdevlab/blog-content/tree/main/Azure%20AD/Scripts/Add%20Devices%20to%20Group).

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
