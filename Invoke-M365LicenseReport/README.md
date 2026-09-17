<div align="center">

# 💻 Invoke-M365LicenseReport

**License allocation and usage report for Microsoft 365 (CSV + HTML).**

Collects subscribed SKUs and per-user assignments from Graph, computes used vs unused counts per plan, and renders console + HTML output.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [License](#-license)

</div>

---

# 📖 Overview

**Invoke-M365LicenseReport** is a PowerShell script that generates a Microsoft 365 license allocation and usage report. It connects to Microsoft Graph, pulls subscribed SKUs and user license assignments, and highlights waste (e.g. licenses on disabled accounts).

---

# ✨ Features

* Subscribed-SKU inventory with used/unused counts per plan
* Per-user license assignment detail
* Waste spotlight: disabled accounts still holding licenses, guests with licenses
* Console summary plus report export beside the script: SKU/inactive/privileged CSVs + Carbon Dark HTML dashboard

---

# 📂 Project Structure

```text
Invoke-M365LicenseReport
│
├── Invoke-M365LicenseReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Invoke-M365LicenseReport.ps1 -outpath "C:\Reports"
```

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `outpath` | String | Yes | — | Output folder for the generated report. |

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
* Microsoft Graph modules (installed on confirmation if missing)

### Permissions
* `User.Read.All`, `AuditLog.Read.All`, `Organization.Read.All`, `RoleManagement.Read.Directory` (Microsoft Graph).

---

# 🛡 Operational Notes
* `-outpath` is mandatory — the folder is created if missing.
* Reclaim licenses surfaced in the waste section before buying more seats.
* For a per-user license matrix instead, see [`Get-EntraLicenseReport`](../Get-EntraLicenseReport/README.md).

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
