<div align="center">

# 📊 Get-EntraRiskyUsers

**Reports risky users and risk detections from Entra ID Protection.**

Pulls risky users and risk detection events showing risk level, risk state, risk detail, last detection time, and remediation status. Requires Entra ID P2 for full data.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.1-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Structure](#-project-structure) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get-EntraRiskyUsers** is a PowerShell reporting script that pulls risky users and risk detection events showing risk level, risk state, risk detail, last detection time, and remediation status. Requires Entra ID P2 for full data.

It queries identityProtection/riskyUsers and riskDetections to surface high, medium, and low risk users with state (atRisk, confirmedCompromised) and recent detections. Use it for SOC triage and risk remediation tracking; requires Entra ID P2.

---

# ✨ Features

* Pulls risky users with risk level, state, and detail
* Shows risk detection events with time and remediation status
* Filters dismissed versus active risk via -IncludeDismissed
* Handles P2 licensing gracefully when data is unavailable
* Exports CSV + Carbon Dark HTML dashboard (shared timestamp) beside the script for SOC handoff

---

# 📂 Project Structure

```text
Get-EntraRiskyUsers
│
├── Get-EntraRiskyUsers.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-EntraRiskyUsers.ps1
```

### Example 1
```powershell
.\Get-EntraRiskyUsers.ps1
```
Reports active risky users and detections.

### Example 2
```powershell
.\Get-EntraRiskyUsers.ps1 -IncludeDismissed
```
Includes dismissed risky users in the report.

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `IncludeDismissed` | Switch | No | False | Include users whose risk has been dismissed. |
| `ExportPath` | String | No | Beside script | Optional CSV export path. |

### Exit Codes
| Code | Status |
| ---- | ------ |
| 0 | Success |
| 1 | Failure |

---

# ⚙️ Requirements

### Operating System
* Windows 10 / Windows 11

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* `IdentityRiskyUser.Read.All, IdentityRiskEvent.Read.All`

### Logging
* `C:\ProgramData\Get-EntraRiskyUsers\Logs\`

---

# 🛡 Operational Notes
* Read-only; never modifies risk state.
* Requires Entra ID P2 for full risky user and detection data.
* Dismissed risk is excluded by default; use -IncludeDismissed for complete history.

---

## 👤 Author
**Mohammad Abdelkader Omar**  
GitHub: [@mabdulkadr](https://github.com/mabdulkadr)  
Website: [momar.tech](https://momar.tech)
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