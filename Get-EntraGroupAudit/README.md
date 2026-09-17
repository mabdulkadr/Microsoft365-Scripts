<div align="center">

# 📊 Get-EntraGroupAudit

**Audits Entra ID group membership, ownership, nesting, and configuration.**

Provides detailed group analysis including members, direct versus transitive membership, owners, dynamic rules, nested hierarchy, license assignments, and group type classification. Supports single-group deep dive or bulk health audit for empty, large, or ownerless groups.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.1-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Structure](#-project-structure) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get-EntraGroupAudit** is a PowerShell reporting script that provides detailed group analysis including members, direct versus transitive membership, owners, dynamic rules, nested hierarchy, license assignments, and group type classification. Supports single-group deep dive or bulk health audit for empty, large, or ownerless groups.

It supports single-group deep dive (with member listing and owner resolution) and bulk audit mode that flags empty groups, ownerless groups, very large groups, and stale dynamic groups across the tenant.

---

# ✨ Features

* Resolves members, owners, and transitive versus direct membership
* Evaluates dynamic membership rules and nested group hierarchy
* Classifies group type (security, M365, dynamic, role-assignable)
* Bulk audit flags empty, ownerless, and oversized groups
* Exports CSV + Carbon Dark HTML dashboard (shared timestamp) beside the script for governance reviews

---

# 📂 Project Structure

```text
Get-EntraGroupAudit
│
├── Get-EntraGroupAudit.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Get-EntraGroupAudit.ps1
```

### Example 1
```powershell
.\Get-EntraGroupAudit.ps1 -GroupName "SG-Intune-Windows-Devices"
```
Runs a deep dive on a single group.

### Example 2
```powershell
.\Get-EntraGroupAudit.ps1 -GroupName "SG-Intune-Pilot" -IncludeMembers -ExportPath "C:\temp\members.csv"
```
Exports members of a group to CSV.

### Example 3
```powershell
.\Get-EntraGroupAudit.ps1 -BulkAudit -ExportPath "C:\temp\group_health.csv"
```
Runs a health audit across all groups.

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `GroupName` | String | No* |  - | Audit a specific group by display name (ByName set). |
| `GroupId` | String | No* |  - | Audit a specific group by object ID (ById set). |
| `BulkAudit` | Switch | No* | False | Run health audit across all groups. |
| `IncludeMembers` | Switch | No | False | List individual members in single-group mode. |
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
* `Directory.Read.All, Group.Read.All, GroupMember.Read.All, User.Read.All`

### Logging
* `C:\ProgramData\Get-EntraGroupAudit\Logs\`

---

# 🛡 Operational Notes
* Read-only; never modifies groups or memberships.
* Bulk audit skips member expansion for performance; use IncludeMembers only for single-group mode.
* Dynamic rule evaluation is read from the group object; sync latency may apply.

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