<div align="center">

# 📦 Microsoft365-Scripts

**Enterprise PowerShell toolkit for Microsoft 365 / Entra ID administration.**

28 tools across 6 domains — every tool lives in its own folder with its script and README.

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Tools](https://img.shields.io/badge/Tools-28%20Tools-10B981?style=for-the-badge)](#-tool-catalog)

[Catalog](#-tool-catalog) • [Layout](#-layout) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Microsoft365-Scripts** is a collection of production-grade PowerShell CLI tools for Microsoft 365 and Entra ID operations: reporting, auditing, group management, guest lifecycle, licensing, and bulk administration. Each tool is self-contained (script + README) and follows the same conventions: canonical headers, structured logging, beside-the-script reports, and PS 5.1 compatibility.

---

# 🗂️ Tool Catalog

## 📊 Reporting & Audit (Entra ID)

| Tool | What it does |
|------|--------------|
| [Get-EntraAdminRoleReport](Get-EntraAdminRoleReport/README.md) | Admin role assignments and privileged-access posture |
| [Get-EntraAppRegistrationAudit](Get-EntraAppRegistrationAudit/README.md) | App registration and credential hygiene audit |
| [Get-EntraCAReport](Get-EntraCAReport/README.md) | Conditional Access policy coverage report |
| [Get-EntraDirectoryAudit](Get-EntraDirectoryAudit/README.md) | Directory audit log review with sensitive-operation flags |
| [Get-EntraGroupAudit](Get-EntraGroupAudit/README.md) | Group inventory with optional member export |
| [Get-EntraGuestUserAudit](Get-EntraGuestUserAudit/README.md) | Guest user posture audit |
| [Get-EntraLicenseReport](Get-EntraLicenseReport/README.md) | Per-user license assignment matrix + waste spotlight |
| [Get-EntraRiskyUsers](Get-EntraRiskyUsers/README.md) | Risky users + 14-day risk detections |
| [Get-EntraSignInReport](Get-EntraSignInReport/README.md) | Sign-in log analysis (failures, MFA, legacy auth, risk) |

## 📊 Reporting (Exchange / Mailbox)

| Tool | What it does |
|------|--------------|
| [Get-GroupMailboxActivityReport](Get-GroupMailboxActivityReport/README.md) | Per-group mailbox activity (7/30/90d) with HTML dashboard + email |
| [Get-M365UserActivityReport](Get-M365UserActivityReport/README.md) | Exchange Online last-logon per mailbox |
| [Get-M365UserLastActivityReport](Get-M365UserLastActivityReport/README.md) | Real last-logon report with inactivity filters |
| [Get-PasswordExpiryReport](Get-PasswordExpiryReport/README.md) | Password change/expiry reports, six angles |
| [GetM365InactiveUserReport](GetM365InactiveUserReport/README.md) | Inactive users from Graph sign-in activity |
| [HybridUserAudit](HybridUserAudit/README.md) | Merged on-prem AD + Entra ID user report |
| [Invoke-M365LicenseReport](Invoke-M365LicenseReport/README.md) | License allocation and usage (SKU-level) |

## 👥 Group Management

| Tool | What it does |
|------|--------------|
| [Add-DevicesToAzureADGroup](Add-DevicesToAzureADGroup/README.md) | Bulk-add devices to a group by device name |
| [Add-UsersToAzureADGroup](Add-UsersToAzureADGroup/README.md) | Bulk-add users to a security group from CSV |
| [AzureAD-DeviceGroupManagement](AzureAD-DeviceGroupManagement/README.md) | Static Intune device groups (interactive flow) |
| [AzureAD-DeviceGroupManagement-AppAuth](AzureAD-DeviceGroupManagement-AppAuth/README.md) | Static Intune device groups (app-only, unattended) |
| [New-DynamicGroupFromCsv](New-DynamicGroupFromCsv/README.md) | Bulk dynamic groups from CSV |
| [New-DynamicGroupByKeyword](New-DynamicGroupByKeyword/README.md) | One dynamic group by device-name keyword |
| [New-DynamicGroupFromAdOu](New-DynamicGroupFromAdOu/README.md) | Mirror on-prem OUs as dynamic device groups |
| [Add-ExchangeOnlineUsersToDistributionGroup](Add-ExchangeOnlineUsersToDistributionGroup/README.md) | Bulk-add users to an EXO distribution group |

## 🧑‍💼 Guest Lifecycle

| Tool | What it does |
|------|--------------|
| [Get-GuestUserReport](Get-GuestUserReport/README.md) | Guest inventory export (audit first) |
| [Remove-GuestUsersFromCsv](Remove-GuestUsersFromCsv/README.md) | Bulk guest removal with safety checks (remove second) |

## 🛡️ Bulk Administration

| Tool | What it does |
|------|--------------|
| [BulkUserResetTool](BulkUserResetTool/README.md) | Bulk password reset, session revoke, sign-in block, device disable |

## 📝 Documentation

| Tool | What it does |
|------|--------------|
| [Invoke-M365Documentation](Invoke-M365Documentation/README.md) | Word documentation for Intune + Entra ID via menu picker |

---

# 📂 Layout

```text
Microsoft365-Scripts
│
├── README.md                          ← this catalog
├── Get-EntraAdminRoleReport/
│   ├── Get-EntraAdminRoleReport.ps1
│   └── README.md
├── Remove-GuestUsersFromCsv/
│   ├── Remove-GuestUsersFromCSV.ps1
│   ├── exportUsers_2025-7-3.csv
│   └── README.md
└── … (one folder per tool, same shape)
```

Every tool folder contains its script plus its own README with parameters, exit codes, requirements, and operational notes.

---

# ⚙️ Requirements

* Windows 10 / Windows 11 with PowerShell **5.1 or later**
* Per-tool modules and Graph permissions are listed in each tool's README
* Logs: `C:\ProgramData\<ToolName>\Logs\` (general CLI tools)

---

# 🛡 Operational Notes
* Test every tool in staging before production — see each README's Disclaimer.
* Guest lifecycle order: `Get-GuestUserReport` (audit) → `Remove-GuestUsersFromCsv` (remove).
* Dynamic-group variants: CSV bulk, single keyword, or on-prem OU mirror — pick one per need.

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
