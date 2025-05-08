# HybridUserAudit.ps1

![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Status](https://img.shields.io/badge/stability-stable-brightgreen.svg)

## Overview

`HybridUserAudit.ps1` is a scalable and memory-efficient PowerShell script that generates a unified user report by comparing on-premises **Active Directory (AD)** users with cloud-based **Microsoft Entra ID (Azure AD)** users.

The script is built for **large enterprise environments** (100k+ users), with support for Arabic content, error handling, and real-time progress output.

---

## 🔍 Key Features

* ✅ Hybrid user audit from both AD and Entra ID
* ✅ Real-time progress and per-user console output
* ✅ Avoids AD enumeration errors by paginating alphabetically (`a*`, `b*`, ..., `0*`)
* ✅ Arabic-safe CSV export using UTF-8 encoding
* ✅ Transcript logging for every execution
* ✅ Memory-optimized: Streams data directly to CSV

---

## 📁 Output Files

* **FullUserReport-YYYY-MM-DD\_HH-MM.csv**: Main audit report saved to user's Desktop (fallback: `C:\Temp`)
* **HybridUserAuditLog-YYYY-MM-DD\_HH-MM.txt**: Full transcript/log of the execution process

---

## 🔧 Attributes Collected

| Attribute                       | Description                             |
| ------------------------------- | --------------------------------------- |
| Username                        | AD `SamAccountName` or Entra prefix     |
| DisplayName                     | AD or Entra display name                |
| Department / Title / Email      | Sourced from AD or Entra                |
| InAD / InEntraID                | Shows presence of user in each system   |
| AD\_Enabled / Entra\_Enabled    | Shows if account is enabled             |
| AD\_Created / Entra\_Created    | Account creation dates                  |
| AD\_LastLogon                   | Last AD logon date                      |
| Entra\_LastInteractiveSignIn    | Last sign-in from Entra (interactive)   |
| Entra\_LastNonInteractiveSignIn | Last background/non-interactive sign-in |
| AD\_WhenChanged / PwdLastSet    | Account change history                  |
| AD\_Description / DN            | Extra AD metadata                       |

---

## 🛠 Requirements

* Windows PowerShell 5.1+
* Domain-joined machine with AD RSAT tools
* `Microsoft.Graph` module installed (`Install-Module Microsoft.Graph -Scope CurrentUser`)
* Permissions to query both AD and Microsoft Graph

---

## 🚀 How to Run

1. Open PowerShell as Administrator
2. Run the script:

```powershell
.\HybridUserAudit.ps1
```

> 🔒 Microsoft Graph authentication prompt will appear on first run.

---

## ✨ Example Output

```csv
Username,DisplayName,Department,Title,Email,InAD,AD_Enabled,AD_Created,...
m.ahmad,محمد أحمد,IT,Admin,m.ahmad@domain.com,Yes,Enabled,2021-06-12,...
```

Console:

```
[EntraID] 1 - m.ahmad : محمد أحمد
[✔] 1 - m.ahmad : محمد أحمد
```

---

## 📌 Notes

* Entra ID data is cached in a hashtable for efficient lookup.
* AD users are processed alphabetically to avoid `invalid enumeration context` errors.
* The script skips over failed lookups without stopping execution.

---

## 📄 License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT)

---

## Disclaimer
⚠ **Use this script at your own risk!**  
Always test in a **non-production environment** before applying changes to **Active Directory**.
