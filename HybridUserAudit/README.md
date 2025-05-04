
# HybridUserAudit.ps1

![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Status](https://img.shields.io/badge/status-Stable-success)

## Overview

`HybridUserAudit.ps1` is a PowerShell script designed to audit and compare user accounts across **on-premises Active Directory** and **Microsoft Entra ID (Azure AD)**. It generates a unified, timestamped CSV report that includes key attributes from both environments, supporting multilingual environments including Arabic.

---

## Features

- 🔁 Merges user data from AD and Entra ID by `username`
- 📊 Highlights whether users exist in AD, Entra ID, or both
- 🕒 Tracks:
  - Account creation date
  - Last logon
  - Last password set
  - Last account modification
- ✅ Indicates enabled/disabled status in each directory
- 📁 Exports UTF-8 BOM encoded CSV to Desktop (Arabic-compatible)
- 🕓 Appends timestamp to filename
- 📂 Automatically opens the Desktop folder after saving

---

## Prerequisites

- PowerShell 5.1 or higher
- ActiveDirectory PowerShell module
- Microsoft Graph PowerShell SDK

Install Graph module if needed:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
````

---

## How to Use

```powershell
.\HybridUserAudit.ps1
```

> The script will:
>
> 1. Authenticate with Microsoft Graph (interactive sign-in).
> 2. Query all users from Active Directory and Entra ID.
> 3. Merge and export the report to your Desktop.

---

## Output Details

* 📄 Filename: `FullUserReport-YYYY-MM-DD_HH-MM.csv`
* 📌 Columns:

  * `Username`
  * `DisplayName`, `Department`, `Title`, `Email`
  * `InAD`, `AD_Enabled`, `AD_Created`, `AD_LastLogon`, `AD_WhenChanged`, `AD_PwdLastSet`, `AD_Description`, `AD_DistinguishedName`
  * `InEntraID`, `Entra_Enabled`, `Entra_Created`, `Entra_LastInteractiveSignIn`, `Entra_LastNonInteractiveSignIn`

---

## License

This script is released under the [MIT License](https://opensource.org/licenses/MIT).

---

**Disclaimer:** Always test in a development environment before use in production. Provided as-is with no warranties.
