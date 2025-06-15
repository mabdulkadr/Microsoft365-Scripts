
# Guest User Management Scripts
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Version](https://img.shields.io/badge/version-1.10-green.svg)

A PowerShell toolkit for Microsoft Entra ID (Azure AD) guest user lifecycle management:  
- **Export** detailed reports of guest users and their memberships  
- **Bulk-remove** guest users based on a provided list

---

## 📦 Overview

Microsoft Entra ID (Azure AD) tenants can accumulate a large number of guest accounts, many of which may no longer be needed.  
This toolkit provides PowerShell scripts to help IT administrators audit, document, and remove guest users efficiently and safely.

---

## 🗂️ Scripts Included

### 1. `GuestUserReport.ps1`

**Description:**  
Exports a comprehensive CSV report of all guest users in your Microsoft Entra ID tenant.  
The report includes:

- Display Name
- User Principal Name (UPN)
- Email address
- Company name
- Invitation and redemption status
- Account creation date
- Group memberships

**Key Features:**

- Connects to Microsoft Graph securely (prompts for authentication)
- Fetches all guest users (`UserType = Guest`)
- Optionally filters by account age (e.g., show only stale or recently created guests)
- Collects group membership details for each guest
- Saves the report in `C:\Temp\GuestUserReport_<timestamp>.csv`
- Notifies and prompts you to open the exported file

**How to Use:**

```powershell
# Run in a PowerShell window with Microsoft.Graph module installed
.\GuestUserReport.ps1
````

---

### 2. `Remove-GuestUsersFromCSV.ps1.ps1`

**Description:**
Bulk-removes guest user accounts from your Entra ID tenant using a provided CSV file.
The script matches each entry in the CSV by `UserPrincipalName`, confirms the account is a guest (`UserType = Guest`), and deletes it.

**Key Features:**

* Interactive file dialog to select your CSV file
* Reads a list of guest users from the CSV file
* For each row:

  * Looks up the account in Entra ID using `UserPrincipalName`
  * Confirms the user is of type `Guest`
  * Deletes the guest account (skips members)
  * Logs each action (deleted, skipped, not found, or failed)
* Prints a summary to the console at the end

**CSV Requirements:**

* The input CSV **must** have at least these columns:

  * `DisplayName` (for your reference)
  * `UserPrincipalName` (UPN, e.g., `m.abdelkader_upm.xyz.com#EXT#@abc.onmicrosoft.com`)
* You may include additional columns as needed; they will be ignored.

**Example CSV:**

```csv
DisplayName,UserPrincipalName
m.abdelkader,m.abdelkader_upm.xyz.com#EXT#@abc.onmicrosoft.com
372113519,372113519_cloud.sa#EXT#@abc.onmicrosoft.com
...
```

**How to Use:**

```powershell
# Run in a PowerShell window with Microsoft.Graph module installed
.\Remove-GuestUsersFromCSV.ps1.ps1
```

* Select your CSV file when prompted.
* Watch the console for status and summary.

---

## 🚦 Requirements

* **PowerShell 5.1+** (Windows 10/11 or Windows Server 2016+ recommended)
* **Microsoft Graph PowerShell SDK** installed
  (Install via: `Install-Module Microsoft.Graph -Scope CurrentUser`)
* **Entra ID admin credentials** to connect and perform delete operations

---

## 📝 Script Output

* **GuestUserReport.ps1:**

  * CSV report in `C:\Temp\GuestUserReport_<timestamp>.csv`
  * File contains detailed guest user information and group memberships
  * Option to open the file after completion

* **Remove-GuestUsersFromCSV.ps1.ps1:**

  * Console summary showing:

    * Number of users processed
    * Number of guests deleted
    * Number of entries skipped (not guests)
    * Number of entries not found or failed

---

## 📋 Best Practices

* **Test with a small set** of users or a non-production tenant before running bulk operations
* Always keep a backup of your original CSV files and reports
* Review the exported report to confirm user identity and account status before deletion
* Use the report as an audit/compliance record if needed

---

## 📜 License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## ❗ Disclaimer

These scripts are provided as-is.
**Test them thoroughly in a staging or lab environment before running in production.**
The author is not responsible for unintended consequences or data loss.

---
