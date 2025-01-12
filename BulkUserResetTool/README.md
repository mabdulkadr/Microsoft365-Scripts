
# Bulk User Access Control Script

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/powershell-7.0%2B-blue.svg)
![Version](https://img.shields.io/badge/version-1.10-green.svg)

## Overview

This PowerShell script is designed to streamline the management of user accounts in Microsoft 365 environments. It enables administrators to perform bulk actions such as resetting user passwords, disabling devices, signing out active sessions, and blocking sign-ins. By leveraging Microsoft Graph PowerShell, the script ensures secure and efficient execution of these critical tasks.

🔗 Original article: [Force sign out users in Microsoft 365](https://www.alitajran.com/force-sign-out-users-microsoft-365/)

---

## Features

- **Reset Passwords**: Reset user passwords with an option to force password change at next sign-in.
- **Sign-Out Sessions**: Revoke all active sessions and refresh tokens for specified users.
- **Block Sign-In**: Prevent selected users from signing in.
- **Disable Devices**: Disable devices registered under a user's account.
- **Exclude Users**: Skip actions for specific users using the `-Exclude` parameter.
- **Bulk or Targeted Actions**: Perform actions for all users or specified users via `-All` or `-UserPrincipalNames` parameters.

---

## Parameters

- `-All`  
  Perform actions for all users in the directory.

- `-ResetPassword`  
  Reset passwords for selected users.

- `-DisableDevices`  
  Disable devices associated with selected users.

- `-SignOut`  
  Revoke all active sessions and refresh tokens.

- `-BlockSignIn`  
  Block sign-ins for selected users.

- `-Exclude`  
  Exclude specific users from bulk actions.

- `-UserPrincipalNames`  
  Specify the user principal names (UPNs) of users to target.

---

## How to Use

### Prerequisites

1. Install the **Microsoft.Graph** PowerShell module:
   ```powershell
   Install-Module -Name Microsoft.Graph -Scope CurrentUser
   ```
2. Connect to Microsoft Graph API with the required permissions:
   ```powershell
   Connect-MgGraph -Scopes Directory.AccessAsUser.All
   ```

### Script Execution

1. Clone or download this repository.
2. Open PowerShell as an administrator.
3. Run the script with desired parameters. Example:
   ```powershell
   .\Set-SignOut.ps1 -All -SignOut -BlockSignIn
   ```

### Example Commands

- Reset passwords for specific users:
  ```powershell
  .\Set-SignOut.ps1 -UserPrincipalNames user1@domain.com, user2@domain.com -ResetPassword
  ```
- Disable devices and sign out all sessions for all users:
  ```powershell
  .\Set-SignOut.ps1 -All -DisableDevices -SignOut
  ```
- Block sign-ins and exclude specific users:
  ```powershell
  .\Set-SignOut.ps1 -All -BlockSignIn -Exclude admin@domain.com, support@domain.com
  ```

---

## Outputs

- **Success Messages**: Displays actions completed for each user or device.
- **Error Messages**: Highlights users or devices that could not be processed.
- **Exclusions**: Confirms which users were excluded from processing.

---

## Notes

- Ensure you have the necessary permissions to perform these actions in Microsoft Entra ID (formerly Azure AD).
- Test the script in a non-production environment before deploying it in production.

---

## Change Log

- **V1.00** (06/18/2023): Initial release.
- **V1.10** (07/24/2023): Updated for Microsoft Graph PowerShell changes.

---

## Author

- **Ali Tajran**  
  Website: [alitajran.com](https://www.alitajran.com)  
  LinkedIn: [linkedin.com/in/alitajran](https://linkedin.com/in/alitajran)

---

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

**Disclaimer**: Use this script responsibly. The author is not liable for any unintended consequences.
```

This `README.md` file provides a detailed and professional overview of your script, covering its features, parameters, and usage. Let me know if you need further adjustments!