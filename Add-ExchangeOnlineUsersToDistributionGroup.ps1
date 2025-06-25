<#
.SYNOPSIS
    This script adds multiple users from a CSV file to a specified distribution group in Exchange Online.

.DESCRIPTION
    The script reads user email addresses from a CSV file and adds each user to a specified distribution group in Exchange Online. 
    It ensures that the necessary modules are installed, prompts the user to select the CSV file, and to provide the Group ID or name.
    It also asks the user if they want to save a log file and, if confirmed, prompts for the location to save the log file.

.REQUIREMENTS
    - ExchangeOnlineManagement Module
    - Microsoft 365 admin privileges

.NOTES
    The CSV file should have a header "Email".

.EXAMPLE
    .\Add-ExchangeOnlineUsersToDistributionGroup.ps1
#>

# Ensure necessary modules are installed
function Ensure-Module {
    param (
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Install-Module -Name $ModuleName -Force
    }
    Import-Module $ModuleName
}

# Ensure ExchangeOnlineManagement module is installed
Ensure-Module -ModuleName "ExchangeOnlineManagement"

# Import necessary types for dialogs
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# Connect to Exchange Online
function Connect-ExchangeOnlineSafely {
    try {
        Connect-ExchangeOnline -UserPrincipalName (Get-Credential).UserName
    } catch {
        Write-Host "Failed to connect to Exchange Online. Exiting."
        exit
    }
}
Connect-ExchangeOnlineSafely

# Function to open a file dialog to select the CSV file
function Get-CSVFilePath {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $OpenFileDialog.FilterIndex = 1
    $OpenFileDialog.Multiselect = $false
    
    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $OpenFileDialog.FileName
    } else {
        Write-Host "No file selected. Exiting."
        exit
    }
}

# Function to open a save file dialog for the log file
function Get-LogFilePath {
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $SaveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $SaveFileDialog.FilterIndex = 1

    if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $SaveFileDialog.FileName
    } else {
        Write-Host "No save file selected. Exiting."
        exit
    }
}

# Function to prompt user for input
function Get-UserInput {
    param (
        [string]$message
    )
    return [Microsoft.VisualBasic.Interaction]::InputBox($message, "Input Required")
}

# Function to prompt the user with Yes/No dialog
function Get-UserConfirmation {
    param (
        [string]$message
    )
    $result = [System.Windows.Forms.MessageBox]::Show($message, "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

# Get the CSV file path
$csvPath = Get-CSVFilePath

# Prompt user for Distribution Group Name
$groupName = Get-UserInput -message "Please enter the Distribution Group Name:"

if ([string]::IsNullOrWhiteSpace($groupName)) {
    Write-Host "No Distribution Group Name provided. Exiting."
    exit
}

# Import the CSV file
try {
    $users = Import-Csv -Path $csvPath
} catch {
    Write-Host "Failed to import CSV file. Exiting."
    exit
}

# Prepare log file
$log = @()

# Function to log and format the output
function Log-Result {
    param (
        [string]$Status,
        [string]$Email,
        [string]$Message
    )
    $log += [PSCustomObject]@{
        Status = $Status
        Email = $Email
        Message = $Message
    }
}

# Add users to the distribution group
foreach ($user in $users) {
    $userPrincipalName = $user.Email
    try {
        # Validate email format
        if (-not [System.Text.RegularExpressions.Regex]::IsMatch($userPrincipalName, "^[^@\s]+@[^@\s]+\.[^@\s]+$")) {
            throw "Invalid email format: $userPrincipalName"
        }

        # Add user to distribution group
        Add-DistributionGroupMember -Identity $groupName -Member $userPrincipalName
        
        # Log success
        Log-Result -Status "SUCCESS" -Email $userPrincipalName -Message "Added to the distribution group."
    } catch {
        # Log failure
        Log-Result -Status "FAILURE" -Email $userPrincipalName -Message $_.Exception.Message
    }
}

# Display results in a table
$log | Format-Table -AutoSize

# Ask the user if they want to save the log file
if (Get-UserConfirmation -message "Do you want to save the log file?") {
    # Get the log file path
    $logFilePath = Get-LogFilePath

    # Save log to file in table format
    try {
        $log | Out-File -FilePath $logFilePath -Force
        Write-Host "Process completed. Log file saved to $logFilePath."
    } catch {
        Write-Host "Failed to save log file. Exiting."
    }
} else {
    Write-Host "Process completed. Log file not saved."
}
