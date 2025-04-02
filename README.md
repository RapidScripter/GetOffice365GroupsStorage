# GetOffice365GroupsStorage - Get Storage Used by Office 365 Groups

## Description
This PowerShell script retrieves the storage usage of Office 365 groups and exports the data to a CSV file. 

### Script Highlights
- Uses modern authentication to connect to Exchange Online.
- Supports execution with MFA-enabled accounts.
- Automatically installs the required PowerShell modules (`EXO V2` and `SharePoint PnP PowerShell`) upon confirmation.
- Accepts credentials as parameters for scheduler-friendly execution.
- Exports the report to a CSV file.
- Lists storage details of each Office 365 group.

## Prerequisites
- **PowerShell 5.1 or later**
- **Exchange Online PowerShell Module**
- **PnP PowerShell Module for SharePoint**
- **Microsoft 365 Admin Permissions** (Exchange Admin & SharePoint Admin)

## Installation
Before running the script, ensure the required PowerShell modules are installed. The script will prompt to install missing modules automatically.

To manually install the required modules, run:
```powershell
Install-Module ExchangeOnlineManagement -Repository PSGallery -Force
Install-Module PnP.PowerShell -Repository PSGallery -Force
```

## Usage
Run the script using one of the following methods:

### 1. With MFA Authentication
```powershell
.\Get-O365GroupsStorage.ps1
```
The script will prompt for authentication.

### 2. Without MFA (Passing Credentials as Parameters)
```powershell
.\Get-O365GroupsStorage.ps1 -NoMFA -UserName "admin@contoso.com" -Password "yourpassword" -TenantName "contoso"
```

### 3. Scheduled Execution (Non-MFA)
For automation, save credentials securely and pass them as parameters:
```powershell
.\Get-O365GroupsStorage.ps1 -NoMFA -UserName "admin@contoso.com" -Password "yourpassword" -TenantName "contoso"
```

## Output
The script generates a CSV file containing:
- Group Name
- Group Email
- Privacy Type
- Storage Used (GB)
- Storage Limit (GB)
- Created Date

### Example Output
```
Office365GroupsStorageSizeReport_Jul-01 10-30 AM.csv
```

After execution, the script prompts to open the generated CSV file.

## Notes
- Ensure you have the necessary permissions to retrieve SharePoint and Exchange data.
- If the storage values return "N/A", verify SharePoint permissions and connectivity.

## Author
**Kashyap Patel**  
[GitHub Profile](https://github.com/RapidScripter)
