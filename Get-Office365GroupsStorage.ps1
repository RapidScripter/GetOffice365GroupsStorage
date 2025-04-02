<#
=============================================================================================
Name:           Get storage used by Office 365 groups
Description:    This script finds Office 365 groups' storage size and exports the report to a CSV file.

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Supports modern authentication (MFA & non-MFA accounts).
2. Automatically installs and updates required modules (EXO V2 & PnP PowerShell).
3. Credentials can be passed as parameters for scheduling.
4. Error handling & validation included.
5. Provides real-time progress updates.
=============================================================================================
#>

# PARAMETERS
param ( 
   [Parameter(Mandatory = $false)]
   [Switch] $NoMFA,
   [String] $UserName = $null, 
   [String] $Password = $null,
   [String] $TenantName = $null #(Example: For 'contoso.com', enter 'contoso')
)

# Function to check and install PowerShell modules
function Ensure-Module {
    param (
        [string]$ModuleName,
        [string]$RequiredVersion = "0.0.0"
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName module is missing. Installing..." -ForegroundColor Magenta
        Install-Module $ModuleName -Repository PSGallery -Force -AllowClobber
    }
    elseif ((Get-Module -ListAvailable -Name $ModuleName).Version -lt $RequiredVersion) {
        Write-Host "$ModuleName is outdated. Updating..." -ForegroundColor Yellow
        Update-Module $ModuleName -Force
    }
    Import-Module $ModuleName -Force
}

# Ensure required modules are installed
Ensure-Module -ModuleName "PnP.PowerShell" -RequiredVersion "2.0.0"
Ensure-Module -ModuleName "ExchangeOnlineManagement" -RequiredVersion "2.0.5"

# Connecting to Exchange Online & SharePoint PnP PowerShell
Write-Host "Connecting to Office 365 services..." -ForegroundColor Cyan
if ($NoMFA.IsPresent) {
    if (($UserName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
    } else {
        $Credential = Get-Credential -Credential $null
    }
    if ($TenantName -eq "") {
        $TenantName = Read-Host "Enter your Tenant Name (e.g., 'contoso')"
    }
    $AdminUrl = "https://$TenantName-admin.sharepoint.com"
    Connect-PnPOnline -Url $AdminUrl -Credentials $Credential
    Connect-ExchangeOnline -Credential $Credential
} else {
    $TenantName = Read-Host "Enter your Tenant Name (e.g., 'contoso')"
    $AdminUrl = "https://$TenantName-admin.sharepoint.com"
    Connect-PnPOnline -Url $AdminUrl -Interactive
    Connect-ExchangeOnline
}

# Fetching Office 365 Group Storage Data
Write-Host "Retrieving Office 365 Groups storage usage..." -ForegroundColor Cyan
$OutputCsv = "./Office365GroupsStorageReport_$((Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')).csv"
$GroupSites = Get-PnPTenantSite -GroupIdDefined $true | Select-Object StorageUsageCurrent, StorageQuota, Url
$GroupCount = 0

Get-UnifiedGroup -ResultSize Unlimited | ForEach-Object {
    $GroupName = $_.DisplayName
    Write-Progress -Activity "Processed Group Count: $GroupCount" "Currently Processing Group: $GroupName"
    $SharePointSiteUrl = $_.SharePointSiteUrl
    
    if ($SharePointSiteUrl) {
        $GroupSite = $GroupSites | Where-Object { $_.Url -eq $SharePointSiteUrl }
        $StorageUsed = if ($GroupSite) { [math]::round($GroupSite.StorageUsageCurrent / 1024, 4) } else { "N/A" }
        $StorageLimit = if ($GroupSite) { $GroupSite.StorageQuota / 1024 } else { "N/A" }
    } else {
        $StorageUsed = "Group not used yet"
        $StorageLimit = "Group not used yet"
    }
    
    $GroupStorage = [PSCustomObject]@{
        'Group Name' = $GroupName
        'Group Email' = $_.PrimarySmtpAddress
        'Group Privacy' = $_.AccessType
        'Storage Used (GB)' = $StorageUsed
        'Storage Limit (GB)' = $StorageLimit
        'Created On' = $_.WhenCreated
    }
    
    $GroupStorage | Export-Csv -Path $OutputCsv -NoTypeInformation -Append
    $GroupCount++
}

# Display group count summary
if ($GroupCount -ne 0) {
    Write-Host "$GroupCount Office 365 groups found in this organization." -ForegroundColor Green
} else {
    Write-Host "No Office 365 groups found." -ForegroundColor Red
}

# Open the report after execution
if (Test-Path -Path $OutputCsv) {
    Write-Host "Report generated: " -NoNewline -ForegroundColor Yellow; Write-Host $OutputCsv
    Write-Host "\n~~ Script prepared by Kashyap Patel ~~\n" -ForegroundColor Green
    Write-Host "Check out " -NoNewline -ForegroundColor Green; Write-Host "https://github.com/RapidScripter" -ForegroundColor Yellow -NoNewline
    Write-Host " for more Microsoft 365 PowerShell scripts. \n\n" -ForegroundColor Green
    
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open the output file?", 0, "Open Output File", 4)
    If ($UserInput -eq 6) {
        Invoke-Item "$OutputCsv"
    }
}

# Cleanup sessions
Disconnect-PnPOnline
Get-PSSession | Remove-PSSession
