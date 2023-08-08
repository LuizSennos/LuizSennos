<#
Script: Check_VM.ps1
Creator: Luiz Sennos
Date: July 03, 2022
Description: This script performs the following tasks:
  - Checks for and installs required PowerShell modules (SharePointPnPPowerShellOnline, ImportExcel).
  - Imports data from an Excel spreadsheet located at "C:\Temp\vms.xlsx".
  - Tests connectivity to hosts specified in the spreadsheet using both hostname and IP address.
  - Connects to a SharePoint site using provided credentials and downloads a Users list file.
  - Removes a specific file ("Vmstest.xlsx") from the SharePoint document library.
  - Exports the results of host connectivity tests to an Excel spreadsheet at "C:\path\to\Vmstest.xlsx".
  - Uploads the results spreadsheet to a specific location in the SharePoint site.
  - Cleans up temporary files.
  - Displays a prompt for the user to exit.

Note: Replace placeholders such as "URL here" with actual URLs or paths.
#>






# For Each Modules we need 
$modulesArray = @(
    "SharePointPnPPowerShellOnline",
    "ImportExcel"
)

# Loop through the array modules and install if not installed
foreach ($mod in $modulesArray) {
    if (Get-Module -ListAvailable $mod) {
        # Module exists
        Write-Host "Module '$mod' is already installed"
    } else {
        # Module does not exist, install it
        Write-Host "Installing '$mod'"
        Install-Module $mod -Scope CurrentUser -ErrorAction Stop
    }
}

# Import data from Excel spreadsheet into a PowerShell variable
$data = Import-Excel -Path "C:\Temp\vms.xlsx"

# Initialize an empty array to store results
$results = @()

# Loop through each row of the spreadsheet
foreach ($row in $data) {
    # Test ping connection for hostname in current row
    $accessibleByHostname = Test-Connection -ComputerName $row.Hostname -Count 1 -Quiet

    # Test IP connection for IP in current row
    $accessibleByIP = Test-Connection -ComputerName $row.IP -Count 1 -Quiet

    # Add results to results array
    $results += [pscustomobject]@{
        Hostname = $row.Hostname
        IP = $row.IP
        "Accessible by Hostname" = $accessibleByHostname
        "Accessible by IP" = $accessibleByIP
    }
}

# Connect to SharePoint with credentials and download the Users list file
$credentials = Get-Credential -Message "Please Enter SharePoint Online credentials"
$Site = "URL here"

try {
    Connect-PnPOnline -Url $Site -Credentials $credentials -WarningAction Ignore
    if (-not (Get-PnpContext)) {
        Write-Host "Unable to connect to SharePoint"
        return
    }
} catch {
    Write-Host "Error connecting to SharePoint Online: $_.Exception.Message" -foregroundcolor black -backgroundcolor Red
    return
}

# Remove the "Vmstext.xlsx" from SharePoint
Remove-PnPFile -ServerRelativeUrl "/teams/RPA/Shared Documents/RPA Support/Automation/Robots/PowerAutomate/Project 06 - VMs Daily Test Connection/Vmstest.xlsx" -force

# Export results to Excel spreadsheet
$results | Export-Excel -Path "C:\path\to\Vmstest.xlsx" -WorksheetName "Table1" -AutoSize -AutoFilter -TableStyle "Medium2"
$finalPath = "C:\path\to\Vmstest.xlsx"
$File = Get-ChildItem $finalPath

Write-Host "Saving to SharePoint"

Add-PnPFile -Folder "Shared Documents/RPA Support/Automation/Robots/PowerAutomate/Project 06 - VMs Daily Test Connection" -Path $File.FullName

# Remove all temp files from C:\Temp (Users and ADExport)
Remove-Item $finalPath

# Prompt to exit
Read-Host -Prompt "Press Enter to exit"
Pause
