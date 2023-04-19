<#
.SYNOPSIS
This PowerShell script updates Active Directory user passwords and exports user properties to Excel, then uploads the Excel file to SharePoint.

.DESCRIPTION
This script performs the following tasks:
1. Checks for the existence of required PowerShell modules and installs them if necessary.
2. Collects user properties from an input CSV file and exports them to a local CSV file.
3. Connects to SharePoint Online, downloads a CSV file from SharePoint, and extracts user properties from it.
4. Creates an Excel table with the extracted user properties, calculates remaining days for password expiration, and saves the Excel file locally.
5. Uploads the Excel file to a specified location in SharePoint.
6. Removes temporary files created during the script execution.

.PARAMETER None
This script does not require any parameters to be passed.

.EXAMPLE
.\Update-ADUserPassword.ps1
Runs the script with default settings.

.NOTES
- This script requires SharePointPnPPowerShellOnline, ActiveDirectory, and ImportExcel modules to be installed.
- SharePoint Online credentials and SharePoint site URL need to be provided.
- Input CSV file with user properties should be available in SharePoint.
- Temporary files are stored in C:\Temp directory and are removed at the end of the script execution.
- This script should be run with appropriate permissions to update Active Directory user passwords and access SharePoint.

.AUTHOR
Luiz Sennos
The author of this script.

#>


Write-Host "Transaction started" -BackgroundColor Green  -ForegroundColor Black

Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -force

#Checks if Modules Exists #Gets Tls12 protocol

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12  #Checks If TLS protocol available

# For Each Modules we need 
    $modulesArray = @(
        "SharePointPnPPowerShellOnline",
        "ActiveDirectory",
        "ImportExcel"
    )

    #  Loop through the array modules and installs if not installed
    foreach($mod in $modulesArray) {
        if(Get-Module -ListAvailable $mod) {
            # Module exists
            Write-Host "Module '$mod' is already installed"
        } else {
            # Module does not exist, install it
            Write-Host "Installing '$mod '"
            Install-Module $mod -Scope CurrentUser -ErrorAction Stop
        }
    }
    
  

#1 - Colects Users properties from InpuUsers and exports to excel

#Variables
$credentials = Get-Credential -Message “Please Enter SharePoint Online credentials” 
$Site= ”Enter SharePoint Site here"
$TimeNow = Get-Date
$TimeExe = (get-date -f yyyyMMdd_HHmm)
$AdExport = "C:\Temp\AD_Updated $TimeExe.csv"  #Change the temp directory if wanted


 
#Conects to SharePoint with credentials and download the Users list file
 try {
    Connect-PnPOnline -Url $Site -Credentials $credentials -WarningAction Ignore 
    if(-not (get-PnpContext)) {
    Write-Host "Unable to conect to SharePoint"
    return
} }
 catch { 
     Write-Host "Error connecting to SharePoint Online: $_.Exception.Message" -foregroundcolor black -backgroundcolor Red
    return
}

#endregion ConnectPnPOnline



Write-Host "Connected with success" 

Get-PnPFile "Sharepoint path of .csv input file here" -Path "C:\Temp" -FileName "name of the file .csv" -AsFile -Force   


#ForEach row in the download user files, finds and selects its properties.
Import-Csv -Encoding UTF8 -path C:\Temp\users.csv |`

ForEach-Object {
     
 #Get-aduser - gets the properties from the InputUsers csv downloaded file
   
 get-aduser -Identity $_.user -Properties samaccountname,PasswordLastSet, "msDS-UserPasswordExpiryTimeComputed", GivenName  |
  
   
 Select-Object samaccountname,PasswordLastSet, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")  } 
 }, GivenName  |

 
 Export-Csv -Encoding UTF8 $AdExport -Append  -NoTypeInformation -Force  |Format-Table -AutoSize 
 
 }
 
 Write-Host "AD User saved to local temp storage." -BackgroundColor Green -ForegroundColor Black  
 
 Write-Host "Creating Table1" -BackgroundColor Green -ForegroundColor Black  

 ################################
#2 - Exports to Sharepoint Step
    
#CREATES TABLE AND EXPORTS TO XLSX#
#Define locations and delimiter

$params = @{
    AutoSize      = $true
    TableName     = 'Table1'
    TableStyle    = 'Light2' # => Choose the style
    BoldTopRow    = $true
    WorksheetName = 'Sheet1'
    PassThru      = $true
    Path          = "C:\Temp\AD_Updated $TimeExe.xlsx" # => Change for wanted -
}

$xlsx = Import-Csv $AdExport | Export-Excel @params
$ws   = $xlsx.Workbook.Worksheets[$params.Worksheetname]
$ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
Close-ExcelPackage $xlsx 

#Inserts RemainingDays Formula Into Sheet
(Get-ChildItem "C:\Temp\AD_Updated $TimeExe.xlsx")| #  Same as Path from params
    foreach-object {
        $xl=New-Object -ComObject Excel.Application
        $wb=$xl.workbooks.open($_)
        $ws = $wb.worksheets.Item(1)
        $ws.Cells.Item(1,5) ='RemainingDays'
        
        $ws.Cells.Item(2,5) = "=INT(C2)-TODAY()"
        $ws.Activate()
        $xl.Columns.item('e').NumberFormat = "0"

        $wb.Save()
        $xl.Quit()
        }
    

    $finalPath = "C:\Temp\AD_Updated $TimeExe.xlsx" #Change if wanted

  
#Gets xlsx final file and exports to Sharepoint 
$File = Get-ChildItem $finalPath

Write-Host "File's Ready. Saving to SharePoint" 

    Add-PnPFile -Folder "Shared Documents/path to upload" -Path $File.FullName #change path to wanted

Write-Host "File $finalPath Uploaded to Sharepoint" 


#Remove all temp Files from c/temp (Users and ADExport)

Write-Host "Removing temp files" -BackgroundColor Yellow -ForegroundColor Black

Remove-Item C:\Temp\Users.csv
Remove-Item $AdExport
Remove-Item $finalPath

Write-Host "Execution Finished. AD Password Updated with Success" -BackgroundColor Green -ForegroundColor Black

Read-Host -Prompt "Press Enter to exit"

pause
