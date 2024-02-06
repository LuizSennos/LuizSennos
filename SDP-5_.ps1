function Select-TextItem 
{ 
    PARAM  
    ( 
        [Parameter(Mandatory=$true)] 
        $options, 
        $displayProperty 
    ) 

    [int]$optionPrefix = 1 
    # Create menu list 
    foreach ($option in $options) 
    { 
        if ($displayProperty -eq $null) 
        { 
            Write-Host ("{0,3}: {1}" -f $optionPrefix,$option) 
        } 
        else 
        { 
            Write-Host ("{0,3}: {1}" -f $optionPrefix,$option.$displayProperty) 
        } 
        $optionPrefix++ 
    } 
    Write-Host ("{0,3}: {1}" -f 0,"To cancel")  
    [int]$response = Read-Host "Select which server to Shadow" 
    $val = $null 
    if ($response -gt 0 -and $response -le $options.Count) 
    { 
        $val = $options[$response-1] 
    } 
    return $val 
}    

# Read CSV file and create object array
$vmList = Import-Csv -Encoding UTF8 -Path C:\Temp\vms.csv | ForEach-Object {
    [PSCustomObject]@{
        VM = $_.vm
        IP = $_.ip
    }
}

# Create menu options from object array
$menuOptions = $vmList | ForEach-Object {
    "$($_.VM) "
}

# Main loop
do {
    # Let user select menu option
    $selectedOption = Select-TextItem $menuOptions

    if ($selectedOption) {
        # Get IP address from selected option
        $selectedVM = $vmList | Where-Object { "$($_.VM) " -eq $selectedOption }
        $selectedIP = $selectedVM.IP

        Write-Host "Connecting to $selectedOption"
        Write-Host "The IP address of the selected VM is $selectedIP"

        if ($selectedIP -eq "Other")
        { 
            [string]$response = Read-Host "Insert the name of computer" 
            $val = $response 
            query.exe session /server:$val
            $sessionid=Read-Host -Prompt "Select the Session Number"
            Mstsc.exe /V:$val /shadow:$sessionid /noConsentPrompt
        }
        else
        {
            query.exe session /server:$selectedIP
            $sessionid=Read-Host -Prompt "Select the Session Number"
            Mstsc.exe /V:$selectedIP /shadow:$sessionid /noConsentPrompt
        }

        # Prompt user to return to main menu or exit
        do {
            $choice = Read-Host "Press 'M' to return to main menu or 'E' to exit"
        } while ($choice -notmatch "[ME]")

        if ($choice -eq "E") {
            break
        }
    }
} while ($true)
