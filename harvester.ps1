############################################################# DONT TOUCH #############################################################

# for pop-up
Add-Type -AssemblyName PresentationFramework

############################################################# NEED to EDIT #############################################################

# Create CSV for gathered data
$csvPath = 'C:\Path\to\your\file.csv'
############################################################# GET ADMIN #############################################################
# Function to check if the script is running as an administrator
function Get-AdminRights {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

#simple try-catch
try {
    # If the script is not running as an administrator, restart it with administrator rights
    if (-not (Get-AdminRights)) {
        $newProcess = Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -PassThru
        $newProcess.WaitForExit()
        exit
    }
}
catch {

    #get exception and print it in a messagebox
    $message = $_
    [System.Windows.MessageBox]::Show("Something went wrong getting sudo rights: $message", "ADMIN-RIGHTS")
    exit
}

############################################################# GET INFOS #############################################################

#simple try-catch again
try {

    $selectedProp = @(
        "Asset No",
        "csName",
        "CsPCSystemType",
        "Device Status",
        "CsManufacturer",
        "csModel",
        "CsSystemFamily", # just in case
        "OsOrganization",
        "Room",
        "Location",
        "OsLocalDateTime",
        "Last Known User",
        "Printer Connection",
        "Ip Address",
        "BiosSeralNumber"
        "OsName",
        "OSDisplayVersion",
        "OsBuildNumber",
        "Win Updates",
        "Defender AV Def"
        )

    # Get PC-Infos
    Write-Host "getting PC-Infos..."
    $pcInfo = Get-ComputerInfo | Select-Object $selectedProp

    if (-not (Test-Path $csvPath)) {
        $pcInfo | Export-Csv -Path $csvPath -NoTypeInformation
    } else {
        $pcInfo | Export-Csv -Path $csvPath -Append -NoTypeInformation
    }
    Write-Host "CSV successfully saved!"
}
catch {

    #get exception and print it in a messagebox
    $message = $_
    [System.Windows.MessageBox]::Show("Something went wrong with the csv-part: $message", "CSV-Error")
    exit
}

############################################################# GET UPDATES #############################################################

#simple try-catch again 
try {

    # Search for available updates
    Write-Host "checking for updates..."
    $getUpdates = Get-WindowsUpdate -MicrosoftUpdate

    # Check if updates are available
    if ($getUpdates.Count -gt 0) {

        # Download and install updates
        Write-Host "Downloading and installing updates..."

        # Install updates
        Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -AutoReboot

        # Restart the computer if necessary
        if ($getUpdates.RequiresReboot) {
            Write-Host "Restarting the system..."
          Restart-Computer -Force
        }
    } else {
        # No updates available
        Write-Host "No new updates found."
    }
}
catch {

    #get exception and print it in a messagebox
    $message = $_
    [System.Windows.MessageBox]::Show("Something went wrong during the update-part: $message", "Update-Error")
    exit
}
