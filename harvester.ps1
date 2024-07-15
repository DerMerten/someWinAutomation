############################################################# DONT TOUCH #############################################################

# for pop-up
Add-Type -AssemblyName PresentationFramework

############################################################# NEED to EDIT #############################################################

# Create CSV for gathered data
$csvPathPc = 'C:\Path\To\Your\PCFile.csv'
$csvPathPrinter = 'C:\Path\To\Your\printerInfo.csv'

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

    $selectedInfoProp = @(
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
        "OsName", # for the printer-part
        "OSDisplayVersion",
        "OsBuildNumber",
        "Win Updates",
        "Defender AV Def"
        )

    # Get PC-Infos
    Write-Host "getting PC-Infos..."
    $pcInfo = Get-ComputerInfo | Select-Object $selectedInfoProp

    if (-not (Test-Path $csvPathPc)) {
        $pcInfo | Export-Csv -Path $csvPathPc -NoTypeInformation
    } else {
        $pcInfo | Export-Csv -Path $csvPathPc -Append -NoTypeInformation
    }
    Write-Host "PC-Infos successfully saved in a CSV-File!"
}
catch {

    #get exception and print it in a messagebox
    $message = $_
    [System.Windows.MessageBox]::Show("Something went wrong getting PC-Infos: $message", "PCINFOS-ERROR")
    exit
}

############################################################# GET PRINTERS #############################################################

# simple try-catch again
try {

    # Get all printers
    $printers = Get-Printer

    # Check if printers exist
    if ($printers -eq 0) {
        Write-Host "No printers found"
        return
    }

    # Initialize a list for printer information
    $printerInfoList = @()

    # Retrieve port information for each printer
    foreach ($printer in $printers) {
        $printerName = $printer.Name
        $driverName = $printer.DriverName
        $portName = $printer.PortName
        $shared = $printer.Shared

        # Skip default MS printer / i dont wanna see it
        if ($printerName -eq "OneNote (Desktop)" -or $printerName -eq "Microsoft Print to PDF" -or $printerName -eq "Microsoft XPS Document Writer" -or $printerName -eq "Fax") {
            continue
        }

        # Get port information
        $port = Get-PrinterPort -Name $portName

        # Create an object with printer name and port description
        $printerInfo = [PSCustomObject]@{
            PrinterName     = $printerName
            DriverName      = $driverName
            PortName        = $portName
            PortDescription = $port.Description
            Shared          = $shared
        }

        # Add the object to the list
        $printerInfoList += $printerInfo
    }

    # Write the printer information to the CSV file
    if (-not (Test-Path $csvPathPrinter)) {
        $printerInfoList | Export-Csv -Path $csvPathPrinter -NoTypeInformation
    } else {
        $printerInfoList | Export-Csv -Path $csvPathPrinter -Append -NoTypeInformation
    }

    Write-Host "Printers successfully saved in a CSV-File!"
}
catch {

    # Get exception and display it in a messagebox
    $message = $_
    [System.Windows.MessageBox]::Show("Something went wrong getting printers: $message", "PRINTING-ERROR")
    exit
}

############################################################# GET UPDATES #############################################################

#simple try-catch again 
try {

    # Install PSWindowsUpdate Module
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Host "PSWindowsUpdate module not found. Installing..."
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser
    } else {
        Write-Host "PSWindowsUpdate module is already installed."
    }

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