# Function to check if the script is running as an administrator
function Get-AdminRights {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# If the script is not running as an administrator, restart it with administrator rights
if (-not (Get-AdminRights)) {
    $newProcess = Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -PassThru
    $newProcess.WaitForExit()
    exit
}

# Search for available updates
$Updates = Get-WindowsUpdate -MicrosoftUpdate

# Check if updates are available
if ($Updates.Count -gt 0) {
    # Download and install updates
    Write-Host "Downloading and installing updates..."

    # Progress bar for downloads
    $totalProgress = 0
    foreach ($Update in $Updates) {
        $updateProgress = $Update.DownloadProgress * 100
        Write-Progress -Activity "Downloading" -Status "Update $Update.Title: $updateProgress%" -PercentComplete $totalProgress
        $totalProgress += $updateProgress
    }

    # Install updates
    Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -AutoReboot

    # Restart the computer if necessary
    if ($Updates.RequiresReboot) {
        Write-Host "Restarting the system..."
        Restart-Computer -Force
    }
} else {
    # No updates available
    Write-Host "No new updates found."
}
