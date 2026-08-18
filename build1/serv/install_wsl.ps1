$logFile = "C:\Users\Олександр\rm\serv\wsl_install.log"
"Starting WSL installation at $(Get-Date)" | Out-File $logFile

try {
    "Enabling Microsoft-Windows-Subsystem-Linux..." | Out-File $logFile -Append
    dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart 2>&1 | Out-File $logFile -Append

    "Enabling VirtualMachinePlatform..." | Out-File $logFile -Append
    dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart 2>&1 | Out-File $logFile -Append

    "Updating WSL..." | Out-File $logFile -Append
    wsl --update 2>&1 | Out-File $logFile -Append

    "Setting default version 2..." | Out-File $logFile -Append
    wsl --set-default-version 2 2>&1 | Out-File $logFile -Append

    "Installing Ubuntu..." | Out-File $logFile -Append
    wsl --install -d Ubuntu 2>&1 | Out-File $logFile -Append

    "Listing distros..." | Out-File $logFile -Append
    wsl --list --verbose 2>&1 | Out-File $logFile -Append

    "Done!" | Out-File $logFile -Append
} catch {
    "Error: $_" | Out-File $logFile -Append
}
