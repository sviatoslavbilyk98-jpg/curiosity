@echo off
echo Enabling WSL features...
dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart
dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart
echo.
echo Installing Ubuntu...
wsl --install -d Ubuntu
echo.
echo Done! Please restart your computer.
pause
