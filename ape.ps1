# Use TLS 1.2 for downloads
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Allow running downloaded scripts in this session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Make sure NuGet provider is installed
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force

# Trust the PowerShell Gallery to avoid prompts
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

# Update PowerShellGet so Install-Script behaves consistently
Install-Module PowerShellGet -Force
Import-Module PowerShellGet

# Install the Autopilot script
Install-Script -Name Get-WindowsAutopilotInfo -Force

# Locate the installed script path and run it online
$script = Get-Command Get-WindowsAutopilotInfo.ps1 -ErrorAction SilentlyContinue
if (-not $script) {
    $possible = @(
        "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1",
        "$env:ProgramFiles\WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1",
        "$env:ProgramFiles\PowerShell\Scripts\Get-WindowsAutopilotInfo.ps1"
    )
    foreach ($p in $possible) { if (Test-Path $p) { $script = Get-Item $p; break } }
}
& $script.FullName -Online
