# Use TLS 1.2 for downloads
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Allow running the script in this session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Make sure the PSGallery repository is available
Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

# Install the Autopilot info script from the gallery
Install-Script -Name Get-WindowsAutopilotInfo -Force

# Find where it was installed, then run it
$script = Get-Command Get-WindowsAutopilotInfo.ps1 -ErrorAction SilentlyContinue
if (-not $script) {
    # Common install locations
    $possible = @(
        "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1",
        "$env:ProgramFiles\WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1",
        "$env:ProgramFiles\PowerShell\Scripts\Get-WindowsAutopilotInfo.ps1"
    )
    foreach ($p in $possible) { if (Test-Path $p) { $script = Get-Item $p; break } }
}
