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

# Try to resolve the command first (most reliable)
$cmd = Get-Command Get-WindowsAutopilotInfo -ErrorAction SilentlyContinue

if ($cmd) {
    & $cmd.Source -Online
}
else {
    # Fallback: search common install paths for Install-Script
    $possible = @(
        Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1',
        Join-Path ${env:ProgramFiles} 'WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1',
        Join-Path ${env:ProgramFiles} 'PowerShell\Scripts\Get-WindowsAutopilotInfo.ps1'
    )

    $path = $possible | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $path) {
        throw "Get-WindowsAutopilotInfo.ps1 not found after installation. Check Install-Script output and `$env:PSModulePath / PATH."
    }

    & $path -Online
}
