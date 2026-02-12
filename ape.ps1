#requires -Version 5.1
<#
.SYNOPSIS
  OOBE-friendly bootstrap for Get-WindowsAutopilotInfo -Online on Windows PowerShell 5.1.

.NOTES
  - Intended to be executed via: powershell -NoProfile -ep Bypass -c "iex(iwr -useb '<url>')"
  - Minimises prompts, sets TLS 1.2, installs prerequisites and runs autopilot registration online.
#>

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

try {
    # --- Networking / security defaults for OOBE ---
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Avoid any policy blocks in this session only
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null

    # Reduce noise, keep failures explicit
    $ProgressPreference = 'SilentlyContinue'

    # --- Ensure NuGet provider (needed for PSGallery installs) ---
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    # --- Trust PSGallery (avoid prompts) ---
    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    # --- Ensure PowerShellGet is available (Install-Script reliability) ---
    # On some builds, PSGet is present but old; forcing update reduces "Install-Script" oddities.
    try {
        Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser | Out-Null
        Import-Module PowerShellGet -Force
    }
    catch {
        # If update fails (e.g. network restrictions), continue and try with existing PSGet.
        Import-Module PowerShellGet -ErrorAction SilentlyContinue | Out-Null
    }

    # --- Install the Autopilot script ---
    Install-Script -Name Get-WindowsAutopilotInfo -Force -SkipPublisherCheck | Out-Null

    # --- Run it: prefer command resolution, fall back to known install paths ---
    $cmd = Get-Command Get-WindowsAutopilotInfo -ErrorAction SilentlyContinue

    if ($cmd -and $cmd.Source -and (Test-Path -LiteralPath $cmd.Source)) {
        & $cmd.Source -Online
        exit 0
    }

    # Fallback paths (OOBE can have odd PATH/profile behaviours)
    $possible = @(
        (Join-Path $env:USERPROFILE   'Documents\WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1'),
        (Join-Path $env:ProgramFiles  'WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1'),
        (Join-Path $env:ProgramFiles  'PowerShell\Scripts\Get-WindowsAutopilotInfo.ps1')
    )

    $path = $possible | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $path) {
        throw "Get-WindowsAutopilotInfo.ps1 installed, but script path wasn't discoverable (PATH/profile limitations in OOBE)."
    }

    & $path -Online
    exit 0
}
catch {
    Write-Error $_
    exit 1
}
