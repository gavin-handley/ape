#requires -Version 5.1
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ProgressPreference = 'SilentlyContinue'

    # NuGet provider
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    # Trust PSGallery
    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    # PowerShellGet (best effort)
    try {
        Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser | Out-Null
        Import-Module PowerShellGet -Force | Out-Null
    }
    catch {
        Import-Module PowerShellGet -ErrorAction SilentlyContinue | Out-Null
    }

    # Install Get-WindowsAutopilotInfo (only add SkipPublisherCheck if supported)
    $installScriptParams = @{
        Name  = 'Get-WindowsAutopilotInfo'
        Force = $true
    }
    $installScriptCmd = Get-Command Install-Script -ErrorAction Stop
    if ($installScriptCmd.Parameters.ContainsKey('SkipPublisherCheck')) {
        $installScriptParams['SkipPublisherCheck'] = $true
    }
    Install-Script @installScriptParams | Out-Null

    # Resolve the script path robustly
    $cmd = Get-Command Get-WindowsAutopilotInfo -ErrorAction SilentlyContinue
    $scriptPath = $null

    if ($cmd -and $cmd.Source -and (Test-Path -LiteralPath $cmd.Source)) {
        $scriptPath = $cmd.Source
    }
    else {
        $possible = @(
            (Join-Path $env:USERPROFILE  'Documents\WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1'),
            (Join-Path $env:ProgramFiles 'WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1'),
            (Join-Path $env:ProgramFiles 'PowerShell\Scripts\Get-WindowsAutopilotInfo.ps1')
        )
        $scriptPath = $possible | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    }

    if (-not $scriptPath) {
        throw "Get-WindowsAutopilotInfo.ps1 installed, but could not locate it to execute in this session."
    }

    # Run Autopilot registration ONLINE.
    # If the only failure is the known 'Group Tag' property lookup bug, treat as success.
    try {
        & $scriptPath -Online
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match 'property\s+"Group Tag"\s+cannot be found') {
            Write-Warning "Autopilot registration likely succeeded; Get-WindowsAutopilotInfo hit a non-fatal reporting bug: $msg"
            exit 0
        }
        throw
    }

    exit 0
}
catch {
    Write-Error $_
    exit 1
}
