#requires -Version 5.1
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ProgressPreference = 'SilentlyContinue'

    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    try {
        Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser | Out-Null
        Import-Module PowerShellGet -Force | Out-Null
    }
    catch {
        Import-Module PowerShellGet -ErrorAction SilentlyContinue | Out-Null
    }

    # Install Get-WindowsAutopilotInfo with conditional SkipPublisherCheck support
    $installScriptParams = @{
        Name  = 'Get-WindowsAutopilotInfo'
        Force = $true
    }
    $installScriptCmd = Get-Command Install-Script -ErrorAction Stop
    if ($installScriptCmd.Parameters.ContainsKey('SkipPublisherCheck')) {
        $installScriptParams['SkipPublisherCheck'] = $true
    }
    Install-Script @installScriptParams | Out-Null

    # Resolve path robustly
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

    # Run -Online, capture all output and filter known benign 'Group Tag' error
    $results = & $scriptPath -Online 2>&1

    $errorRecords = @($results | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
    $normalOutput = @($results | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] })

    if ($normalOutput.Count -gt 0) {
        $normalOutput | ForEach-Object { Write-Output $_ }
    }

    if ($errorRecords.Count -gt 0) {
        # Allowlist the specific noisy message
        $benign = $errorRecords | Where-Object { $_.Exception.Message -match 'property\s+"Group Tag"\s+cannot be found' }
        $real   = $errorRecords | Where-Object { $_.Exception.Message -notmatch 'property\s+"Group Tag"\s+cannot be found' }

        if ($real.Count -eq 0) {
            Write-Warning "Autopilot action completed; suppressed a known non-fatal output issue in Get-WindowsAutopilotInfo (Group Tag property)."
            exit 0
        }

        # Surface real errors
        $real | ForEach-Object { Write-Error $_ }
        exit 1
    }

    exit 0
}
catch {
    Write-Error $_
    exit 1
}
