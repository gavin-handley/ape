#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # Suppress Microsoft Graph SDK welcome/banner noise (if Graph modules respect it)
    $env:MG_SHOW_WELCOME_MESSAGE = 'false'

    # Networking in OOBE
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Allow script execution in this process only
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null

    # Reduce noisy progress output
    $ProgressPreference = 'SilentlyContinue'

    # Ensure NuGet provider exists
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    # Trust PSGallery (avoid prompts)
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

    # --- Run in a separate PowerShell process to guarantee we can suppress specific noisy output ---
    $outFile = Join-Path $env:TEMP ("autopilot_out_{0}.txt" -f ([guid]::NewGuid().ToString()))
    $errFile = Join-Path $env:TEMP ("autopilot_err_{0}.txt" -f ([guid]::NewGuid().ToString()))

    # Ensure files exist so Get-Content won't fail later
    New-Item -Path $outFile -ItemType File -Force | Out-Null
    New-Item -Path $errFile -ItemType File -Force | Out-Null

    # Run Get-WindowsAutopilotInfo -Online in a child process; capture stdout+stderr
    $psExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $argList = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-Command',
        "& '$scriptPath' -Online"
    )

    $p = Start-Process -FilePath $psExe -ArgumentList $argList -Wait -PassThru `
        -WindowStyle Hidden `
        -RedirectStandardOutput $outFile `
        -RedirectStandardError  $errFile

    $stdout = Get-Content -LiteralPath $outFile -ErrorAction SilentlyContinue
    $stderr = Get-Content -LiteralPath $errFile -ErrorAction SilentlyContinue

    # Clean up temp files
    Remove-Item -LiteralPath $outFile, $errFile -Force -ErrorAction SilentlyContinue

    # Filter known benign "Group Tag" message (it may appear in stderr OR stdout)
    $filterRegex = 'property\s+"Group Tag"\s+cannot be found'
    $stderrFiltered = @($stderr | Where-Object { $_ -notmatch $filterRegex })
    $stdoutFiltered = @($stdout | Where-Object { $_ -notmatch $filterRegex })

    # Emit normal output (you can comment this out if you want near-silent success)
    if ($stdoutFiltered.Count -gt 0) {
        $stdoutFiltered | ForEach-Object { Write-Output $_ }
    }

    # If the child process returned non-zero, decide if it's a real failure
    if ($p.ExitCode -ne 0) {
        # If the only stderr was the benign Group Tag message, treat as success
        if ($stderrFiltered.Count -eq 0 -and ($stderr.Count -gt 0)) {
            Write-Warning "Autopilot action completed; suppressed a known non-fatal output issue in Get-WindowsAutopilotInfo (Group Tag property)."
            exit 0
        }

        # Otherwise, surface errors
        if ($stderrFiltered.Count -gt 0) {
            $stderrFiltered | ForEach-Object { Write-Error $_ }
        }
        else {
            Write-Error "Get-WindowsAutopilotInfo failed with exit code $($p.ExitCode)."
        }

        exit 1
    }

    # Child succeeded; if it produced only the benign message, keep it clean
    if ($stderrFiltered.Count -gt 0) {
        # Some warnings/errors may still be meaningful; print as warnings rather than hard fail
        $stderrFiltered | ForEach-Object { Write-Warning $_ }
    }

    exit 0
}
catch {
    Write-Error $_
    exit 1
}
