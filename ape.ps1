#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # Suppress Microsoft Graph SDK welcome/banner noise (where respected)
    $env:MG_SHOW_WELCOME_MESSAGE = 'false'

    # Networking in OOBE
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Allow script execution in this process only
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
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

    # PowerShellGet (best effort) – suppress PackageManagement-in-use noise in OOBE
    try {
        Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop | Out-Null
        Import-Module PowerShellGet -Force -ErrorAction Stop | Out-Null
    }
    catch {
        if ($_.Exception.Message -match "PackageManagement.*currently in use") {
            # Expected in OOBE; continue with in-box modules
        }
        Import-Module PowerShellGet -ErrorAction SilentlyContinue | Out-Null
    }

    # ------------------------------------------------------------
    # PRE-CLEAN: remove any prior Get-WindowsAutopilotInfo.ps1 copies
    # ------------------------------------------------------------
    $pathsToRemove = New-Object System.Collections.Generic.List[string]

    $existingCmd = Get-Command Get-WindowsAutopilotInfo -ErrorAction SilentlyContinue
    if ($existingCmd -and $existingCmd.Source) { [void]$pathsToRemove.Add($existingCmd.Source) }

    $common = @(
        (Join-Path $env:USERPROFILE  'Documents\WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1'),
        (Join-Path $env:ProgramFiles 'WindowsPowerShell\Scripts\Get-WindowsAutopilotInfo.ps1'),
        (Join-Path $env:ProgramFiles 'PowerShell\Scripts\Get-WindowsAutopilotInfo.ps1')
    )
    foreach ($p in $common) { [void]$pathsToRemove.Add($p) }

    foreach ($p in @($pathsToRemove | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $p) {
            try { Remove-Item -LiteralPath $p -Force -ErrorAction Stop } catch { }
        }
    }

    # ------------------------------------------------------------
    # INSTALL: Get-WindowsAutopilotInfo (conditional SkipPublisherCheck)
    # ------------------------------------------------------------
    $installScriptParams = @{
        Name  = 'Get-WindowsAutopilotInfo'
        Force = $true
    }

    $installScriptCmd = Get-Command Install-Script -ErrorAction Stop
    if ($installScriptCmd.Parameters.ContainsKey('SkipPublisherCheck')) {
        $installScriptParams['SkipPublisherCheck'] = $true
    }

    Install-Script @installScriptParams | Out-Null

    # Resolve script path after reinstall
    $cmd = Get-Command Get-WindowsAutopilotInfo -ErrorAction SilentlyContinue
    $scriptPath = $null

    if ($cmd -and $cmd.Source -and (Test-Path -LiteralPath $cmd.Source)) {
        $scriptPath = $cmd.Source
    }
    else {
        $scriptPath = $common | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    }

    if (-not $scriptPath) {
        throw "Get-WindowsAutopilotInfo.ps1 installed, but could not locate it to execute in this session."
    }

    # ------------------------------------------------------------
    # RUN: child process capture (controls console output)
    # ------------------------------------------------------------
    $outFile = Join-Path $env:TEMP ("autopilot_out_{0}.txt" -f ([guid]::NewGuid()))
    $errFile = Join-Path $env:TEMP ("autopilot_err_{0}.txt" -f ([guid]::NewGuid()))

    New-Item -Path $outFile -ItemType File -Force | Out-Null
    New-Item -Path $errFile -ItemType File -Force | Out-Null

    $psExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'

    # Child command: also suppress Graph welcome there
    $childCommand = @"
`$env:MG_SHOW_WELCOME_MESSAGE='false'
& '$scriptPath' -Online
"@

    $p = Start-Process -FilePath $psExe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-Command', $childCommand) `
        -Wait -PassThru -WindowStyle Hidden `
        -RedirectStandardOutput $outFile `
        -RedirectStandardError  $errFile

    $stdout = @((Get-Content -LiteralPath $outFile -ErrorAction SilentlyContinue))
    $stderr = @((Get-Content -LiteralPath $errFile -ErrorAction SilentlyContinue))

    Remove-Item -LiteralPath $outFile, $errFile -Force -ErrorAction SilentlyContinue

    $all = @($stdout + $stderr)
    $allText = ($all -join "`n")

    # ------------------------------------------------------------
    # FILTER: suppress Graph/WAM chatter + benign Group Tag noise
    # ------------------------------------------------------------
    $noisePatterns = @(
        '^Welcome to Microsoft Graph!',
        '^Connected via delegated access',
        '^Readme:',
        '^SDK Docs:',
        '^API Docs:',
        '^NOTE:',
        'Web Account Manager',
        '\bWAM\b',
        'Set-MgGraphOption',
        'clientId'
    )

    $benignPatterns = @(
        'property\s+"Group Tag"\s+cannot be found'
    )

    $filtered = $all
    foreach ($pat in $noisePatterns)  { $filtered = @($filtered | Where-Object { $_ -notmatch $pat }) }
    foreach ($pat in $benignPatterns) { $filtered = @($filtered | Where-Object { $_ -notmatch $pat }) }

    # ------------------------------------------------------------
    # PRINT: only the two progress lines you care about (if present)
    # ------------------------------------------------------------
    $progress = @(
        $filtered | Where-Object { $_ -match '^Connected to Intune tenant\b' }
    ) + @(
        $filtered | Where-Object { $_ -match '^Gathered details for device with serial number:\b' }
    )

    # De-dup while preserving order
    $seen = @{}
    foreach ($line in @($progress)) {
        if (-not [string]::IsNullOrWhiteSpace($line) -and -not $seen.ContainsKey($line)) {
            $seen[$line] = $true
            Write-Output $line
        }
    }

    # ------------------------------------------------------------
    # RESULT: deterministic mapping
    # ------------------------------------------------------------

    # Your known "already there" condition
    $alreadyImported =
        ($allText -match 'ZtdDeviceAlreadyAssigned') -or
        ($allText -match '\b806\b.*ZtdDeviceAlreadyAssigned') -or
        ($allText -match ':\s*806\s+ZtdDeviceAlreadyAssigned')

    # Best-effort "new import succeeded" signals (different versions emit different strings)
    $importSucceeded =
        ($allText -match '(?i)\bimport(ed)?\b.*\bsuccess') -or
        ($allText -match '(?i)\bsuccessfully\b.*\bimport') -or
        ($allText -match '(?i)\bdevice(s)?\b.*\bimported\b') -or
        ($allText -match '(?i)\bupload(ed)?\b.*\b(hash|hardware)\b')

    if ($alreadyImported) {
        Write-Output "Device already imported."
        exit 0
    }

    if ($p.ExitCode -eq 0 -and $importSucceeded) {
        Write-Output "1 device imported successfully."
        exit 0
    }

    if ($p.ExitCode -eq 0) {
        # Completed but couldn't classify from text (rare)
        Write-Output "Completed (verify in Autopilot/Intune if required)."
        exit 0
    }

    # Non-zero exit: surface remaining filtered stderr as errors, otherwise generic fail
    $errFiltered = @($stderr)
    foreach ($pat in $noisePatterns)  { $errFiltered = @($errFiltered | Where-Object { $_ -notmatch $pat }) }
    foreach ($pat in $benignPatterns) { $errFiltered = @($errFiltered | Where-Object { $_ -notmatch $pat }) }

    if (@($errFiltered).Count -gt 0) {
        $errFiltered | ForEach-Object { Write-Error $_ }
    }
    else {
        Write-Error "Get-WindowsAutopilotInfo failed with exit code $($p.ExitCode)."
    }
    exit 1
}
catch {
    Write-Error $_
    exit 1
}
