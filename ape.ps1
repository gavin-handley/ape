#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # Suppress Microsoft Graph SDK welcome/banner noise (if respected by Graph modules)
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

    # PowerShellGet (best effort) – suppress PackageManagement-in-use noise in OOBE
    try {
        Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop | Out-Null
        Import-Module PowerShellGet -Force -ErrorAction Stop | Out-Null
    }
    catch {
        if ($_.Exception.Message -match "PackageManagement.*currently in use") {
            # Expected in OOBE; continue with in-box modules
        }
        else {
            # Keep quiet unless you want to see it
            # Write-Warning "PowerShellGet update/import issue (continuing): $($_.Exception.Message)"
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

    $pathsToRemove = @($pathsToRemove | Select-Object -Unique)

    foreach ($p in $pathsToRemove) {
        if (Test-Path -LiteralPath $p) {
            try { Remove-Item -LiteralPath $p -Force -ErrorAction Stop }
            catch { }
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
    # RUN: child process capture (guarantees we can hide Graph banners)
    # ------------------------------------------------------------
    $outFile = Join-Path $env:TEMP ("autopilot_out_{0}.txt" -f ([guid]::NewGuid().ToString()))
    $errFile = Join-Path $env:TEMP ("autopilot_err_{0}.txt" -f ([guid]::NewGuid().ToString()))

    New-Item -Path $outFile -ItemType File -Force | Out-Null
    New-Item -Path $errFile -ItemType File -Force | Out-Null

    $psExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'

    # In the child process, suppress Graph welcome banner as well
    $childCommand = @"
`$env:MG_SHOW_WELCOME_MESSAGE='false'
& '$scriptPath' -Online
"@

    $argList = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-Command', $childCommand
    )

    $p = Start-Process -FilePath $psExe -ArgumentList $argList -Wait -PassThru `
        -WindowStyle Hidden `
        -RedirectStandardOutput $outFile `
        -RedirectStandardError  $errFile

    $stdout = Get-Content -LiteralPath $outFile -ErrorAction SilentlyContinue
    $stderr = Get-Content -LiteralPath $errFile -ErrorAction SilentlyContinue

    Remove-Item -LiteralPath $outFile, $errFile -Force -ErrorAction SilentlyContinue

    # Always treat these as string arrays (handles $null + single line output)
    $stdout = @($stdout)
    $stderr = @($stderr)

    # ------------------------------------------------------------
    # OUTPUT FILTERING: whitelist only what you want engineers to see
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
        '\bNoWelcome\b',
        'Set-MgGraphOption',
        'clientId'
    )

    $benignPatterns = @(
        'property\s+"Group Tag"\s+cannot be found'
    )

    function Remove-Noise {
        param([string[]]$lines)

        $lines = @($lines)  # force array
        if ($lines.Count -eq 0) { return @() }

        $filtered = $lines
        foreach ($pat in $noisePatterns)  { $filtered = @($filtered | Where-Object { $_ -notmatch $pat }) }
        foreach ($pat in $benignPatterns) { $filtered = @($filtered | Where-Object { $_ -notmatch $pat }) }

        return @($filtered)
    }

    $stdout = Remove-Noise -lines $stdout
    $stderr = Remove-Noise -lines $stderr

    # Whitelist only the two key progress lines
    $keepPatterns = @(
        '^Connected to Intune tenant\b',
        '^Gathered details for device with serial number:\b'
    )

    $kept = @()
    foreach ($pat in $keepPatterns) {
        $kept += @($stdout | Where-Object { $_ -match $pat })
        $kept += @($stderr | Where-Object { $_ -match $pat })
    }

    # De-duplicate while preserving order, and keep as array
    $seen = @{}
    $keptUnique = @()
    foreach ($line in @($kept)) {
        if (-not $seen.ContainsKey($line)) {
            $seen[$line] = $true
            $keptUnique += $line
        }
    }

    if (@($keptUnique).Count -gt 0) {
        $keptUnique | ForEach-Object { Write-Output $_ }
    }

    # ------------------------------------------------------------
    # FINAL STATUS LINE (controlled by us)
    # ------------------------------------------------------------

    $allText = (@($stdout + $stderr) -join "`n")

    $looksAlready  = $allText -match '(already\s+(imported|exists|registered|present))'
    $looksImported = $allText -match '(import(ed)?\s+success|successfully\s+import|uploaded\s+success)'

    if ($p.ExitCode -eq 0) {
        if ($looksImported) {
            Write-Output "Result: Imported successfully."
        }
        elseif ($looksAlready) {
            Write-Output "Result: Already imported / no change."
        }
        else {
            Write-Output "Result: Completed (verify in Intune if required)."
        }
        exit 0
    }

    # Non-zero exit: show remaining stderr (still filtered)
    if (@($stderr).Count -gt 0) {
        $stderr | ForEach-Object { Write-Error $_ }
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
