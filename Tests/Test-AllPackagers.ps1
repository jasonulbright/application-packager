#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Automated test harness: stages, installs, verifies detection + process names,
    uninstalls, and verifies removal for all Application Packager scripts.

.PARAMETER IncludeOnly
    Array of packager base names to test (e.g., "package-notepadplusplus").
    If omitted, tests all packagers.

.PARAMETER SkipList
    Array of packager base names to skip.

.PARAMETER StageOnly
    Only stage (download) -- do not install or uninstall.

.PARAMETER DownloadRoot
    Local staging root. Default: C:\temp\ap
#>
param(
    [string[]]$IncludeOnly,
    [string[]]$SkipList = @(
        'package-vmwareworkstation',   # User excluded
        'package-m365apps-x64',       # Requires O365 license
        'package-m365apps-x86',
        'package-m365project-x64',
        'package-m365project-x86',
        'package-m365visio-x64',
        'package-m365visio-x86',
        'package-vs2026',             # Multi-GB, hours to install
        'package-vs2026community',
        'package-citrixworkspacecr',  # Requires Citrix infra
        'package-citrixworkspaceltsr',
        'package-xencenter',          # Requires XenServer
        'package-xenservervmtools',
        'package-tableaudesktop',     # Requires license
        'package-tableauprep',
        'package-tableaureader',
        'package-ssms',               # 1GB+, long install
        'package-powerbidesktop',     # 500MB+, long install
        'package-anaconda',           # 1GB+
        'package-pycharm'             # 800MB+
    ),
    [switch]$StageOnly,
    [string]$DownloadRoot = 'C:\temp\ap'
)

$ErrorActionPreference = 'Continue'
$packagerDir = 'c:\projects\applicationpackager\Packagers'

# Results tracking
$results = [System.Collections.ArrayList]::new()

function Write-TestLog {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'HH:mm:ss'
    $color = switch ($Level) {
        'PASS' { 'Green' }
        'FAIL' { 'Red' }
        'WARN' { 'Yellow' }
        'SKIP' { 'DarkGray' }
        default { 'White' }
    }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color
}

# Get packager list
$allPackagers = Get-ChildItem "$packagerDir\package-*.ps1" | Where-Object {
    $_.BaseName -notin @('package-7zip') # 7zip already validated
} | Sort-Object Name

if ($IncludeOnly) {
    $allPackagers = $allPackagers | Where-Object { $_.BaseName -in $IncludeOnly }
}
$allPackagers = $allPackagers | Where-Object { $_.BaseName -notin $SkipList }

Write-TestLog "Testing $($allPackagers.Count) packagers" -Level INFO
Write-TestLog "DownloadRoot: $DownloadRoot" -Level INFO
Write-TestLog ""

foreach ($packager in $allPackagers) {
    $name = $packager.BaseName
    Write-TestLog "=== $name ===" -Level INFO

    $result = [PSCustomObject]@{
        Name              = $name
        StageResult       = 'SKIP'
        Version           = ''
        InstallResult     = 'SKIP'
        DetectionResult   = 'SKIP'
        ProcessVerified   = 'SKIP'
        ActualProcess     = ''
        ManifestProcess   = ''
        UninstallResult   = 'SKIP'
        RemovalVerified   = 'SKIP'
        Error             = ''
    }

    # --- STAGE ---
    try {
        $stageOutput = powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "
            cd '$packagerDir'
            . '.\$name.ps1' -StageOnly -DownloadRoot '$DownloadRoot' 2>&1
        " 2>&1

        # Find the staged content path (last line of output that's a path)
        $contentPath = ($stageOutput | Where-Object { $_ -match '^C:\\' } | Select-Object -Last 1)

        if ($contentPath -and (Test-Path $contentPath)) {
            $result.StageResult = 'PASS'

            # Read manifest
            $manifestFile = Join-Path $contentPath 'stage-manifest.json'
            if (Test-Path $manifestFile) {
                $manifest = Get-Content $manifestFile -Raw | ConvertFrom-Json
                $result.Version = $manifest.SoftwareVersion
                $result.ManifestProcess = ($manifest.RunningProcess -join ', ')
            }
        } else {
            $result.StageResult = 'FAIL'
            $result.Error = 'No content path returned from stage'
        }
    } catch {
        $result.StageResult = 'FAIL'
        $result.Error = $_.Exception.Message
    }

    if ($StageOnly -or $result.StageResult -ne 'PASS') {
        $null = $results.Add($result)
        Write-TestLog "  Stage: $($result.StageResult) $(if ($result.Version) { "v$($result.Version)" })" -Level $result.StageResult
        continue
    }

    # --- INSTALL ---
    $installBat = Join-Path $contentPath 'install.bat'
    if (Test-Path $installBat) {
        try {
            Write-TestLog "  Installing..." -Level INFO
            $proc = Start-Process cmd.exe -ArgumentList @('/c', $installBat) -Wait -PassThru -NoNewWindow -WorkingDirectory $contentPath
            if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
                $result.InstallResult = 'PASS'
            } else {
                $result.InstallResult = 'FAIL'
                $result.Error = "Install exit code: $($proc.ExitCode)"
            }
        } catch {
            $result.InstallResult = 'FAIL'
            $result.Error = $_.Exception.Message
        }
    }

    # --- DETECTION ---
    if ($result.InstallResult -eq 'PASS' -and $manifest.Detection) {
        Start-Sleep -Seconds 2
        $det = $manifest.Detection
        switch ($det.Type) {
            'File' {
                $detPath = Join-Path $det.FilePath $det.FileName
                if (Test-Path $detPath) {
                    $result.DetectionResult = 'PASS'
                } else {
                    $result.DetectionResult = 'FAIL'
                    $result.Error = "Detection file not found: $detPath"
                }
            }
            'RegistryKeyValue' {
                $regPath = "HKLM:\$($det.RegistryKeyRelative)"
                $regVal = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
                if ($regVal) {
                    $result.DetectionResult = 'PASS'
                } else {
                    $result.DetectionResult = 'FAIL'
                    $result.Error = "Detection registry key not found: $regPath"
                }
            }
            'RegistryKey' {
                $regPath = "HKLM:\$($det.RegistryKeyRelative)"
                if (Test-Path $regPath) {
                    $result.DetectionResult = 'PASS'
                } else {
                    $result.DetectionResult = 'FAIL'
                    $result.Error = "Detection registry key not found: $regPath"
                }
            }
            'Script' {
                $scriptResult = powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $det.ScriptText 2>&1
                if ($scriptResult -match 'Installed') {
                    $result.DetectionResult = 'PASS'
                } else {
                    $result.DetectionResult = 'FAIL'
                    $result.Error = "Detection script returned: $scriptResult"
                }
            }
            default {
                $result.DetectionResult = 'SKIP'
            }
        }
    }

    # --- PROCESS NAME VERIFICATION ---
    if ($result.InstallResult -eq 'PASS' -and $manifest.RunningProcess) {
        $procs = @($manifest.RunningProcess)
        if ($procs.Count -gt 0 -and $procs[0] -ne '') {
            # Check if the EXE actually exists on disk matching the process name
            $verified = @()
            foreach ($p in $procs) {
                $found = Get-ChildItem "C:\Program Files\*\$p.exe" -Recurse -ErrorAction SilentlyContinue |
                    Select-Object -First 1
                if (-not $found) {
                    $found = Get-ChildItem "C:\Program Files (x86)\*\$p.exe" -Recurse -ErrorAction SilentlyContinue |
                        Select-Object -First 1
                }
                if (-not $found) {
                    $found = Get-ChildItem "C:\ProgramData\*\$p.exe" -Recurse -ErrorAction SilentlyContinue |
                        Select-Object -First 1
                }
                if ($found) {
                    $verified += "$p (OK: $($found.FullName))"
                } else {
                    $verified += "$p (NOT FOUND)"
                }
            }
            $result.ActualProcess = $verified -join '; '
            if ($verified -match 'NOT FOUND') {
                $result.ProcessVerified = 'FAIL'
            } else {
                $result.ProcessVerified = 'PASS'
            }
        } else {
            $result.ProcessVerified = 'SKIP'
        }
    }

    Write-TestLog "  Stage: $($result.StageResult) | Install: $($result.InstallResult) | Detect: $($result.DetectionResult) | Process: $($result.ProcessVerified)" -Level $(if ($result.InstallResult -eq 'PASS' -and $result.DetectionResult -eq 'PASS') { 'PASS' } else { 'WARN' })

    # --- UNINSTALL ---
    if ($result.InstallResult -eq 'PASS') {
        $uninstallBat = Join-Path $contentPath 'uninstall.bat'
        if (Test-Path $uninstallBat) {
            try {
                Write-TestLog "  Uninstalling..." -Level INFO
                $proc = Start-Process cmd.exe -ArgumentList @('/c', $uninstallBat) -Wait -PassThru -NoNewWindow -WorkingDirectory $contentPath
                if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
                    $result.UninstallResult = 'PASS'
                } else {
                    $result.UninstallResult = 'FAIL'
                    $result.Error += " | Uninstall exit code: $($proc.ExitCode)"
                }
            } catch {
                $result.UninstallResult = 'FAIL'
                $result.Error += " | $($_.Exception.Message)"
            }
        }

        # --- VERIFY REMOVAL ---
        if ($result.UninstallResult -eq 'PASS') {
            Start-Sleep -Seconds 3
            switch ($manifest.Detection.Type) {
                'File' {
                    $detPath = Join-Path $manifest.Detection.FilePath $manifest.Detection.FileName
                    if (-not (Test-Path $detPath)) {
                        $result.RemovalVerified = 'PASS'
                    } else {
                        $result.RemovalVerified = 'FAIL'
                        $result.Error += ' | Detection file still exists after uninstall'
                    }
                }
                'RegistryKeyValue' {
                    $regPath = "HKLM:\$($manifest.Detection.RegistryKeyRelative)"
                    if (-not (Test-Path $regPath)) {
                        $result.RemovalVerified = 'PASS'
                    } else {
                        $result.RemovalVerified = 'WARN'
                    }
                }
                'Script' {
                    $scriptResult = powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $manifest.Detection.ScriptText 2>&1
                    if (-not ($scriptResult -match 'Installed')) {
                        $result.RemovalVerified = 'PASS'
                    } else {
                        $result.RemovalVerified = 'FAIL'
                    }
                }
            }
        }

        Write-TestLog "  Uninstall: $($result.UninstallResult) | Removal: $($result.RemovalVerified)" -Level $(if ($result.UninstallResult -eq 'PASS') { 'PASS' } else { 'WARN' })
    }

    $null = $results.Add($result)
}

# --- SUMMARY ---
Write-Host ""
Write-Host "=" * 80
Write-Host "TEST SUMMARY"
Write-Host "=" * 80

$results | Format-Table Name, StageResult, Version, InstallResult, DetectionResult, ProcessVerified, UninstallResult, RemovalVerified -AutoSize

# Export detailed results
$results | ConvertTo-Json -Depth 4 | Set-Content 'c:\temp\packager-test-results.json' -Encoding UTF8

$passed = ($results | Where-Object { $_.InstallResult -eq 'PASS' -and $_.DetectionResult -eq 'PASS' -and $_.UninstallResult -eq 'PASS' }).Count
$failed = ($results | Where-Object { $_.InstallResult -eq 'FAIL' -or $_.DetectionResult -eq 'FAIL' }).Count
$skipped = ($results | Where-Object { $_.StageResult -eq 'SKIP' -or $_.StageResult -eq 'FAIL' }).Count
$processIssues = ($results | Where-Object { $_.ProcessVerified -eq 'FAIL' })

Write-Host ""
Write-Host "Passed: $passed | Failed: $failed | Skipped: $skipped | Process issues: $($processIssues.Count)"

if ($processIssues.Count -gt 0) {
    Write-Host ""
    Write-Host "Process name mismatches:"
    $processIssues | ForEach-Object {
        Write-Host "  $($_.Name): $($_.ActualProcess)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Detailed results: c:\temp\packager-test-results.json"
