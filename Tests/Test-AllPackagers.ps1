#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Single-app test harness for Application Packager regression testing.
    Designed to be called per-app by the orchestrator (Invoke-PackagerRegressionTest.ps1).

.DESCRIPTION
    Runs one packager script through stage/install/detect/process/uninstall/verify-removal.
    Returns a structured result object. Does NOT handle reboots -- the orchestrator does that.

    Exit codes from install/uninstall (0, 3010, other) are captured in the result so the
    orchestrator can decide whether to reboot before continuing.

.PARAMETER PackagerName
    Base name of the packager (e.g., "package-7zip").

.PARAMETER Phase
    Which phase(s) to run. Default: All.
    - Stage        : Download and create manifest only
    - Install      : Run install.bat (requires prior Stage)
    - Detect       : Run detection check (requires prior Install)
    - Uninstall    : Run uninstall.bat (requires prior Install)
    - VerifyRemoval: Check detection is gone (requires prior Uninstall)
    - All          : Run full cycle

.PARAMETER DownloadRoot
    Local staging root. Default: C:\temp\ap

.PARAMETER Cleanup
    Delete staged content folder after test completes.

.PARAMETER TimeoutSec
    Per-phase timeout in seconds. Default: 600 (10 min).
#>
param(
    [Parameter(Mandatory)]
    [string]$PackagerName,

    [ValidateSet('Stage','Install','Detect','Uninstall','VerifyRemoval','All')]
    [string]$Phase = 'All',

    [string]$DownloadRoot = 'C:\temp\ap',

    [switch]$Cleanup,

    [int]$TimeoutSec = 600
)

$ErrorActionPreference = 'Continue'
$packagerDir = 'c:\projects\applicationpackager\Packagers'

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

# --- Initialize result object ---
$result = [PSCustomObject]@{
    Name              = $PackagerName
    StageResult       = 'SKIP'
    Version           = ''
    ContentPath       = ''
    InstallResult     = 'SKIP'
    InstallExitCode   = -1
    DetectionResult   = 'SKIP'
    ProcessVerified   = 'SKIP'
    ActualProcess     = ''
    ManifestProcess   = ''
    UninstallResult   = 'SKIP'
    UninstallExitCode = -1
    RemovalVerified   = 'SKIP'
    Error             = ''
    NeedsReboot       = $false
    RebootPhase       = ''
}

$manifest = $null
$contentPath = ''

# --- Resolve content path from prior stage if not running Stage phase ---
if ($Phase -ne 'Stage' -and $Phase -ne 'All') {
    # Find content path from packager's subfolder
    $scriptPath = Join-Path $packagerDir "$PackagerName.ps1"
    if (Test-Path $scriptPath) {
        # Read the script to find the app subfolder name (BaseDownloadRoot pattern)
        $scriptContent = Get-Content $scriptPath -Raw
        if ($scriptContent -match '\$BaseDownloadRoot\s*=\s*Join-Path\s+\$DownloadRoot\s+"([^"]+)"') {
            $appFolder = $Matches[1]
            $appRoot = Join-Path $DownloadRoot $appFolder
            if (Test-Path $appRoot) {
                $versionDirs = Get-ChildItem $appRoot -Directory -ErrorAction SilentlyContinue |
                    Where-Object { Test-Path (Join-Path $_.FullName 'stage-manifest.json') } |
                    Sort-Object LastWriteTime -Descending
                if ($versionDirs) {
                    $contentPath = $versionDirs[0].FullName
                    $manifestFile = Join-Path $contentPath 'stage-manifest.json'
                    $manifest = Get-Content $manifestFile -Raw | ConvertFrom-Json
                    $result.ContentPath = $contentPath
                    $result.Version = $manifest.SoftwareVersion
                }
            }
        }
    }
}

# ─── STAGE ───────────────────────────────────────────────────────────────────

if ($Phase -eq 'Stage' -or $Phase -eq 'All') {
    Write-TestLog "  Staging $PackagerName..." -Level INFO
    try {
        $stageOutput = powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "
            cd '$packagerDir'
            . '.\$PackagerName.ps1' -StageOnly -DownloadRoot '$DownloadRoot' 2>&1
        " 2>&1

        $contentPath = ($stageOutput | Where-Object { $_ -match '^C:\\' } | Select-Object -Last 1)

        if ($contentPath -and (Test-Path $contentPath)) {
            $result.StageResult = 'PASS'
            $result.ContentPath = "$contentPath"

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

    Write-TestLog "  Stage: $($result.StageResult) $(if ($result.Version) { "v$($result.Version)" })" -Level $result.StageResult

    if ($Phase -eq 'Stage' -or $result.StageResult -ne 'PASS') {
        return $result
    }
}

# ─── INSTALL ─────────────────────────────────────────────────────────────────

if ($Phase -eq 'Install' -or $Phase -eq 'All') {
    $installBat = Join-Path $contentPath 'install.bat'
    if (Test-Path $installBat) {
        try {
            Write-TestLog "  Installing..." -Level INFO
            $proc = Start-Process cmd.exe -ArgumentList @('/c', $installBat) `
                -Wait -PassThru -NoNewWindow -WorkingDirectory $contentPath
            $result.InstallExitCode = $proc.ExitCode

            if ($proc.ExitCode -eq 0) {
                $result.InstallResult = 'PASS'
            } elseif ($proc.ExitCode -eq 3010) {
                $result.InstallResult = 'PASS'
                $result.NeedsReboot = $true
                $result.RebootPhase = 'Install'
                Write-TestLog "  Install returned 3010 (reboot required)" -Level WARN
                return $result
            } else {
                $result.InstallResult = 'FAIL'
                $result.Error = "Install exit code: $($proc.ExitCode)"
            }
        } catch {
            $result.InstallResult = 'FAIL'
            $result.Error = $_.Exception.Message
        }
    } else {
        $result.InstallResult = 'FAIL'
        $result.Error = 'install.bat not found'
    }

    if ($result.InstallResult -ne 'PASS') {
        return $result
    }
}

# ─── DETECT ──────────────────────────────────────────────────────────────────

if ($Phase -eq 'Detect' -or $Phase -eq 'All') {
    if ($manifest -and $manifest.Detection) {
        $det = $manifest.Detection

        # x86 apps register in WOW6432Node on 64-bit OS
        $is64Bit = if ($null -ne $det.Is64Bit) { $det.Is64Bit } else { $true }

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
                $relKey = $det.RegistryKeyRelative
                if (-not $is64Bit -and $relKey -match '^SOFTWARE\\' -and $relKey -notmatch 'WOW6432Node') {
                    $relKey = $relKey -replace '^SOFTWARE\\', 'SOFTWARE\WOW6432Node\'
                }
                $regPath = "HKLM:\$relKey"
                $regVal = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
                if ($regVal) {
                    $result.DetectionResult = 'PASS'
                } else {
                    $result.DetectionResult = 'FAIL'
                    $result.Error = "Detection registry key not found: $regPath"
                }
            }
            'RegistryKey' {
                $relKey = $det.RegistryKeyRelative
                if (-not $is64Bit -and $relKey -match '^SOFTWARE\\' -and $relKey -notmatch 'WOW6432Node') {
                    $relKey = $relKey -replace '^SOFTWARE\\', 'SOFTWARE\WOW6432Node\'
                }
                $regPath = "HKLM:\$relKey"
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
        Write-TestLog "  Detect: $($result.DetectionResult)" -Level $result.DetectionResult
    } else {
        $result.DetectionResult = 'SKIP'
        Write-TestLog "  Detect: SKIP (no detection in manifest)" -Level SKIP
    }

    # --- PROCESS NAME VERIFICATION ---
    if ($manifest -and $manifest.RunningProcess) {
        $procs = @($manifest.RunningProcess)
        if ($procs.Count -gt 0 -and $procs[0] -ne '') {
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

    Write-TestLog "  Install: $($result.InstallResult) | Detect: $($result.DetectionResult) | Process: $($result.ProcessVerified)" -Level $(
        if ($result.InstallResult -eq 'PASS' -and $result.DetectionResult -eq 'PASS') { 'PASS' } else { 'WARN' }
    )

    if ($Phase -eq 'Detect') {
        return $result
    }
}

# ─── UNINSTALL ───────────────────────────────────────────────────────────────

if ($Phase -eq 'Uninstall' -or $Phase -eq 'All') {
    $uninstallBat = Join-Path $contentPath 'uninstall.bat'
    if (Test-Path $uninstallBat) {
        try {
            Write-TestLog "  Uninstalling..." -Level INFO
            $proc = Start-Process cmd.exe -ArgumentList @('/c', $uninstallBat) `
                -Wait -PassThru -NoNewWindow -WorkingDirectory $contentPath
            $result.UninstallExitCode = $proc.ExitCode

            if ($proc.ExitCode -eq 0) {
                $result.UninstallResult = 'PASS'
            } elseif ($proc.ExitCode -eq 3010) {
                $result.UninstallResult = 'PASS'
                $result.NeedsReboot = $true
                $result.RebootPhase = 'Uninstall'
                Write-TestLog "  Uninstall returned 3010 (reboot required)" -Level WARN
                return $result
            } else {
                $result.UninstallResult = 'FAIL'
                $result.Error += " | Uninstall exit code: $($proc.ExitCode)"
            }
        } catch {
            $result.UninstallResult = 'FAIL'
            $result.Error += " | $($_.Exception.Message)"
        }
    } else {
        $result.UninstallResult = 'SKIP'
        Write-TestLog "  Uninstall: SKIP (no uninstall.bat)" -Level SKIP
    }

    if ($result.UninstallResult -ne 'PASS') {
        return $result
    }
}

# ─── VERIFY REMOVAL ──────────────────────────────────────────────────────────

if ($Phase -eq 'VerifyRemoval' -or $Phase -eq 'All') {
    if ($manifest -and $manifest.Detection) {
        Start-Sleep -Seconds 3
        $remIs64 = if ($null -ne $manifest.Detection.Is64Bit) { $manifest.Detection.Is64Bit } else { $true }
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
                $relKey = $manifest.Detection.RegistryKeyRelative
                if (-not $remIs64 -and $relKey -match '^SOFTWARE\\' -and $relKey -notmatch 'WOW6432Node') {
                    $relKey = $relKey -replace '^SOFTWARE\\', 'SOFTWARE\WOW6432Node\'
                }
                $regPath = "HKLM:\$relKey"
                if (-not (Test-Path $regPath)) {
                    $result.RemovalVerified = 'PASS'
                } else {
                    $result.RemovalVerified = 'WARN'
                }
            }
            'RegistryKey' {
                $relKey = $manifest.Detection.RegistryKeyRelative
                if (-not $remIs64 -and $relKey -match '^SOFTWARE\\' -and $relKey -notmatch 'WOW6432Node') {
                    $relKey = $relKey -replace '^SOFTWARE\\', 'SOFTWARE\WOW6432Node\'
                }
                $regPath = "HKLM:\$relKey"
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
    } else {
        $result.RemovalVerified = 'SKIP'
    }

    Write-TestLog "  Uninstall: $($result.UninstallResult) | Removal: $($result.RemovalVerified)" -Level $(
        if ($result.UninstallResult -eq 'PASS' -and $result.RemovalVerified -ne 'FAIL') { 'PASS' } else { 'WARN' }
    )
}

# --- Return result ---
return $result
