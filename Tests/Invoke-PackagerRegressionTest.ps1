#Requires -RunAsAdministrator
#Requires -Modules Hyper-V
<#
.SYNOPSIS
    Remote orchestrator for Application Packager regression testing on CLIENT01.

.DESCRIPTION
    Runs from the Hyper-V host. Optionally reverts CLIENT01 to a clean snapshot,
    deploys the current packager scripts via WinRM, and executes Test-AllPackagers.ps1
    per-app on the remote VM.

    Handles reboots: if install or uninstall returns 3010, the orchestrator reboots
    CLIENT01, waits for WinRM, then continues (detect after install-3010, verify
    removal after uninstall-3010).

    Cleans up staged content on the remote VM after each successful test.

.PARAMETER ResetClient
    Revert CLIENT01 to the Deployment-Complete snapshot before testing.

.PARAMETER StageOnly
    Only test staging (download + manifest). No install/uninstall.

.PARAMETER IncludeOnly
    Test specific packagers only. Array of base names (e.g., "package-7zip").

.PARAMETER SkipList
    Override the default skip list.

.PARAMETER ResultsPath
    Local host directory for collected results. Default: C:\temp\regression-results

.PARAMETER VMName
    Hyper-V VM name. Default: CLIENT01

.PARAMETER CheckpointName
    Snapshot name to restore when using -ResetClient. Default: Deployment-Complete

.PARAMETER Credential
    PSCredential for WinRM to the VM.

.PARAMETER AppTimeoutSec
    Per-app phase timeout in seconds. Default: 600 (10 min).

.PARAMETER WhatIf
    Show what would happen without executing.

.EXAMPLE
    .\Invoke-PackagerRegressionTest.ps1 -ResetClient
    Full regression: reset, deploy, run all apps with reboot handling.

.EXAMPLE
    .\Invoke-PackagerRegressionTest.ps1 -ResetClient -StageOnly
    Fast smoke: reset, stage-only for all apps.

.EXAMPLE
    .\Invoke-PackagerRegressionTest.ps1 -IncludeOnly @('package-7zip','package-vlc')
    Quick test of two specific apps (no reset).
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$ResetClient,
    [switch]$StageOnly,
    [string[]]$IncludeOnly,
    [string[]]$SkipList,
    [string]$ResultsPath = 'C:\temp\regression-results',
    [string]$VMName = 'CLIENT01',
    [string]$CheckpointName = 'Deployment-Complete',
    [PSCredential]$Credential,
    [int]$AppTimeoutSec = 600
)

$ErrorActionPreference = 'Stop'

$localRepoRoot = 'c:\projects\applicationpackager'
$remoteRepoRoot = 'C:\projects\applicationpackager'
$winrmTimeoutSec = 300
$winrmRetrySec = 15

# ─── DEFAULT SKIP LIST ───────────────────────────────────────────────────────

$defaultSkipList = @(
    # --- Per-user installs (won't work headless/SYSTEM) ---
    'package-obsidian'              # NSIS per-user, InstallContext=User
    'package-postman'               # Squirrel per-user, InstallContext=User
    'package-brave'                 # Squirrel, hangs in headless install

    # --- Requires external infrastructure or licenses ---
    'package-vmwareworkstation'     # Requires VMware license
    'package-m365apps-x64'         # Requires O365 license
    'package-m365apps-x86'
    'package-m365project-x64'
    'package-m365project-x86'
    'package-m365visio-x64'
    'package-m365visio-x86'
    'package-citrixworkspacecr'    # Requires Citrix infra
    'package-citrixworkspaceltsr'
    'package-xencenter'            # Requires XenServer
    'package-xenservervmtools'
    'package-tableaudesktop'       # Requires license
    'package-tableauprep'
    'package-tableaureader'

    # --- Very large installs (multi-GB, hours) ---
    'package-vs2026'               # Multi-GB
    'package-vs2026community'
    'package-ssms'                 # 1GB+
    'package-powerbidesktop'       # 500MB+
    'package-anaconda'             # 1GB+
    'package-pycharm'              # 800MB+

    # --- Upstream broken ---
    'package-slack'                # CDN 404 as of 2026-04-10
)

if (-not $PSBoundParameters.ContainsKey('SkipList')) {
    $SkipList = $defaultSkipList
}

# ─── NO-UNINSTALL LIST ──────────────────────────────────────────────────────
# System dependencies that other apps rely on. Install + detect only.

$noUninstallList = @(
    'package-msvcruntimes'          # VC++ Redistributable - everything depends on this
    'package-dotnet8'               # .NET 8 runtime
    'package-Dotnet9x64'            # .NET 9 runtime
    'package-Dotnet10x64'           # .NET 10 runtime
    'package-aspnethostingbundle8'  # ASP.NET runtime
    'package-edge'                  # OS-integrated browser
    'package-webview2'              # Edge WebView2 runtime - many apps depend on this
)

# ─── HELPER FUNCTIONS ────────────────────────────────────────────────────────

function Write-Phase {
    param([string]$Phase, [string]$Message)
    $ts = Get-Date -Format 'HH:mm:ss'
    Write-Host "[$ts] [$Phase] $Message" -ForegroundColor Cyan
}

function Write-Status {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'HH:mm:ss'
    $color = switch ($Level) {
        'PASS' { 'Green' }
        'FAIL' { 'Red' }
        'WARN' { 'Yellow' }
        'SKIP' { 'DarkGray' }
        default { 'White' }
    }
    Write-Host "  [$ts] $Message" -ForegroundColor $color
}

function Wait-ForWinRM {
    param([string]$Computer, [PSCredential]$Cred, [int]$TimeoutSec = 300)
    Write-Status "Waiting for WinRM (timeout: ${TimeoutSec}s)..."
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            $null = Invoke-Command -ComputerName $Computer -Credential $Cred -ScriptBlock { $true } -ErrorAction Stop
            Write-Status 'WinRM connected' -Level PASS
            return $true
        }
        catch {
            Start-Sleep -Seconds $winrmRetrySec
        }
    }
    Write-Status "WinRM timeout after ${TimeoutSec}s" -Level FAIL
    return $false
}

function Invoke-VMReboot {
    param([string]$Computer, [PSCredential]$Cred)
    Write-Status "Rebooting $Computer..."
    try {
        Invoke-Command -ComputerName $Computer -Credential $Cred -ScriptBlock {
            Restart-Computer -Force
        } -ErrorAction SilentlyContinue
    } catch {
        # Expected -- WinRM drops during reboot
    }
    Start-Sleep -Seconds 10
    return (Wait-ForWinRM -Computer $Computer -Cred $Cred -TimeoutSec $winrmTimeoutSec)
}

function Invoke-RemoteTest {
    <#
    .SYNOPSIS
        Runs Test-AllPackagers.ps1 for a single app on the remote VM with a timeout.
        Returns the deserialized result object.
    #>
    param(
        [string]$Computer,
        [PSCredential]$Cred,
        [string]$PackagerName,
        [string]$Phase = 'All',
        [switch]$IsStageOnly,
        [int]$TimeoutSec = 600
    )

    $remoteScript = "$remoteRepoRoot\Tests\Test-AllPackagers.ps1"
    $phaseToRun = if ($IsStageOnly) { 'Stage' } else { $Phase }

    $job = Invoke-Command -ComputerName $Computer -Credential $Cred -AsJob -ScriptBlock {
        param($script, $name, $phase)
        & $script -PackagerName $name -Phase $phase
    } -ArgumentList $remoteScript, $PackagerName, $phaseToRun

    $completed = $job | Wait-Job -Timeout $TimeoutSec
    if (-not $completed -or $job.State -eq 'Running') {
        $job | Stop-Job
        $job | Remove-Job -Force
        return [PSCustomObject]@{
            Name            = $PackagerName
            StageResult     = 'FAIL'
            Version         = ''
            ContentPath     = ''
            InstallResult   = 'SKIP'
            InstallExitCode = -1
            DetectionResult = 'SKIP'
            ProcessVerified = 'SKIP'
            ActualProcess   = ''
            ManifestProcess = ''
            UninstallResult = 'SKIP'
            UninstallExitCode = -1
            RemovalVerified = 'SKIP'
            Error           = "Timed out after ${TimeoutSec}s"
            NeedsReboot     = $false
            RebootPhase     = ''
        }
    }

    $output = $job | Receive-Job
    $job | Remove-Job -Force

    # Receive-Job may return multiple objects (Write-Host output + return value)
    # The result object is the last PSCustomObject
    if ($output -is [array]) {
        $resultObj = $output | Where-Object { $_ -is [PSCustomObject] -and $_.PSObject.Properties['Name'] } | Select-Object -Last 1
        if (-not $resultObj) { $resultObj = $output[-1] }
        return $resultObj
    }
    return $output
}

# ─── PHASE 0: PRE-FLIGHT ────────────────────────────────────────────────────

Write-Phase 'PREFLIGHT' 'Verifying prerequisites'

if (-not (Test-Path "$localRepoRoot\Packagers\package-7zip.ps1")) {
    throw "Local repo not found or incomplete: $localRepoRoot"
}
Write-Status "Local repo: $localRepoRoot"

$vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
if (-not $vm) {
    throw "VM not found: $VMName. Is Hyper-V running?"
}
Write-Status "VM found: $VMName (State: $($vm.State))"

if ($ResetClient) {
    $checkpoint = Get-VMSnapshot -VMName $VMName -Name $CheckpointName -ErrorAction SilentlyContinue
    if (-not $checkpoint) {
        throw "Checkpoint not found: '$CheckpointName' on VM '$VMName'"
    }
    Write-Status "Checkpoint found: $CheckpointName"
}

if (-not $Credential) {
    $Credential = New-Object PSCredential(
        'contoso\LabAdmin',
        (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
    )
}

if (-not (Test-Path $ResultsPath)) {
    New-Item -Path $ResultsPath -ItemType Directory -Force | Out-Null
}

# Build packager list from local repo
$allPackagers = Get-ChildItem "$localRepoRoot\Packagers\package-*.ps1" | Sort-Object Name
if ($IncludeOnly) {
    $allPackagers = $allPackagers | Where-Object { $_.BaseName -in $IncludeOnly }
}
$allPackagers = @($allPackagers | Where-Object { $_.BaseName -notin $SkipList })

Write-Status "Apps to test: $($allPackagers.Count) (skipping $($SkipList.Count))"
Write-Phase 'PREFLIGHT' 'All prerequisites verified'
Write-Host ''

# ─── PHASE 1: RESET CLIENT01 ────────────────────────────────────────────────

if ($ResetClient) {
    Write-Phase 'RESET' "Reverting $VMName to checkpoint: $CheckpointName"

    if ($PSCmdlet.ShouldProcess($VMName, "Stop VM, restore '$CheckpointName', restart")) {
        if ($vm.State -ne 'Off') {
            Write-Status "Stopping $VMName..."
            Stop-VM -Name $VMName -Force -TurnOff
            Start-Sleep -Seconds 5
        }

        Write-Status 'Restoring checkpoint...'
        Restore-VMSnapshot -VMName $VMName -Name $CheckpointName -Confirm:$false

        Write-Status "Starting $VMName..."
        Start-VM -Name $VMName

        if (-not (Wait-ForWinRM -Computer $VMName -Cred $Credential)) {
            throw "WinRM to $VMName did not become available within ${winrmTimeoutSec}s"
        }

        Write-Status 'Checking disk size...'
        Invoke-Command -ComputerName $VMName -Credential $Credential -ScriptBlock {
            $currentMax = (Get-PartitionSupportedSize -DriveLetter C).SizeMax
            $currentSize = (Get-Partition -DriveLetter C).Size
            if ($currentSize -lt $currentMax) {
                Resize-Partition -DriveLetter C -Size $currentMax
            }
        }
        Write-Status 'Disk verified'
    }

    Write-Phase 'RESET' 'CLIENT01 restored and ready'
    Write-Host ''
}

# ─── PHASE 2: DEPLOY SCRIPTS ────────────────────────────────────────────────

Write-Phase 'DEPLOY' 'Deploying packager scripts to CLIENT01'

if ($PSCmdlet.ShouldProcess($VMName, 'Copy applicationpackager repo')) {
    $session = New-PSSession -ComputerName $VMName -Credential $Credential

    try {
        Invoke-Command -Session $session -ScriptBlock {
            param($root)
            New-Item -Path "$root\Packagers" -ItemType Directory -Force | Out-Null
            New-Item -Path "$root\Tests" -ItemType Directory -Force | Out-Null
        } -ArgumentList $remoteRepoRoot

        Write-Status 'Copying Packagers/ ...'
        Copy-Item -Path "$localRepoRoot\Packagers\*" `
            -Destination "$remoteRepoRoot\Packagers\" `
            -ToSession $session -Force -Recurse

        Write-Status 'Copying Tests/ ...'
        Copy-Item -Path "$localRepoRoot\Tests\*" `
            -Destination "$remoteRepoRoot\Tests\" `
            -ToSession $session -Force -Recurse

        $fileCount = Invoke-Command -Session $session -ScriptBlock {
            param($root)
            (Get-ChildItem "$root\Packagers\package-*.ps1").Count
        } -ArgumentList $remoteRepoRoot

        Write-Status "Deployed $fileCount packager scripts" -Level PASS
    }
    finally {
        Remove-PSSession $session -ErrorAction SilentlyContinue
    }
}

Write-Phase 'DEPLOY' 'Scripts deployed'
Write-Host ''

# ─── PHASE 3: RUN TESTS (PER-APP LOOP) ──────────────────────────────────────

Write-Phase 'TEST' "Testing $($allPackagers.Count) packagers on CLIENT01"

$startTime = Get-Date
Write-Status "Start time: $($startTime.ToString('HH:mm:ss'))"
Write-Host ''

$results = [System.Collections.ArrayList]::new()
$rebootCount = 0
$appIndex = 0

foreach ($packager in $allPackagers) {
    $appIndex++
    $name = $packager.BaseName
    Write-Phase 'TEST' "[$appIndex/$($allPackagers.Count)] $name"

    if ($StageOnly) {
        # Stage-only mode
        $r = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
            -PackagerName $name -IsStageOnly -TimeoutSec $AppTimeoutSec
        $null = $results.Add($r)

        # Cleanup staged content
        if ($r.StageResult -eq 'PASS' -and $r.ContentPath) {
            Invoke-Command -ComputerName $VMName -Credential $Credential -ScriptBlock {
                param($path)
                $parent = Split-Path $path -Parent
                if ($parent -and (Test-Path $parent)) {
                    Remove-Item $parent -Recurse -Force -ErrorAction SilentlyContinue
                }
            } -ArgumentList "$($r.ContentPath)"
        }
        continue
    }

    $skipUninstall = $name -in $noUninstallList

    if ($skipUninstall) {
        # ── Install + Detect only (system dependency) ──
        Write-Status '(no-uninstall: system dependency)' -Level SKIP

        # Stage
        $r = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
            -PackagerName $name -Phase 'Stage' -TimeoutSec $AppTimeoutSec
        if ($r.StageResult -ne 'PASS') {
            $null = $results.Add($r)
            continue
        }

        # Install
        $ri = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
            -PackagerName $name -Phase 'Install' -TimeoutSec $AppTimeoutSec
        $r.InstallResult = $ri.InstallResult
        $r.InstallExitCode = $ri.InstallExitCode
        $r.NeedsReboot = $ri.NeedsReboot
        $r.RebootPhase = $ri.RebootPhase
        $r.Error = $ri.Error

        # Reboot if 3010
        if ($r.NeedsReboot -and $r.RebootPhase -eq 'Install') {
            $rebootCount++
            Write-Status "Reboot #$rebootCount (install 3010)..." -Level WARN
            if (-not (Invoke-VMReboot -Computer $VMName -Cred $Credential)) {
                $r.Error += ' | Reboot failed (WinRM timeout)'
                $null = $results.Add($r)
                continue
            }
            $r.NeedsReboot = $false
        }

        if ($r.InstallResult -eq 'PASS') {
            # Detect
            $rd = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
                -PackagerName $name -Phase 'Detect' -TimeoutSec $AppTimeoutSec
            $r.DetectionResult = $rd.DetectionResult
            $r.ProcessVerified = $rd.ProcessVerified
            $r.ActualProcess = $rd.ActualProcess
        }

        # Mark uninstall as intentionally skipped
        $r.UninstallResult = 'SKIP'
        $r.RemovalVerified = 'SKIP'

    } else {
        # ── Full E2E: Stage + Install + Detect + Uninstall ──
        $r = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
            -PackagerName $name -Phase 'All' -TimeoutSec $AppTimeoutSec

        # ── Handle install 3010: reboot then detect ──
        if ($r.NeedsReboot -and $r.RebootPhase -eq 'Install') {
            $rebootCount++
            Write-Status "Reboot #$rebootCount (install 3010)..." -Level WARN
            if (Invoke-VMReboot -Computer $VMName -Cred $Credential) {
                $r2 = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
                    -PackagerName $name -Phase 'Detect' -TimeoutSec $AppTimeoutSec
                $r.DetectionResult = $r2.DetectionResult
                $r.ProcessVerified = $r2.ProcessVerified
                $r.ActualProcess = $r2.ActualProcess
                $r.NeedsReboot = $false

                $r3 = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
                    -PackagerName $name -Phase 'Uninstall' -TimeoutSec $AppTimeoutSec
                $r.UninstallResult = $r3.UninstallResult
                $r.UninstallExitCode = $r3.UninstallExitCode

                if ($r3.NeedsReboot -and $r3.RebootPhase -eq 'Uninstall') {
                    $rebootCount++
                    Write-Status "Reboot #$rebootCount (uninstall 3010)..." -Level WARN
                    if (Invoke-VMReboot -Computer $VMName -Cred $Credential) {
                        $r4 = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
                            -PackagerName $name -Phase 'VerifyRemoval' -TimeoutSec $AppTimeoutSec
                        $r.RemovalVerified = $r4.RemovalVerified
                    }
                } else {
                    $r5 = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
                        -PackagerName $name -Phase 'VerifyRemoval' -TimeoutSec $AppTimeoutSec
                    $r.RemovalVerified = $r5.RemovalVerified
                }
            } else {
                $r.Error += ' | Reboot failed (WinRM timeout)'
            }
        }

        # ── Handle uninstall 3010: reboot then verify removal ──
        if ($r.NeedsReboot -and $r.RebootPhase -eq 'Uninstall') {
            $rebootCount++
            Write-Status "Reboot #$rebootCount (uninstall 3010)..." -Level WARN
            if (Invoke-VMReboot -Computer $VMName -Cred $Credential) {
                $r6 = Invoke-RemoteTest -Computer $VMName -Cred $Credential `
                    -PackagerName $name -Phase 'VerifyRemoval' -TimeoutSec $AppTimeoutSec
                $r.RemovalVerified = $r6.RemovalVerified
            } else {
                $r.Error += ' | Reboot failed (WinRM timeout)'
            }
        }
    }

    # ── Cleanup staged content ──
    if ($r.ContentPath) {
        Invoke-Command -ComputerName $VMName -Credential $Credential -ScriptBlock {
            param($path)
            $parent = Split-Path $path -Parent
            if ($parent -and (Test-Path $parent)) {
                Remove-Item $parent -Recurse -Force -ErrorAction SilentlyContinue
            }
        } -ArgumentList "$($r.ContentPath)"
    }

    # ── Log summary line ──
    $statusColor = if ($r.InstallResult -eq 'PASS' -and $r.DetectionResult -eq 'PASS' -and $r.UninstallResult -eq 'PASS') { 'PASS' } else { 'WARN' }
    if ($r.InstallResult -eq 'FAIL' -or $r.DetectionResult -eq 'FAIL') { $statusColor = 'FAIL' }
    Write-Status ("S:{0} I:{1} D:{2} P:{3} U:{4} R:{5} {6}" -f `
        $r.StageResult, $r.InstallResult, $r.DetectionResult, `
        $r.ProcessVerified, $r.UninstallResult, $r.RemovalVerified, `
        $(if ($r.Error) { "[$($r.Error)]" } else { '' })) -Level $statusColor

    $null = $results.Add($r)

    # ── Save incremental results after each app ──
    $results | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $ResultsPath 'packager-results-incremental.json') -Encoding ASCII
}

$elapsed = (Get-Date) - $startTime

Write-Host ''
Write-Phase 'TEST' "Completed in $($elapsed.ToString('hh\:mm\:ss')) ($rebootCount reboots)"
Write-Host ''

# ─── PHASE 4: RESULTS ───────────────────────────────────────────────────────

Write-Phase 'RESULTS' 'Summary'

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$localResultsFile = Join-Path $ResultsPath "packager-results-$timestamp.json"
$results | ConvertTo-Json -Depth 4 | Set-Content $localResultsFile -Encoding ASCII

$total      = $results.Count
$stagePassed = @($results | Where-Object { $_.StageResult -eq 'PASS' }).Count
$stageFail   = @($results | Where-Object { $_.StageResult -eq 'FAIL' })
$fullPass    = @($results | Where-Object { $_.InstallResult -eq 'PASS' -and $_.DetectionResult -eq 'PASS' -and $_.UninstallResult -eq 'PASS' }).Count
$installFail = @($results | Where-Object { $_.InstallResult -eq 'FAIL' })
$detectFail  = @($results | Where-Object { $_.DetectionResult -eq 'FAIL' })
$processFail = @($results | Where-Object { $_.ProcessVerified -eq 'FAIL' })
$uninstFail  = @($results | Where-Object { $_.UninstallResult -eq 'FAIL' })
$timedOut    = @($results | Where-Object { $_.Error -match 'Timed out' })

Write-Host ''
Write-Host ('=' * 70)
Write-Host 'REGRESSION TEST SUMMARY'
Write-Host ('=' * 70)
Write-Host ''
Write-Host "  Total tested     : $total"
Write-Host "  Staged OK        : $stagePassed"
if (-not $StageOnly) {
    Write-Host "  Full E2E pass    : $fullPass" -ForegroundColor $(if ($fullPass -eq $total) { 'Green' } else { 'Yellow' })
}
Write-Host "  Reboots          : $rebootCount"
Write-Host ''

if ($stageFail.Count -gt 0) {
    Write-Host '  STAGE FAILURES:' -ForegroundColor Red
    $stageFail | ForEach-Object { Write-Host "    $($_.Name): $($_.Error)" -ForegroundColor Red }
    Write-Host ''
}
if ($installFail.Count -gt 0) {
    Write-Host '  INSTALL FAILURES:' -ForegroundColor Red
    $installFail | ForEach-Object { Write-Host "    $($_.Name): $($_.Error)" -ForegroundColor Red }
    Write-Host ''
}
if ($detectFail.Count -gt 0) {
    Write-Host '  DETECTION FAILURES:' -ForegroundColor Red
    $detectFail | ForEach-Object { Write-Host "    $($_.Name): $($_.Error)" -ForegroundColor Red }
    Write-Host ''
}
if ($uninstFail.Count -gt 0) {
    Write-Host '  UNINSTALL FAILURES:' -ForegroundColor Red
    $uninstFail | ForEach-Object { Write-Host "    $($_.Name): $($_.Error)" -ForegroundColor Red }
    Write-Host ''
}
if ($processFail.Count -gt 0) {
    Write-Host '  PROCESS NAME ISSUES:' -ForegroundColor Yellow
    $processFail | ForEach-Object { Write-Host "    $($_.Name): $($_.ActualProcess)" -ForegroundColor Yellow }
    Write-Host ''
}
if ($timedOut.Count -gt 0) {
    Write-Host '  TIMED OUT:' -ForegroundColor Yellow
    $timedOut | ForEach-Object { Write-Host "    $($_.Name)" -ForegroundColor Yellow }
    Write-Host ''
}

Write-Host "  Results: $localResultsFile"
Write-Host ''

# Exit code
$failCount = $installFail.Count + $detectFail.Count + $stageFail.Count
if ($failCount -gt 0) {
    Write-Phase 'DONE' "$failCount failure(s) detected"
    exit 1
}

Write-Phase 'DONE' 'All tests passed'
