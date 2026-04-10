#Requires -RunAsAdministrator
#Requires -Modules Hyper-V
<#
.SYNOPSIS
    Remote orchestrator for Application Packager regression testing on CLIENT01.

.DESCRIPTION
    Runs from the Hyper-V host. Optionally reverts CLIENT01 to a clean snapshot,
    deploys the current packager scripts via WinRM, executes Test-AllPackagers.ps1
    on the remote VM, and collects results back to the host.

    Wraps Test-AllPackagers.ps1 -- does not replace it. The inner harness handles
    stage/install/detect/uninstall/verify. This script handles VM lifecycle,
    file deployment, remote invocation, and result collection.

.PARAMETER ResetClient
    Revert CLIENT01 to the Deployment-Complete snapshot before testing.
    The VM is stopped, restored, restarted, and WinRM connectivity is verified.

.PARAMETER StageOnly
    Only test staging (download + manifest). No install/uninstall.
    Fast smoke test for version scraping and content wrapper generation.

.PARAMETER IncludeOnly
    Test specific packagers only. Array of base names (e.g., "package-7zip").
    Passed through to Test-AllPackagers.ps1.

.PARAMETER SkipList
    Override the default skip list in Test-AllPackagers.ps1.
    Passed through to the inner harness.

.PARAMETER ResultsPath
    Local host directory for collected results.
    Default: C:\temp\regression-results

.PARAMETER VMName
    Hyper-V VM name. Default: CLIENT01

.PARAMETER CheckpointName
    Snapshot name to restore when using -ResetClient.
    Default: Deployment-Complete

.PARAMETER Credential
    PSCredential for WinRM to the VM. If omitted, prompts or uses default
    domain credentials (contoso\LabAdmin).

.PARAMETER WhatIf
    Show what would happen without executing. Useful for verifying connectivity.

.EXAMPLE
    .\Invoke-PackagerRegressionTest.ps1 -ResetClient
    Full regression: reset CLIENT01, deploy scripts, run all ~97 apps, collect results.

.EXAMPLE
    .\Invoke-PackagerRegressionTest.ps1 -ResetClient -StageOnly
    Fast smoke: reset CLIENT01, stage-only for all apps (no install/uninstall).

.EXAMPLE
    .\Invoke-PackagerRegressionTest.ps1 -IncludeOnly @('package-7zip','package-vlc')
    Quick test of two specific apps (no reset, no deploy if already present).

.EXAMPLE
    .\Invoke-PackagerRegressionTest.ps1 -ResetClient -WhatIf
    Dry run: show what would happen without executing.
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
    [PSCredential]$Credential
)

$ErrorActionPreference = 'Stop'

$localRepoRoot = 'c:\projects\applicationpackager'
$remoteRepoRoot = 'C:\projects\applicationpackager'
$remoteResultsJson = 'c:\temp\packager-test-results.json'
$winrmTimeoutSec = 300
$winrmRetrySec = 15
$testTimeoutSec = 14400  # 4 hours

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
        default { 'White' }
    }
    Write-Host "  [$ts] $Message" -ForegroundColor $color
}

# ─── PHASE 0: PRE-FLIGHT ─────────────────────────────────────────────────────

Write-Phase 'PREFLIGHT' 'Verifying prerequisites'

# Verify local repo exists
if (-not (Test-Path "$localRepoRoot\Packagers\package-7zip.ps1")) {
    throw "Local repo not found or incomplete: $localRepoRoot"
}
Write-Status "Local repo: $localRepoRoot"

# Verify VM exists
$vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
if (-not $vm) {
    throw "VM not found: $VMName. Is Hyper-V running?"
}
Write-Status "VM found: $VMName (State: $($vm.State))"

# Verify checkpoint exists (if reset requested)
if ($ResetClient) {
    $checkpoint = Get-VMSnapshot -VMName $VMName -Name $CheckpointName -ErrorAction SilentlyContinue
    if (-not $checkpoint) {
        throw "Checkpoint not found: '$CheckpointName' on VM '$VMName'"
    }
    Write-Status "Checkpoint found: $CheckpointName"
}

# Build credential if not provided
if (-not $Credential) {
    $Credential = New-Object PSCredential(
        'contoso\LabAdmin',
        (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force)
    )
}

# Results directory
if (-not (Test-Path $ResultsPath)) {
    New-Item -Path $ResultsPath -ItemType Directory -Force | Out-Null
}

Write-Phase 'PREFLIGHT' 'All prerequisites verified'
Write-Host ''

# ─── PHASE 1: RESET CLIENT01 ─────────────────────────────────────────────────

if ($ResetClient) {
    Write-Phase 'RESET' "Reverting $VMName to checkpoint: $CheckpointName"

    if ($PSCmdlet.ShouldProcess($VMName, "Stop VM, restore '$CheckpointName', restart")) {
        # Stop VM if running
        if ($vm.State -ne 'Off') {
            Write-Status "Stopping $VMName..."
            Stop-VM -Name $VMName -Force -TurnOff
            Start-Sleep -Seconds 5
        }

        # Restore checkpoint
        Write-Status 'Restoring checkpoint...'
        Restore-VMSnapshot -VMName $VMName -Name $CheckpointName -Confirm:$false

        # Start VM
        Write-Status "Starting $VMName..."
        Start-VM -Name $VMName

        # Wait for WinRM
        Write-Status "Waiting for WinRM (timeout: ${winrmTimeoutSec}s)..."
        $deadline = (Get-Date).AddSeconds($winrmTimeoutSec)
        $connected = $false
        while ((Get-Date) -lt $deadline) {
            try {
                $null = Invoke-Command -ComputerName $VMName -Credential $Credential -ScriptBlock { $true } -ErrorAction Stop
                $connected = $true
                break
            }
            catch {
                Start-Sleep -Seconds $winrmRetrySec
            }
        }
        if (-not $connected) {
            throw "WinRM to $VMName did not become available within ${winrmTimeoutSec}s"
        }
        Write-Status 'WinRM connected' -Level PASS

        # Expand C: drive if needed (matches Deploy-HomeLab.ps1 logic)
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

# ─── PHASE 2: DEPLOY SCRIPTS ─────────────────────────────────────────────────

Write-Phase 'DEPLOY' 'Deploying packager scripts to CLIENT01'

if ($PSCmdlet.ShouldProcess($VMName, 'Copy applicationpackager repo')) {
    $session = New-PSSession -ComputerName $VMName -Credential $Credential

    try {
        # Create remote directory structure
        Invoke-Command -Session $session -ScriptBlock {
            param($root)
            New-Item -Path "$root\Packagers" -ItemType Directory -Force | Out-Null
            New-Item -Path "$root\Tests" -ItemType Directory -Force | Out-Null
        } -ArgumentList $remoteRepoRoot

        # Copy Packagers/ folder
        Write-Status 'Copying Packagers/ ...'
        Copy-Item -Path "$localRepoRoot\Packagers\*" `
            -Destination "$remoteRepoRoot\Packagers\" `
            -ToSession $session -Force -Recurse

        # Copy Tests/ folder
        Write-Status 'Copying Tests/ ...'
        Copy-Item -Path "$localRepoRoot\Tests\*" `
            -Destination "$remoteRepoRoot\Tests\" `
            -ToSession $session -Force -Recurse

        # Verify deployment
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

# ─── PHASE 3: RUN TESTS ──────────────────────────────────────────────────────

Write-Phase 'TEST' 'Running Test-AllPackagers.ps1 on CLIENT01'

if ($PSCmdlet.ShouldProcess($VMName, 'Run Test-AllPackagers.ps1')) {
    $testScript = "$remoteRepoRoot\Tests\Test-AllPackagers.ps1"

    # Build argument hashtable for splatting on the remote side
    $remoteArgs = @{}
    if ($StageOnly) { $remoteArgs['StageOnly'] = $true }
    if ($IncludeOnly) { $remoteArgs['IncludeOnly'] = $IncludeOnly }
    if ($SkipList) { $remoteArgs['SkipList'] = $SkipList }

    $startTime = Get-Date
    Write-Status "Start time: $($startTime.ToString('HH:mm:ss'))"
    Write-Status "Timeout: $($testTimeoutSec / 3600) hours"
    Write-Host ''

    # Run the test harness remotely
    # -InDisconnectedSession is not used -- we want to stream output in real time
    $testOutput = Invoke-Command -ComputerName $VMName -Credential $Credential -ScriptBlock {
        param($script, $args_)
        Set-Location C:\
        & $script @args_
    } -ArgumentList $testScript, $remoteArgs

    $elapsed = (Get-Date) - $startTime

    # Display output
    if ($testOutput) {
        $testOutput | ForEach-Object { Write-Host $_ }
    }

    Write-Host ''
    Write-Phase 'TEST' "Completed in $($elapsed.ToString('hh\:mm\:ss'))"
    Write-Host ''
}

# ─── PHASE 4: COLLECT RESULTS ────────────────────────────────────────────────

Write-Phase 'COLLECT' 'Retrieving results from CLIENT01'

if ($PSCmdlet.ShouldProcess($VMName, 'Collect test results')) {
    $session = New-PSSession -ComputerName $VMName -Credential $Credential

    try {
        # Check if results file exists
        $hasResults = Invoke-Command -Session $session -ScriptBlock {
            param($path)
            Test-Path $path
        } -ArgumentList $remoteResultsJson

        if ($hasResults) {
            $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
            $localResultsFile = Join-Path $ResultsPath "packager-results-$timestamp.json"

            Copy-Item -Path $remoteResultsJson `
                -Destination $localResultsFile `
                -FromSession $session

            Write-Status "Results saved: $localResultsFile" -Level PASS

            # Parse and display summary
            $results = Get-Content $localResultsFile -Raw | ConvertFrom-Json

            $total   = $results.Count
            $passed  = ($results | Where-Object { $_.InstallResult -eq 'PASS' -and $_.DetectionResult -eq 'PASS' -and $_.UninstallResult -eq 'PASS' }).Count
            $staged  = ($results | Where-Object { $_.StageResult -eq 'PASS' }).Count
            $failed  = ($results | Where-Object { $_.InstallResult -eq 'FAIL' -or $_.DetectionResult -eq 'FAIL' }).Count
            $detectFail = ($results | Where-Object { $_.DetectionResult -eq 'FAIL' })
            $installFail = ($results | Where-Object { $_.InstallResult -eq 'FAIL' })
            $processFail = ($results | Where-Object { $_.ProcessVerified -eq 'FAIL' })

            Write-Host ''
            Write-Host ('=' * 70)
            Write-Host 'REGRESSION TEST SUMMARY'
            Write-Host ('=' * 70)
            Write-Host ''
            Write-Host "  Total tested     : $total"
            Write-Host "  Staged OK        : $staged"
            Write-Host "  Full pass        : $passed" -ForegroundColor $(if ($passed -eq $total) { 'Green' } else { 'Yellow' })
            Write-Host "  Failed           : $failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Green' })
            Write-Host ''

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
            if ($processFail.Count -gt 0) {
                Write-Host '  PROCESS NAME ISSUES:' -ForegroundColor Yellow
                $processFail | ForEach-Object { Write-Host "    $($_.Name): $($_.ActualProcess)" -ForegroundColor Yellow }
                Write-Host ''
            }

            Write-Host "  Results file: $localResultsFile"
            Write-Host ''

            # Exit code
            if ($failed -gt 0) {
                Write-Phase 'DONE' "$failed failures detected"
                exit 1
            }
        }
        else {
            Write-Status 'No results file found on CLIENT01 (test may not have completed)' -Level WARN
            exit 1
        }
    }
    finally {
        Remove-PSSession $session -ErrorAction SilentlyContinue
    }
}

Write-Phase 'DONE' 'Regression test complete — all passed'
