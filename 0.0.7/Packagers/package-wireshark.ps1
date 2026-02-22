<#
Vendor: The Wireshark developer community
App: Wireshark
CMName: Wireshark

.SYNOPSIS
    Wireshark Auto-Packager (EXE) for MECM

.DESCRIPTION
    - Finds latest Wireshark "Stable Release" version from wireshark.org
    - Downloads Wireshark-<version>-x64.exe
    - Copies to: <FileServerPath>\Applications\Wireshark Foundation\Wireshark\<Version>\
    - Creates install.bat / uninstall.bat in that folder
    - Temporarily installs Wireshark locally to extract uninstall registry metadata:
        DisplayName, DisplayVersion, Publisher, InstallLocation, QuietUninstallString, UninstallString, PSChildName (key name), root (WOW6432 vs native)
      Then uninstalls to return packaging machine to clean state.
    - Creates MECM Application + Script Deployment Type
    - Detection: Registry string DisplayVersion MUST equal packaged version (e.g., 4.6.3)
      (Uses the actual uninstall hive discovered during extraction, even if WOW6432Node)
    - DeploymentType "Content" settings enforced:
        - Allow clients to use DP from default site boundary group (Content Fallback)
        - Deployment options: Download content from DP and run locally (Slow network mode: Download)

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM"). The PSDrive is assumed to already exist.
.PARAMETER Comment
    Work order or comment string applied to the MECM application description.
.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").
.PARAMETER GetLatestVersionOnly
    Outputs only the latest version string and exits.

.NOTES
    Requires:
      - PowerShell 5.1
      - ConfigMgr Admin Console installed (for ConfigurationManager.psd1)
      - RBAC rights to create Applications and Deployment Types
      - Local administrator (script temporarily installs Wireshark to extract registry metadata)
      - Write access to FileServerPath
    File/share operations NEVER occur while current location is the CM PSDrive (e.g., MCM:).
#>

# ----------------------------
# CONFIG
# ----------------------------
param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$WiresharkDownloadPage = "https://www.wireshark.org/download.html"
$WiresharkWin64Root    = "https://www.wireshark.org/download/win64"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\Wireshark_installers"
$WiresharkRootNetworkPath = Join-Path $FileServerPath "Applications\Wireshark Foundation\Wireshark"

# Wireshark installer switches you want (these are passed to the EXE)
$DesktopIconSetting     = "no"  # yes|no
$QuickLaunchIconSetting = "no"  # yes|no

# Polling
$InitialInstallBufferSeconds = 60
$PollSleepSeconds            = 60
$MaxRegistryPollRetries      = 20

# MECM DT runtime
$EstimatedRuntimeMins = 10
$MaximumRuntimeMins   = 30

# MECM DT Content tab settings (the two "jank network" items)
$EnableContentFallback      = $true     # Allow fallback to default site boundary group DPs
$SlowNetworkDeploymentMode  = "Download" # Download content from DP and run locally

# ----------------------------
# HELPERS / FUNCTIONS
# ----------------------------

function Test-IsAdmin {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Warning "Failed to check admin privileges: $($_.Exception.Message)"
        return $false
    }
}

function Push-FSLocation {
    # Ensures file ops don’t happen on CM PSDrive.
    param([string]$Why)

    $cur = Get-Location
    Write-Host "Current location before FS op ($Why): $cur"
    Set-Location $PSScriptRoot -ErrorAction Stop
    Write-Host "Set location to script directory for FS op ($Why): $PSScriptRoot"
    return $cur
}

function Pop-Location {
    param([object]$OriginalLocation, [string]$Why)
    try {
        Set-Location $OriginalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored location after ($Why) to: $OriginalLocation"
    }
    catch {
        Write-Warning "Failed restoring location after ($Why): $($_.Exception.Message)"
    }
}

function Test-NetworkShareAccess {
    param ([string]$Path)

    $orig = Push-FSLocation -Why "Network share validation"
    try {
        if (-not $Path) { throw "Network path is null/empty." }
        if (-not (Test-Path $Path -ErrorAction Stop)) { throw "Network path not accessible: $Path" }

        $testFile = Join-Path $Path "test_$(Get-Random).txt"
        Set-Content -Path $testFile -Value "Test" -ErrorAction Stop
        Remove-Item -Path $testFile -ErrorAction Stop

        Write-Host "Network share is accessible and writable: $Path"
        return $true
    }
    catch {
        Write-Error "Network share validation failed: $($_.Exception.Message)"
        return $false
    }
    finally {
        Pop-Location -OriginalLocation $orig -Why "Network share validation"
    }
}

function Get-LatestWiresharkVersion {
    param([switch]$Quiet)

    if (-not $Quiet) { Write-Host "Fetching Wireshark download page: $WiresharkDownloadPage" }
    try {
        $html = (curl.exe -L --fail --silent --show-error $WiresharkDownloadPage) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Wireshark download page: $WiresharkDownloadPage" }
        if ($html -match 'Stable Release:\s*([0-9]+\.[0-9]+\.[0-9]+)') {
            $v = $matches[1]
            if (-not $Quiet) { Write-Host "Found latest Wireshark stable version: $v" }
            return $v
        }
        throw "Could not parse Stable Release version from download page."
    }
    catch {
        Write-Error "Failed to determine latest Wireshark version: $($_.Exception.Message)"
        return $null
    }
}

function Split-CommandLine {
    param([string]$CommandLine)
    if (-not $CommandLine) { return $null }

    $cmd = $CommandLine.Trim()

    if ($cmd.StartsWith('"')) {
        $secondQuote = $cmd.IndexOf('"', 1)
        if ($secondQuote -gt 1) {
            $exe  = $cmd.Substring(1, $secondQuote - 1)
            $args = $cmd.Substring($secondQuote + 1).Trim()
            return @{ FilePath = $exe; Arguments = $args }
        }
    }

    $parts = $cmd.Split(@(' '), 2, [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($parts.Count -eq 1) { return @{ FilePath = $parts[0]; Arguments = "" } }
    return @{ FilePath = $parts[0]; Arguments = $parts[1] }
}

function Convert-RegRootToCMKeyName {
    <#
        Input:
          HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall
        Output:
          SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\<KeyName>
    #>
    param(
        [Parameter(Mandatory)][string]$UninstallRootPSPath,
        [Parameter(Mandatory)][string]$PSChildName
    )
    $cmBase = $UninstallRootPSPath -replace '^HKLM:\\', ''
    return "$cmBase\$PSChildName"
}

function Install-And-Extract-RegistryData {
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [Parameter(Mandatory)][string]$DisplayNamePrefix
    )

    $orig = Push-FSLocation -Why "Temp install + registry crawl"
    try {
        Write-Host "Temporarily installing Wireshark to extract uninstall metadata..."
        Write-Host "  Installer: $InstallerPath"
        $installArgs = "/S /desktopicon=$DesktopIconSetting /quicklaunchicon=$QuickLaunchIconSetting"
        Write-Host "  Args     : $installArgs"

        Start-Process -FilePath $InstallerPath -ArgumentList $installArgs -Wait -NoNewWindow

        Write-Host "Initial buffer after install: $InitialInstallBufferSeconds seconds"
        Start-Sleep -Seconds $InitialInstallBufferSeconds

        # You discovered it lands under WOW6432Node; search that FIRST (still scan both).
        $uninstallRoots = @(
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        )

        $retry = 0
        $data = $null
        $pattern = "$DisplayNamePrefix*"

        do {
            $retry++
            Write-Host "Registry poll attempt $retry/$MaxRegistryPollRetries (pattern: '$pattern')"

            foreach ($root in $uninstallRoots) {
                Write-Host "  Scanning root: $root"
                $keys = Get-ChildItem -Path $root -ErrorAction SilentlyContinue
                if (-not $keys) { continue }

                foreach ($k in $keys) {
                    $p = $k.PSPath
                    $dn = (Get-ItemProperty -Path $p -Name "DisplayName" -ErrorAction SilentlyContinue).DisplayName
                    if ($dn -and ($dn -like $pattern)) {
                        $props = Get-ItemProperty -Path $p -ErrorAction SilentlyContinue

                        $data = [ordered]@{
                            UninstallRoot        = $root
                            PSChildName          = $k.PSChildName
                            DisplayName          = $props.DisplayName
                            DisplayVersion       = $props.DisplayVersion
                            Publisher            = $props.Publisher
                            InstallLocation      = $props.InstallLocation
                            QuietUninstallString = $props.QuietUninstallString
                            UninstallString      = $props.UninstallString
                        }

                        Write-Host "  MATCH FOUND:"
                        Write-Host "    Root               : $($data.UninstallRoot)"
                        Write-Host "    Key                : $($data.PSChildName)"
                        Write-Host "    DisplayName        : $($data.DisplayName)"
                        Write-Host "    DisplayVersion     : $($data.DisplayVersion)"
                        Write-Host "    Publisher          : $($data.Publisher)"
                        Write-Host "    InstallLocation    : $($data.InstallLocation)"
                        Write-Host "    QuietUninstallString: $($data.QuietUninstallString)"
                        Write-Host "    UninstallString    : $($data.UninstallString)"
                        break
                    }
                }

                if ($data) { break }
            }

            if (-not $data -and $retry -lt $MaxRegistryPollRetries) {
                Write-Host "  No match yet. Sleeping $PollSleepSeconds seconds..."
                Start-Sleep -Seconds $PollSleepSeconds
            }

        } while (-not $data -and $retry -lt $MaxRegistryPollRetries)

        if (-not $data) {
            throw "No uninstall registry entry found for '$DisplayNamePrefix' after $MaxRegistryPollRetries polls."
        }

        # Uninstall to clean machine
        Write-Host "Uninstalling Wireshark to return packaging machine to clean state..."

        $uninstallCmd = $null
        if ($data.QuietUninstallString) {
            $uninstallCmd = $data.QuietUninstallString
            Write-Host "  Using QuietUninstallString."
        }
        elseif ($data.UninstallString) {
            $uninstallCmd = $data.UninstallString
            Write-Host "  QuietUninstallString missing; using UninstallString."
        }
        else {
            $fallback = Join-Path $env:ProgramFiles "Wireshark\uninstall.exe"
            if (Test-Path $fallback) {
                $uninstallCmd = "`"$fallback`" /S"
                Write-Host "  Registry uninstall strings missing; using fallback: $uninstallCmd"
            }
        }

        if ($uninstallCmd) {
            $parsed = Split-CommandLine -CommandLine $uninstallCmd
            if ($parsed -and $parsed.FilePath) {
                Write-Host "  Uninstall FilePath : $($parsed.FilePath)"
                Write-Host "  Uninstall Args     : $($parsed.Arguments)"
                Start-Process -FilePath $parsed.FilePath -ArgumentList $parsed.Arguments -Wait -NoNewWindow
                Start-Sleep -Seconds 30
            }
            else {
                Write-Warning "  Could not parse uninstall command: $uninstallCmd"
            }
        }
        else {
            Write-Warning "  No uninstall command found. Machine may not be clean."
        }

        return $data
    }
    finally {
        Pop-Location -OriginalLocation $orig -Why "Temp install + registry crawl"
    }
}

function Create-BatchFiles {
    param(
        [Parameter(Mandatory)][string]$NetworkPath,
        [Parameter(Mandatory)][string]$InstallerFileName,
        [Parameter(Mandatory)][string]$DisplayNamePrefix
    )

    $orig = Push-FSLocation -Why "Create batch files"
    try {
        $installBatPath   = Join-Path $NetworkPath "install.bat"
        $uninstallBatPath = Join-Path $NetworkPath "uninstall.bat"

        $installBat = @"
@echo off
setlocal

REM Wireshark silent install
start /wait "" "%~dp0$InstallerFileName" /S /desktopicon=$DesktopIconSetting /quicklaunchicon=$QuickLaunchIconSetting

exit /b %ERRORLEVEL%
"@

        # Uninstall wrapper:
        # - scans BOTH uninstall roots
        # - finds first DisplayName that begins with "$DisplayNamePrefix"
        # - prefers QuietUninstallString
        $uninstallBat = @"
@echo off
setlocal EnableExtensions

set "ROOT1=HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
set "ROOT2=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
set "PREFIX=$DisplayNamePrefix"

set "KEY="

for %%R in ("%ROOT1%" "%ROOT2%") do (
  for /f "delims=" %%K in ('reg query "%%~R" 2^>nul') do (
    for /f "tokens=2,*" %%A in ('reg query "%%K" /v DisplayName 2^>nul ^| find /i "DisplayName"') do (
      echo %%B | findstr /i /b "%PREFIX%" >nul
      if !errorlevel!==0 (
        set "KEY=%%K"
        goto :FOUND
      )
    )
  )
)

:FOUND
if defined KEY (
  for /f "tokens=2,*" %%A in ('reg query "%KEY%" /v QuietUninstallString 2^>nul ^| find /i "QuietUninstallString"') do set "QUIET=%%B"
  if defined QUIET (
    cmd.exe /c %QUIET%
    exit /b %ERRORLEVEL%
  )
  for /f "tokens=2,*" %%A in ('reg query "%KEY%" /v UninstallString 2^>nul ^| find /i "UninstallString"') do set "UNINST=%%B"
  if defined UNINST (
    cmd.exe /c %UNINST%
    exit /b %ERRORLEVEL%
  )
)

if exist "%ProgramFiles%\Wireshark\uninstall.exe" (
  "%ProgramFiles%\Wireshark\uninstall.exe" /S
  exit /b %ERRORLEVEL%
)

exit /b 0
"@

        Set-Content -Path $installBatPath   -Value $installBat   -Encoding ASCII -ErrorAction Stop
        Set-Content -Path $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop

        Write-Host "Created:"
        Write-Host "  $installBatPath"
        Write-Host "  $uninstallBatPath"
    }
    finally {
        Pop-Location -OriginalLocation $orig -Why "Create batch files"
    }
}

function Connect-CMSite {
    param([Parameter(Mandatory)][string]$SiteCode)

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
        Write-Host "Connected to CM site: $SiteCode"
        return $true
    }
    catch {
        Write-Error "Failed to connect to CM site: $($_.Exception.Message)"
        return $false
    }
}

function Remove-CMApplicationRevisionHistoryByCIId {
    param(
        [Parameter(Mandatory)][UInt32]$CI_ID,
        [UInt32]$KeepLatest = 1
    )
    $history = Get-CMApplicationRevisionHistory -Id $CI_ID -ErrorAction SilentlyContinue
    if (-not $history) { return }
    $revs = @()
    foreach ($h in @($history)) {
        if ($h.PSObject.Properties.Name -contains 'Revision')  { $revs += [UInt32]$h.Revision;   continue }
        if ($h.PSObject.Properties.Name -contains 'CIVersion') { $revs += [UInt32]$h.CIVersion; continue }
    }
    $revs = $revs | Sort-Object -Unique -Descending
    if ($revs.Count -le $KeepLatest) { return }
    foreach ($rev in ($revs | Select-Object -Skip $KeepLatest)) {
        Remove-CMApplicationRevisionHistory -Id $CI_ID -Revision $rev -Force -ErrorAction Stop
    }
}

function New-WiresharkMECMApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$Publisher,
        [Parameter(Mandatory)][string]$ContentLocation,
        [Parameter(Mandatory)][string]$UninstallRootPSPath,
        [Parameter(Mandatory)][string]$UninstallKeyName
    )

    $orig = Get-Location
    Write-Host "Current location before MECM work: $orig"

    try {
        if (-not (Connect-CMSite -SiteCode $SiteCode)) {
            throw "Could not connect to CM site."
        }

        # IMPORTANT: We are now in CM PSDrive. NO FILE OPS beyond this point.
        Write-Host "MECM App Create/Update:"
        Write-Host "  AppName       : $AppName"
        Write-Host "  Version       : $Version"
        Write-Host "  Publisher     : $Publisher"
        Write-Host "  ContentLoc    : $ContentLocation"
        Write-Host "  ContentFallback: $EnableContentFallback"
        Write-Host "  SlowNetworkMode: $SlowNetworkDeploymentMode"
        Write-Host "  Detection root : $UninstallRootPSPath"
        Write-Host "  Detection key  : $UninstallKeyName"
        Write-Host "  Detection value: DisplayVersion == $Version"

        $existing = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if (-not $existing) {
            Write-Host "Creating new CM Application: $AppName"
            $cmApp = New-CMApplication -Name $AppName -Publisher $Publisher -SoftwareVersion $Version -Description $Comment -LocalizedApplicationName $AppName -ErrorAction Stop
        }
        else {
            Write-Host "Application already exists: $AppName"
            $cmApp = $existing
        }

        $cmKeyName = Convert-RegRootToCMKeyName -UninstallRootPSPath $UninstallRootPSPath -PSChildName $UninstallKeyName

        # THIS is the detection you asked for: DisplayVersion string equals packaged version.
        $detect = New-CMDetectionClauseRegistryKey `
            -Hive LocalMachine `
            -KeyName $cmKeyName `
            -ValueName "DisplayVersion" `
            -PropertyType String `
            -ExpectedValue $Version `
            -ExpressionOperator IsEquals

        $dtName = "$AppName Script DT"

        # If DT exists, enforce the two Content settings too (don’t rely on defaults).
        $dtExisting = Get-CMDeploymentType -ApplicationName $AppName -DeploymentTypeName $dtName -ErrorAction SilentlyContinue

        if ($dtExisting) {
            Write-Host "Deployment Type already exists: $dtName"
            Write-Host "Enforcing DT Content settings (fallback + download/run locally)..."

            # Some environments use Set-CMScriptDeploymentType for updates; enforce what we can.
            Set-CMScriptDeploymentType `
                -ApplicationName $AppName `
                -DeploymentTypeName $dtName `
                -ContentFallback:$EnableContentFallback `
                -SlowNetworkDeploymentMode $SlowNetworkDeploymentMode `
                -ErrorAction Stop | Out-Null

            Write-Host "NOTE: Existing DT detected. If you need to update detection/version on an existing DT, delete/recreate DT per your standard."
            return
        }

        Write-Host "Creating Script Deployment Type: $dtName"
        Write-Host "  InstallCommand  : install.bat"
        Write-Host "  UninstallCommand: uninstall.bat"
        Write-Host "  ContentLocation : $ContentLocation"
        Write-Host "  Detection       : HKLM\$cmKeyName DisplayVersion == $Version"
        Write-Host "  ContentFallback : $EnableContentFallback"
        Write-Host "  SlowNetworkMode : $SlowNetworkDeploymentMode"

        $params = @{
            ApplicationName          = $AppName
            DeploymentTypeName       = $dtName
            InstallCommand           = "install.bat"
            UninstallCommand         = "uninstall.bat"
            ContentLocation          = $ContentLocation
            InstallationBehaviorType = "InstallForSystem"
            LogonRequirementType     = "WhetherOrNotUserLoggedOn"
            EstimatedRuntimeMins     = $EstimatedRuntimeMins
            MaximumRuntimeMins       = $MaximumRuntimeMins
            AddDetectionClause       = $detect
            ErrorAction              = "Stop"
        }

        # Apply your two mandatory content settings
        if ($EnableContentFallback) {
            $params["ContentFallback"] = $true
        }
        $params["SlowNetworkDeploymentMode"] = $SlowNetworkDeploymentMode

        Write-Host "Calling Add-CMScriptDeploymentType with parameters:"
        $params.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host ("  {0}: {1}" -f $_.Name, $_.Value) }

        Add-CMScriptDeploymentType @params | Out-Null
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Created MECM Application + DT successfully."
    }
    finally {
        # Always restore out of CM PSDrive so later file ops can’t break
        Set-Location $orig -ErrorAction SilentlyContinue
        Write-Host "Restored location after MECM work to: $orig"
    }
}

# ----------------------------
# MAIN
# ----------------------------
try {
    Write-Host "============================================================"
    Write-Host " Wireshark Auto-Packager (EXE)"
    Write-Host "============================================================"

    if ($GetLatestVersionOnly) {
        $v = Get-LatestWiresharkVersion -Quiet
        Write-Output $v
        return
    }


    if (-not (Test-IsAdmin)) { throw "Run PowerShell as Administrator." }

    if (-not (Test-Path $BaseDownloadRoot)) {
        $o = Push-FSLocation -Why "Create base download root"
        New-Item -ItemType Directory -Path $BaseDownloadRoot -Force | Out-Null
        Write-Host "Created: $BaseDownloadRoot"
        Pop-Location -OriginalLocation $o -Why "Create base download root"
    }

    if (-not (Test-NetworkShareAccess -Path $WiresharkRootNetworkPath)) {
        throw "Network root path not accessible/writable: $WiresharkRootNetworkPath"
    }

    $LatestVersion = Get-LatestWiresharkVersion
    if (-not $LatestVersion) { throw "No version discovered." }

    $DisplayNamePrefix = "Wireshark"
    $InstallerFileName = "Wireshark-$LatestVersion-x64.exe"
    $DownloadUrl       = "$WiresharkWin64Root/$InstallerFileName"

    $LocalVersionFolder   = Join-Path $BaseDownloadRoot $LatestVersion
    $LocalInstallerPath   = Join-Path $LocalVersionFolder $InstallerFileName

    $NetworkVersionFolder = Join-Path $WiresharkRootNetworkPath $LatestVersion
    $NetworkInstallerPath = Join-Path $NetworkVersionFolder $InstallerFileName

    Write-Host "Resolved variables:"
    Write-Host "  Version             : $LatestVersion"
    Write-Host "  InstallerFileName   : $InstallerFileName"
    Write-Host "  DownloadUrl         : $DownloadUrl"
    Write-Host "  LocalVersionFolder  : $LocalVersionFolder"
    Write-Host "  NetworkVersionFolder: $NetworkVersionFolder"
    Write-Host "  ContentFallback     : $EnableContentFallback"
    Write-Host "  SlowNetworkMode     : $SlowNetworkDeploymentMode"
    Write-Host ""

    # --- Directory prep (FS context) ---
    $orig = Push-FSLocation -Why "Directory prep"
    if (-not (Test-Path $LocalVersionFolder)) {
        Write-Host "Creating local folder: $LocalVersionFolder"
        New-Item -ItemType Directory -Path $LocalVersionFolder -Force | Out-Null
    }
    if (-not (Test-Path $NetworkVersionFolder)) {
        Write-Host "Creating network folder: $NetworkVersionFolder"
        New-Item -ItemType Directory -Path $NetworkVersionFolder -Force | Out-Null
    }
    Pop-Location -OriginalLocation $orig -Why "Directory prep"

    # --- Download (FS context) ---
    $orig = Push-FSLocation -Why "Download"
    if (Test-Path $LocalInstallerPath) {
        Write-Host "Local installer exists; skipping download: $LocalInstallerPath"
    }
    else {
        Write-Host "Downloading Wireshark..."
        Write-Host "  From: $DownloadUrl"
        Write-Host "  To  : $LocalInstallerPath"
        curl.exe -L --fail --silent --show-error -o $LocalInstallerPath $DownloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $DownloadUrl" }
        Write-Host "Download complete."
    }
    Pop-Location -OriginalLocation $orig -Why "Download"

    # --- Copy to share (FS context) ---
    $orig = Push-FSLocation -Why "Copy to share"
    if (Test-Path $NetworkInstallerPath) {
        Write-Host "Network installer exists; skipping copy: $NetworkInstallerPath"
    }
    else {
        Write-Host "Copying to network share..."
        Write-Host "  From: $LocalInstallerPath"
        Write-Host "  To  : $NetworkInstallerPath"
        Copy-Item -Path $LocalInstallerPath -Destination $NetworkInstallerPath -Force -ErrorAction Stop
        Write-Host "Copy complete."
    }
    Pop-Location -OriginalLocation $orig -Why "Copy to share"

    # --- Batch wrappers (FS context) ---
    Create-BatchFiles -NetworkPath $NetworkVersionFolder -InstallerFileName $InstallerFileName -DisplayNamePrefix $DisplayNamePrefix

    # --- Temp install + registry crawl (FS context) ---
    $registryData = Install-And-Extract-RegistryData -InstallerPath $NetworkInstallerPath -DisplayNamePrefix $DisplayNamePrefix

    # --- Build final app metadata ---
    $FinalAppName = $registryData.DisplayName
    if (-not $FinalAppName) { $FinalAppName = "Wireshark $LatestVersion (x64)" }

    $FinalPublisher = $registryData.Publisher
    if (-not $FinalPublisher) { $FinalPublisher = "Wireshark Foundation" }

    # THIS is the packaged version you want detection to match:
    $PackagedVersion = $LatestVersion

    Write-Host ""
    Write-Host "Final packaging metadata:"
    Write-Host "  AppName           : $FinalAppName"
    Write-Host "  Publisher         : $FinalPublisher"
    Write-Host "  Packaged Version  : $PackagedVersion"
    Write-Host "  Uninstall Root    : $($registryData.UninstallRoot)"
    Write-Host "  Uninstall KeyName : $($registryData.PSChildName)"
    Write-Host "  DisplayVersion (found): $($registryData.DisplayVersion)"
    Write-Host ""

    # If Wireshark writes a different DisplayVersion than the download version, you WANT TO KNOW.
    if ($registryData.DisplayVersion -and ($registryData.DisplayVersion -ne $PackagedVersion)) {
        Write-Warning "DisplayVersion found '$($registryData.DisplayVersion)' does not equal packaged version '$PackagedVersion'."
        Write-Warning "Detection will still be set to packaged version per your request. If installs report different DisplayVersion, detection will fail."
    }

    # --- MECM App + DT (CM context only inside function) ---
    New-WiresharkMECMApplication `
        -AppName $FinalAppName `
        -Version $PackagedVersion `
        -Publisher $FinalPublisher `
        -ContentLocation $NetworkVersionFolder `
        -UninstallRootPSPath $registryData.UninstallRoot `
        -UninstallKeyName $registryData.PSChildName

    Write-Host "============================================================"
    Write-Host " DONE"
    Write-Host "============================================================"
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}