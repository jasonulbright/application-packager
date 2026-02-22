<#
Vendor: Wireshark Foundation
App: Wireshark (x64)
CMName: Wireshark

.SYNOPSIS
    Packages Wireshark (x64) for MECM.

.DESCRIPTION
    Downloads the latest Wireshark x64 installer from the official download
    server, stages content to a versioned network location, temporarily
    installs the product to extract registry metadata, and creates an MECM
    Application with registry-based version detection.
    Detection uses DisplayVersion string equals on the discovered uninstall
    registry key.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Wireshark Foundation\Wireshark\<Version>

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Wireshark version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Local administrator
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Configuration ---
$WiresharkDownloadPage = "https://www.wireshark.org/download.html"
$WiresharkWin64Root    = "https://www.wireshark.org/download/win64"

$VendorFolder = "Wireshark Foundation"
$AppFolder    = "Wireshark"

$DisplayNamePrefix      = "Wireshark"
$DesktopIconSetting     = "no"
$QuickLaunchIconSetting = "no"

$InitialInstallBufferSeconds = 60
$PollSleepSeconds            = 60
$MaxRegistryPollRetries      = 20

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\Wireshark"

$EstimatedRuntimeMins = 10
$MaximumRuntimeMins   = 30

# --- Functions ---
function Test-IsAdmin {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Warning "Admin check failed: $($_.Exception.Message)"
        return $false
    }
}

function Connect-CMSite {
    param([Parameter(Mandatory)][string]$SiteCode)

    try {
        if (-not (Get-Module -Name ConfigurationManager -ErrorAction SilentlyContinue)) {
            $cmModulePath = Join-Path $env:SMS_ADMIN_UI_PATH "..\ConfigurationManager.psd1"
            if (Test-Path -LiteralPath $cmModulePath) {
                Import-Module $cmModulePath -ErrorAction Stop
            }
            else {
                Import-Module ConfigurationManager -ErrorAction Stop
            }
        }

        if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
            throw "Configuration Manager PSDrive '$SiteCode' is not available."
        }

        Set-Location "${SiteCode}:" -ErrorAction Stop
        Write-Host "Connected to CM site: $SiteCode"
        return $true
    }
    catch {
        Write-Error "Failed to connect to CM site: $($_.Exception.Message)"
        return $false
    }
}

function Initialize-Folder {
    param([Parameter(Mandatory)][string]$Path)

    $origLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)

    $origLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
            Write-Error "Network path does not exist or is inaccessible: $Path"
            return $false
        }

        try {
            $tmp = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
            Set-Content -LiteralPath $tmp -Value "test" -Encoding ASCII -ErrorAction Stop
            Remove-Item -LiteralPath $tmp -ErrorAction Stop
            return $true
        }
        catch {
            Write-Error "Network share is not writable: $Path ($($_.Exception.Message))"
            return $false
        }
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

function Get-LatestWiresharkVersion {
    param([switch]$Quiet)

    if (-not $Quiet) {
        Write-Host "Wireshark download page      : $WiresharkDownloadPage"
    }

    try {
        $html = (curl.exe -L --fail --silent --show-error $WiresharkDownloadPage) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Wireshark download page: $WiresharkDownloadPage" }

        if ($html -match 'Stable Release:\s*([0-9]+\.[0-9]+\.[0-9]+)') {
            $v = $matches[1]
            if (-not $Quiet) {
                Write-Host "Latest Wireshark version     : $v"
            }
            return $v
        }

        throw "Could not parse Stable Release version from download page."
    }
    catch {
        Write-Error "Failed to get Wireshark version: $($_.Exception.Message)"
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
            $exe       = $cmd.Substring(1, $secondQuote - 1)
            $arguments = $cmd.Substring($secondQuote + 1).Trim()
            return @{ FilePath = $exe; Arguments = $arguments }
        }
    }

    $parts = $cmd.Split(@(' '), 2, [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($parts.Count -eq 1) { return @{ FilePath = $parts[0]; Arguments = "" } }
    return @{ FilePath = $parts[0]; Arguments = $parts[1] }
}

function Convert-RegRootToCMKeyName {
    param(
        [Parameter(Mandatory)][string]$UninstallRootPSPath,
        [Parameter(Mandatory)][string]$PSChildName
    )

    $cmBase = $UninstallRootPSPath -replace '^HKLM:\\', ''
    return "$cmBase\$PSChildName"
}

function Invoke-WiresharkMetadataExtraction {
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [Parameter(Mandatory)][string]$Prefix
    )

    Write-Host "Installing temporarily for metadata extraction..."
    $installArgs = "/S /desktopicon=$DesktopIconSetting /quicklaunchicon=$QuickLaunchIconSetting"
    Start-Process -FilePath $InstallerPath -ArgumentList $installArgs -Wait -NoNewWindow

    Write-Host "Waiting $InitialInstallBufferSeconds seconds for installer to complete registration..."
    Start-Sleep -Seconds $InitialInstallBufferSeconds

    $uninstallRoots = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $data    = $null
    $retry   = 0
    $pattern = "$Prefix*"

    do {
        $retry++
        Write-Host "Registry poll attempt $retry/$MaxRegistryPollRetries (pattern: '$pattern')"

        foreach ($root in $uninstallRoots) {
            $keys = Get-ChildItem -Path $root -ErrorAction SilentlyContinue
            if (-not $keys) { continue }

            foreach ($k in $keys) {
                $dn = (Get-ItemProperty -Path $k.PSPath -Name "DisplayName" -ErrorAction SilentlyContinue).DisplayName
                if ($dn -and ($dn -like $pattern)) {
                    $props = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
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
                    Write-Host "Found registry entry: $($data.DisplayName) ($($data.DisplayVersion))"
                    break
                }
            }

            if ($data) { break }
        }

        if (-not $data -and $retry -lt $MaxRegistryPollRetries) {
            Write-Host "No match yet. Sleeping $PollSleepSeconds seconds..."
            Start-Sleep -Seconds $PollSleepSeconds
        }

    } while (-not $data -and $retry -lt $MaxRegistryPollRetries)

    if (-not $data) {
        throw "No uninstall registry entry found for '$Prefix' after $MaxRegistryPollRetries polls."
    }

    # Uninstall to return packaging machine to clean state
    Write-Host "Uninstalling after metadata extraction..."

    $uninstallCmd = $null
    if ($data.QuietUninstallString) {
        $uninstallCmd = $data.QuietUninstallString
    }
    elseif ($data.UninstallString) {
        $uninstallCmd = $data.UninstallString
    }
    else {
        $fallback = Join-Path $env:ProgramFiles "Wireshark\uninstall.exe"
        if (Test-Path -LiteralPath $fallback) {
            $uninstallCmd = "`"$fallback`" /S"
        }
    }

    if ($uninstallCmd) {
        $parsed = Split-CommandLine -CommandLine $uninstallCmd
        if ($parsed -and $parsed.FilePath) {
            Start-Process -FilePath $parsed.FilePath -ArgumentList $parsed.Arguments -Wait -NoNewWindow
            Start-Sleep -Seconds 30
        }
        else {
            Write-Warning "Could not parse uninstall command: $uninstallCmd"
        }
    }
    else {
        Write-Warning "No uninstall command found. Machine may not be clean."
    }

    return $data
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
        if ($h.PSObject.Properties.Name -contains 'Revision') { $revs += [UInt32]$h.Revision; continue }
        if ($h.PSObject.Properties.Name -contains 'CIVersion') { $revs += [UInt32]$h.CIVersion; continue }
    }

    $revs = $revs | Sort-Object -Unique -Descending
    if ($revs.Count -le $KeepLatest) { return }

    foreach ($rev in ($revs | Select-Object -Skip $KeepLatest)) {
        Remove-CMApplicationRevisionHistory -Id $CI_ID -Revision $rev -Force -ErrorAction Stop
    }
}

function New-MECMWiresharkApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$InstallerFileName,
        [Parameter(Mandatory)][string]$Publisher,
        [Parameter(Mandatory)][string]$RegistryKeyName
    )

    $orig = Get-Location

    try {
        if (-not (Test-IsAdmin)) { throw "Run PowerShell as Administrator." }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $existing = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Warning "Application already exists: $AppName"
            return
        }

        Write-Host "Creating CM Application      : $AppName"
        $cmApp = New-CMApplication `
            -Name $AppName `
            -Publisher $Publisher `
            -SoftwareVersion $SoftwareVersion `
            -Description $Comment `
            -AutoInstall $true `
            -ErrorAction Stop

        Write-Host "Application CI_ID            : $($cmApp.CI_ID)"

        Set-Location C: -ErrorAction Stop

        $installBatPath   = Join-Path $ContentPath "install.bat"
        $uninstallBatPath = Join-Path $ContentPath "uninstall.bat"

        if (-not (Test-Path -LiteralPath $installBatPath)) {
            $installBat = @"
@echo off
setlocal
start /wait "" "%~dp0$InstallerFileName" /S /desktopicon=$DesktopIconSetting /quicklaunchicon=$QuickLaunchIconSetting
exit /b 0
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal EnableExtensions EnableDelayedExpansion

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
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        # Detection: DisplayVersion string equals packaged version
        $clause = New-CMDetectionClauseRegistryKeyValue `
            -Hive LocalMachine `
            -KeyName $RegistryKeyName `
            -ValueName "DisplayVersion" `
            -PropertyType String `
            -ExpectedValue $SoftwareVersion `
            -ExpressionOperator IsEquals `
            -Value

        Write-Host "Adding Script Deployment Type: $dtName"
        Add-CMScriptDeploymentType `
            -ApplicationName $AppName `
            -DeploymentTypeName $dtName `
            -ContentLocation $ContentPath `
            -InstallCommand "install.bat" `
            -UninstallCommand "uninstall.bat" `
            -InstallationBehaviorType InstallForSystem `
            -LogonRequirementType WhetherOrNotUserLoggedOn `
            -EstimatedRuntimeMins $EstimatedRuntimeMins `
            -MaximumRuntimeMins $MaximumRuntimeMins `
            -AddDetectionClause @($clause) `
            -ContentFallback `
            -SlowNetworkDeploymentMode Download `
            -ErrorAction Stop | Out-Null

        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Created MECM application     : $AppName"
    }
    finally {
        Set-Location $orig -ErrorAction SilentlyContinue
    }
}

function Get-WiresharkNetworkAppRoot {
    param([Parameter(Mandatory)][string]$FileServerPath)

    $appsRoot   = Join-Path $FileServerPath "Applications"
    $vendorPath = Join-Path $appsRoot $VendorFolder
    $appPath    = Join-Path $vendorPath $AppFolder

    Initialize-Folder -Path $appsRoot
    Initialize-Folder -Path $vendorPath
    Initialize-Folder -Path $appPath

    return $appPath
}

# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $v = Get-LatestWiresharkVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch {
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "Wireshark (x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "FileServerPath               : $FileServerPath"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host "WiresharkDownloadPage        : $WiresharkDownloadPage"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-WiresharkNetworkAppRoot -FileServerPath $FileServerPath

    $version = Get-LatestWiresharkVersion
    if (-not $version) {
        throw "Could not resolve Wireshark version."
    }

    $installerFileName = "Wireshark-${version}-x64.exe"
    $downloadUrl       = "${WiresharkWin64Root}/${installerFileName}"
    $localExe    = Join-Path $BaseDownloadRoot $installerFileName
    $contentPath = Join-Path $networkAppRoot $version

    Initialize-Folder -Path $contentPath

    $netExe = Join-Path $contentPath $installerFileName

    Write-Host "Version                      : $version"
    Write-Host "Local installer              : $localExe"
    Write-Host "ContentPath                  : $contentPath"
    Write-Host "Network installer            : $netExe"
    Write-Host ""

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Host "Downloading installer..."
        curl.exe -L --fail --silent --show-error -o $localExe $downloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $downloadUrl" }
    }
    else {
        Write-Host "Local installer exists. Skipping download."
    }

    if (-not (Test-Path -LiteralPath $netExe)) {
        Write-Host "Copying installer to network..."
        Copy-Item -LiteralPath $localExe -Destination $netExe -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network installer exists. Skipping copy."
    }

    # --- Temporary install for metadata extraction ---
    $registryData = Invoke-WiresharkMetadataExtraction `
        -InstallerPath $netExe `
        -Prefix $DisplayNamePrefix

    $appName   = $registryData.DisplayName
    $publisher = $registryData.Publisher

    if (-not $appName)   { $appName   = "Wireshark $version (x64)" }
    if (-not $publisher) { $publisher = "Wireshark Foundation" }

    if ($registryData.DisplayVersion -and ($registryData.DisplayVersion -ne $version)) {
        Write-Warning "Registry DisplayVersion '$($registryData.DisplayVersion)' differs from download version '$version'. Detection uses download version."
    }

    $registryKeyName = Convert-RegRootToCMKeyName `
        -UninstallRootPSPath $registryData.UninstallRoot `
        -PSChildName $registryData.PSChildName

    Write-Host ""
    Write-Host "CM Application Name          : $appName"
    Write-Host "CM SoftwareVersion           : $version"
    Write-Host "Detection registry key       : $registryKeyName"
    Write-Host ""

    New-MECMWiresharkApplication `
        -AppName $appName `
        -SoftwareVersion $version `
        -ContentPath $contentPath `
        -InstallerFileName $installerFileName `
        -Publisher $publisher `
        -RegistryKeyName $registryKeyName

    Write-Host ""
    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
