<#
Vendor: Microsoft
App: .NET 10 Desktop Runtime (x64)
CMName: Microsoft Windows Desktop Runtime - 10

.SYNOPSIS
    Downloads the latest .NET 10 Desktop Runtime (x64) and creates an MECM application.

.DESCRIPTION
    - Retrieves the latest .NET 10 Desktop Runtime version from releases-index.json
    - Downloads windowsdesktop-runtime-<version>-win-x64.exe
    - Stages content to a versioned network folder
    - Creates install.bat / uninstall.bat
    - Creates MECM Application + Script Deployment Type
    - Sets DT content options:
      - Allow clients to use fallback source location for content
      - Download content from DP and run locally
    - Detection: hostfxr.dll presence in %ProgramFiles%\dotnet\host\fxr\<version>

.PARAMETER GetLatestVersionOnly
    Outputs only the latest version string and exits.
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Configuration ---
$DotnetReleasesJsonUrl = "https://dotnetcli.blob.core.windows.net/dotnet/release-metadata/releases-index.json"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager"
$DesktopRuntimeRootNetworkPath = Join-Path $FileServerPath "Applications\Microsoft\.NET Core"
$DownloadBaseUrl = "https://dotnetcli.azureedge.net/dotnet/"

$Publisher = "Microsoft Corporation"

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
        Set-Location "${SiteCode}:" -ErrorAction Stop
        Write-Host "Connected to CM site: $SiteCode"
        return $true
    }
    catch {
        Write-Error "Failed to connect to CM site: $($_.Exception.Message)"
        return $false
    }
}

function Get-LatestDotnet10DesktopRuntimeVersion {
    param([switch]$Quiet)

    if (-not $Quiet) { Write-Host "Fetching .NET release information from: $DotnetReleasesJsonUrl" }
    try {
        $json = (curl.exe -L --fail --silent --show-error $DotnetReleasesJsonUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch .NET release info: $DotnetReleasesJsonUrl" }

        $releases = ConvertFrom-Json $json

        $idx = $releases.'releases-index' |
            Where-Object { $_.'channel-version' -eq '10.0' } |
            Select-Object -First 1

        if (-not $idx) {
            throw "Channel-version '10.0' not found in releases-index."
        }

        $v = $idx.'latest-runtime'
        if ([string]::IsNullOrWhiteSpace($v)) {
            throw "latest-runtime not present for channel 10.0."
        }

        Write-Host "Latest .NET 10 runtime version: $v"
        return $v
    }
    catch {
        Write-Error "Failed to determine latest .NET 10 version: $($_.Exception.Message)"
        return $null
    }
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)

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

function Create-BatchFiles {
    param(
        [Parameter(Mandatory)][string]$NetworkPath,
        [Parameter(Mandatory)][string]$Version
    )

    $install = @"
@echo off
start /wait "" "%~dp0windowsdesktop-runtime-${Version}-win-x64.exe" /install /quiet /norestart
exit /b 0
"@

    $uninstall = @"
@echo off
start /wait "" "%~dp0windowsdesktop-runtime-${Version}-win-x64.exe" /uninstall /quiet /norestart
exit /b 0
"@

    Set-Content -LiteralPath (Join-Path $NetworkPath "install.bat") -Value $install -Encoding ASCII -ErrorAction Stop
    Set-Content -LiteralPath (Join-Path $NetworkPath "uninstall.bat") -Value $uninstall -Encoding ASCII -ErrorAction Stop
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

function New-MECMApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$RuntimeVersion
    )

    $orig = Get-Location
    try {
        if (-not (Connect-CMSite -SiteCode $SiteCode)) {
            throw "CM site connection failed."
        }
        $existing = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Warning "Application already exists: $AppName"
            return
        }

        Write-Host "Creating CM Application: $AppName"
        $cmApp = New-CMApplication -Name $AppName -Publisher $Publisher -SoftwareVersion $SoftwareVersion -Description $Comment -ErrorAction Stop

        $clause = New-CMDetectionClauseFile `
            -Path ("$env:ProgramFiles\dotnet\host\fxr\{0}" -f $RuntimeVersion) `
            -FileName "hostfxr.dll" `
            -Existence `
            -Is64Bit

        $dtParams = @{
            ApplicationName           = $AppName
            DeploymentTypeName        = "Script Installer"
            InstallCommand            = "install.bat"
            UninstallCommand          = "uninstall.bat"
            ContentLocation           = $ContentPath
            InstallationBehaviorType  = "InstallForSystem"
            LogonRequirementType      = "WhetherOrNotUserLoggedOn"
            EstimatedRuntimeMins      = $EstimatedRuntimeMins
            MaximumRuntimeMins        = $MaximumRuntimeMins
            ContentFallback           = $true
            SlowNetworkDeploymentMode = "Download"
            AddDetectionClause        = @($clause)
            ErrorAction               = "Stop"
        }

        Write-Host "Adding Script Deployment Type: Script Installer"
        Add-CMScriptDeploymentType @dtParams
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Created MECM application: $AppName"
    }
    finally {
        Set-Location $orig -ErrorAction SilentlyContinue
    }
}

# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    try {
        $v = Get-LatestDotnet10DesktopRuntimeVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
        exit 0
    }
    catch { exit 1 }
}

# --- Main ---
try {
    $startLocation = Get-Location

    $runAsUser = "{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME
    $machine   = $env:COMPUTERNAME

    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host ".NET 10 Desktop Runtime (x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host "RunAsUser                    : $runAsUser"
    Write-Host "Machine                      : $machine"
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host "DesktopRuntimeRootNetworkPath: $DesktopRuntimeRootNetworkPath"
    Write-Host "ReleasesIndexUrl             : $DotnetReleasesJsonUrl"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    $LatestVersion = Get-LatestDotnet10DesktopRuntimeVersion
    if (-not $LatestVersion) {
        throw "Could not determine latest version."
    }

    $AppName = "Microsoft Windows Desktop Runtime - $LatestVersion (x64)"
    $SoftwareVersion = $LatestVersion

    $NetworkPath = Join-Path $DesktopRuntimeRootNetworkPath $LatestVersion

    Ensure-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $DesktopRuntimeRootNetworkPath)) {
        throw "Network root path not accessible: $DesktopRuntimeRootNetworkPath"
    }

    Ensure-Folder -Path $NetworkPath

    $FileName    = "windowsdesktop-runtime-$LatestVersion-win-x64.exe"
    $DownloadUrl = "${DownloadBaseUrl}WindowsDesktop/$LatestVersion/$FileName"

    $LocalFile = Join-Path $BaseDownloadRoot $FileName
    $NetFile   = Join-Path $NetworkPath $FileName

    Write-Host "LatestVersion                : $LatestVersion"
    Write-Host "AppName                      : $AppName"
    Write-Host "SoftwareVersion              : $SoftwareVersion"
    Write-Host "DownloadUrl                  : $DownloadUrl"
    Write-Host "LocalFile                    : $LocalFile"
    Write-Host "NetworkPath                  : $NetworkPath"
    Write-Host ""

    if (-not (Test-Path -LiteralPath $LocalFile)) {
        Write-Host "Downloading installer..."
        curl.exe -L --fail --silent --show-error -o $LocalFile $DownloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $DownloadUrl" }
    }
    else {
        Write-Host "Local installer exists. Skipping download."
    }

    if (-not (Test-Path -LiteralPath $NetFile)) {
        Write-Host "Copying installer to network..."
        Copy-Item -LiteralPath $LocalFile -Destination $NetFile -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network installer exists. Skipping copy."
    }

    Create-BatchFiles -NetworkPath $NetworkPath -Version $LatestVersion

    New-MECMApplication -AppName $AppName -SoftwareVersion $SoftwareVersion -ContentPath $NetworkPath -RuntimeVersion $LatestVersion

    Write-Host ""
    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
