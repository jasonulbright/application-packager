<#
Vendor: Cisco
App: Cisco Webex (x64)
CMName: Cisco Webex

.SYNOPSIS
    Automates downloading the latest Cisco Webex x64 MSI and creating an MECM application.

.DESCRIPTION
    Downloads the Cisco Webex MSI from Cisco's distribution server to a local staging folder,
    extracts ProductVersion and ProductCode from the MSI via Windows Installer COM, then creates
    an MECM application with Windows Installer (ProductCode) detection.
    Each version creates a separate MECM application with its own versioned content folder.

    Running with -GetLatestVersionOnly scrapes the Webex release notes page for the latest
    version number — no MSI download required. The main run downloads the MSI to staging,
    extracts the actual ProductVersion and ProductCode, and reuses the staged copy if present.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Work order or comment string applied to the MECM application description.

.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").

.PARAMETER GetLatestVersionOnly
    Scrapes the Webex release notes page for the latest version string and exits.
    Does not download the MSI.

.NOTES
    Requires:
      - PowerShell 5.1
      - ConfigMgr Admin Console installed (for ConfigurationManager.psd1)
      - RBAC rights to create Applications and Deployment Types

.KNOWN ISSUES
    The release notes page (help.webex.com/en-us/article/mqkve8) reports the highest version
    across all channels. The Gold MSI may trail by a build if Cisco hasn't promoted yet —
    the version packaged in MECM is always the authoritative ProductVersion from the MSI itself.
    The download URL is sourced from help.webex.com/en-us/article/nw5p67g. If downloads fail,
    verify the URL against that page. An English-only variant (Webex_en.msi) is at the same path.
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"

# --- Configuration ---
$StagingFolder    = Join-Path $env:USERPROFILE "Downloads\Webex_staging"
$StagingMsi       = Join-Path $StagingFolder "Webex.msi"
$WebexNetworkRoot = Join-Path $FileServerPath "Applications\Cisco\Webex"
$WebexDownloadUrl = "https://binaries.webex.com/WebexOfclDesktop-Win-64-Gold/Webex.msi"
$ReleaseNotesUrl  = "https://help.webex.com/en-us/article/mqkve8/Webex-App-%7C-Release-notes"
$Publisher        = "Cisco Systems, Inc."
$MsiFileName      = "Webex.msi"

# --- Functions ---

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

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)
    $originalLocation = Get-Location
    try {
        if (-not (Test-Path -LiteralPath $Path -ErrorAction Stop)) {
            Write-Error "Network path '$Path' does not exist or is inaccessible."
            return $false
        }
        $testFile = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
        Set-Content -LiteralPath $testFile -Value "Test" -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $testFile -ErrorAction Stop
        return $true
    }
    catch {
        Write-Error "Failed to access network share '$Path': $($_.Exception.Message)"
        return $false
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
    }
}

function Get-MsiProperty {
    param(
        [Parameter(Mandatory)][string]$MsiPath,
        [Parameter(Mandatory)][string]$Property
    )
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database  = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($MsiPath, 0))
        $view      = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @("SELECT Value FROM Property WHERE Property = '$Property'"))
        $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
        $record    = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        if (-not $record) { return $null }
        return $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
    }
    catch {
        Write-Error "Failed to read MSI property '$Property': $($_.Exception.Message)"
        return $null
    }
}

function Invoke-WebexMsiDownload {
    param([switch]$Force)

    if (-not (Test-Path -LiteralPath $StagingFolder)) {
        New-Item -ItemType Directory -Path $StagingFolder -Force | Out-Null
    }

    if ($Force -or -not (Test-Path -LiteralPath $StagingMsi)) {
        Write-Host "Downloading Webex MSI: $WebexDownloadUrl"
        curl.exe -L --fail --silent --show-error -o $StagingMsi $WebexDownloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $WebexDownloadUrl" }
        Write-Host "Downloaded: $StagingMsi"
    }
    else {
        Write-Host "Using cached staging MSI: $StagingMsi"
    }
}

function Get-LatestWebexVersion {
    param([switch]$Quiet)
    try {
        $html     = (curl.exe -L --fail --silent --show-error $ReleaseNotesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch release notes: $ReleaseNotesUrl" }
        $versions = [regex]::Matches($html, '\b(\d+\.\d+\.\d+\.\d+)\b') |
                    ForEach-Object { $_.Value } |
                    Select-Object -Unique |
                    Sort-Object { [version]$_ } -Descending
        $latest = $versions | Select-Object -First 1
        if ([string]::IsNullOrWhiteSpace($latest)) { throw "No version numbers found on release notes page." }
        if (-not $Quiet) { Write-Host "Latest Webex version: $latest" }
        return $latest
    }
    catch {
        Write-Error "Failed to retrieve latest Webex version: $($_.Exception.Message)"
        return $null
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

function New-MECMWebexApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$NetworkPath,
        [Parameter(Mandatory)][string]$ProductCode
    )
    $originalLocation = Get-Location
    try {
        if (-not (Connect-CMSite -SiteCode $SiteCode)) {
            Write-Error "Failed to connect to CM site."
            return
        }

        $existingApp = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existingApp) {
            $dts = Get-CMDeploymentType -ApplicationName $AppName -ErrorAction SilentlyContinue
            if ($dts -and $dts.Count -gt 0) {
                Write-Warning "Application '$AppName' already exists with $($dts.Count) deployment type(s). Skipping."
                return
            }
            $cmApp = $existingApp
        }
        else {
            Write-Host "Creating application: $AppName"
            $cmApp = New-CMApplication `
                -Name $AppName `
                -Publisher $Publisher `
                -SoftwareVersion $Version `
                -LocalizedApplicationName $AppName `
                -Description $Comment `
                -ErrorAction Stop
        }

        $detectionClause = New-CMDetectionClauseWindowsInstaller `
            -ProductCode $ProductCode `
            -Value `
            -PropertyType Version `
            -ExpressionOperator GreaterEquals `
            -ExpectedValue $Version

        $params = @{
            ApplicationName           = $AppName
            DeploymentTypeName        = "$AppName Script DT"
            InstallCommand            = "install.bat"
            UninstallCommand          = "uninstall.bat"
            ContentLocation           = $NetworkPath
            InstallationBehaviorType  = "InstallForSystem"
            LogonRequirementType      = "WhetherOrNotUserLoggedOn"
            MaximumRuntimeMins        = 30
            EstimatedRuntimeMins      = 10
            ContentFallback           = $true
            SlowNetworkDeploymentMode = "Download"
            AddDetectionClause        = $detectionClause
            ErrorAction               = "Stop"
        }

        Add-CMScriptDeploymentType @params | Out-Null
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Created MECM application: $AppName"
    }
    catch {
        Write-Error "Failed to create MECM application: $($_.Exception.Message)"
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
    }
}

# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    $v = Get-LatestWebexVersion -Quiet
    if (-not $v) { exit 1 }
    Write-Output $v
    exit 0
}

# --- Main ---
try {
    if (-not (Test-IsAdmin)) {
        Write-Error "This script must be run as Administrator."
        exit 1
    }

    Set-Location $PSScriptRoot -ErrorAction Stop
    Write-Host "Set initial location to script directory: $PSScriptRoot"

    Invoke-WebexMsiDownload

    $Version     = Get-MsiProperty -MsiPath $StagingMsi -Property 'ProductVersion'
    $ProductCode = Get-MsiProperty -MsiPath $StagingMsi -Property 'ProductCode'

    if (-not $Version) {
        Write-Error "Could not extract ProductVersion from MSI. Exiting."
        exit 1
    }
    if (-not $ProductCode) {
        Write-Error "Could not extract ProductCode from MSI. Exiting."
        exit 1
    }

    Write-Host "Webex version:  $Version"
    Write-Host "ProductCode:    $ProductCode"

    $AppName            = "Cisco Webex $Version"
    $NetworkVersionPath = Join-Path $WebexNetworkRoot $Version
    $NetworkMsi         = Join-Path $NetworkVersionPath $MsiFileName

    if (-not (Test-NetworkShareAccess -Path $WebexNetworkRoot)) {
        Write-Error "Network share '$WebexNetworkRoot' is inaccessible. Exiting."
        exit 1
    }

    if (-not (Test-Path -LiteralPath $NetworkVersionPath)) {
        Write-Host "Creating network directory: $NetworkVersionPath"
        New-Item -ItemType Directory -Path $NetworkVersionPath -Force -ErrorAction Stop | Out-Null
    }

    if (-not (Test-Path -LiteralPath $NetworkMsi)) {
        Write-Host "Copying MSI to network share..."
        Copy-Item -LiteralPath $StagingMsi -Destination $NetworkVersionPath -Force -ErrorAction Stop
        Write-Host "Copied to: $NetworkVersionPath"
    }
    else {
        Write-Host "MSI already exists on network share: $NetworkMsi"
    }

    $installBat   = Join-Path $NetworkVersionPath "install.bat"
    $uninstallBat = Join-Path $NetworkVersionPath "uninstall.bat"
    Set-Content -LiteralPath $installBat   -Value "start /wait msiexec.exe /i `"%~dp0$MsiFileName`" /qn /norestart" -Encoding ASCII
    Set-Content -LiteralPath $uninstallBat -Value "start /wait msiexec.exe /x `"%~dp0$MsiFileName`" /qn /norestart" -Encoding ASCII
    Write-Host "Created install.bat and uninstall.bat in $NetworkVersionPath"

    New-MECMWebexApplication `
        -AppName     $AppName `
        -Version     $Version `
        -NetworkPath $NetworkVersionPath `
        -ProductCode $ProductCode

    Write-Host ""
    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
