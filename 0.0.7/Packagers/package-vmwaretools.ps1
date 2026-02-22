<#
Vendor: Broadcom
App: VMware Tools (x64)
CMName: VMWare Tools

.SYNOPSIS
    Automates downloading the latest VMware Tools installer and creating an MECM application.

.DESCRIPTION
    - Parses https://packages.vmware.com/tools/releases/ to determine the latest version folder.
    - Downloads the latest Windows x64 installer from the "latest" directory.
    - Stages content to a versioned network folder.
    - Creates install.bat/uninstall.bat wrappers plus install.ps1/uninstall.ps1.
    - Creates an MECM Application + Script Deployment Type.
    - Deployment Type content options:
      - Allow clients to use fallback source location for content
      - Download content from DP and run locally
    - Detection: vmtoolsd.exe file version >= packaged version.

.PARAMETER SiteCode
    ConfigMgr site code for the CM PSDrive (e.g., "MCM").
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
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"

# --- Configuration Variables ---
$VmwareToolsReleasesUrl     = "https://packages.vmware.com/tools/releases/"
$VmwareToolsLatestUrl       = "https://packages.vmware.com/tools/releases/latest/windows/x64/"

$BaseDownloadRoot           = Join-Path $env:USERPROFILE "Downloads"
$VmwareToolsRootNetworkPath = Join-Path $FileServerPath "Applications\Broadcom"

$Publisher                  = "Broadcom"
$EstimatedRuntimeMins       = 10
$MaximumRuntimeMins         = 20

# --- Functions ---
function Test-IsAdmin {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Import-ConfigurationManagerModule {
    if (Get-Module -Name ConfigurationManager -ErrorAction SilentlyContinue) {
        return $true
    }

    $defaultModulePath = "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
    $modulePath = $null

    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($path in $uninstallPaths) {
        $subkeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName
        foreach ($subkey in $subkeys) {
            $keyPath = "$path\$subkey"
            $displayName = (Get-ItemProperty -Path $keyPath -Name "DisplayName" -ErrorAction SilentlyContinue).DisplayName
            if ($displayName -like "*Configuration Manager Console*") {
                $installLocation = (Get-ItemProperty -Path $keyPath -Name "InstallLocation" -ErrorAction SilentlyContinue).InstallLocation
                if ($installLocation) {
                    $potentialModulePath = Join-Path $installLocation "bin\ConfigurationManager.psd1"
                    if (Test-Path -LiteralPath $potentialModulePath) {
                        $modulePath = $potentialModulePath
                        break
                    }
                }
            }
        }
        if ($modulePath) { break }
    }

    if (-not $modulePath) {
        $modulePath = $defaultModulePath
    }

    if (-not (Test-Path -LiteralPath $modulePath)) {
        Write-Error "Configuration Manager module not found. Ensure CM Admin Console is installed."
        return $false
    }

    Import-Module $modulePath -Force -ErrorAction Stop
    return $true
}

function Connect-CMSite {
    param([Parameter(Mandatory)][string]$SiteCode)

    if (-not (Import-ConfigurationManagerModule)) {
        return $false
    }

    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        Write-Error "CM site PSDrive '$SiteCode' not found. Verify site code and connectivity."
        return $false
    }

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

function Get-LatestVmwareToolsVersion {
    param(
        [switch]$Quiet
    )

    if (-not $Quiet) { Write-Host "Fetching VMware Tools release information from: $VmwareToolsReleasesUrl" }

    $releasesHtml = (curl.exe -L --fail --silent --show-error $VmwareToolsReleasesUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch VMware Tools releases: $VmwareToolsReleasesUrl" }
    $versionFolders = [regex]::Matches($releasesHtml, 'href="(\d+\.\d+\.\d+/)"') |
        ForEach-Object { $_.Groups[1].Value.Trim('/') } |
        Select-Object -Unique

    if (-not $versionFolders) {
        throw "Could not find any version folders in the releases directory."
    }

    $latestVersion = $versionFolders | Sort-Object { [version]$_ } -Descending | Select-Object -First 1
    if (-not $latestVersion) {
        throw "Could not determine the latest version."
    }

    if (-not $Quiet) { Write-Host "Found latest VMware Tools version: $latestVersion" }
    return $latestVersion
}

function Get-LatestVmwareToolsInstallerFileName {
    Write-Host "Fetching latest Windows x64 directory listing from: $VmwareToolsLatestUrl"

    $installerHtml = (curl.exe -L --fail --silent --show-error $VmwareToolsLatestUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch VMware Tools latest listing: $VmwareToolsLatestUrl" }
    $fileName = ([regex]::Matches($installerHtml, 'href="(VMware-tools-[^"]*-x64\.exe)"') |
        Select-Object -First 1).Groups[1].Value

    if (-not $fileName) {
        throw "Could not find the installer filename in the latest directory."
    }

    return $fileName
}

function Create-BatchAndPS1Files {
    param(
        [Parameter(Mandatory)][string]$NetworkPath,
        [Parameter(Mandatory)][string]$FileName
    )

    $originalLocation = Get-Location
    try {
        Set-Location $PSScriptRoot -ErrorAction Stop

        $installArgs   = '/S /v "/qn REBOOT=R ADDLOCAL=ALL REMOVE=FileIntrospection,NetworkIntrospection"'
        $uninstallArgs = '/S /v "/qn REBOOT=R REMOVE=ALL"'

        $InstallBatContent = @"
@echo off
setlocal
start /wait "" "%~dp0$FileName" $installArgs
timeout /t 300 /nobreak >nul
exit /b 3010
"@

        $UninstallBatContent = @"
@echo off
setlocal
start /wait "" "%~dp0$FileName" $uninstallArgs
exit /b 0
"@

        $InstallPs1Content = @"
Start-Process `"$PSScriptRoot\$FileName`" -ArgumentList '$installArgs' -Wait
"@

        $UninstallPs1Content = @"
Start-Process `"$PSScriptRoot\$FileName`" -ArgumentList '$uninstallArgs' -Wait
"@

        Set-Content -LiteralPath (Join-Path $NetworkPath "install.bat")   -Value $InstallBatContent   -Encoding ASCII -ErrorAction Stop
        Set-Content -LiteralPath (Join-Path $NetworkPath "uninstall.bat") -Value $UninstallBatContent -Encoding ASCII -ErrorAction Stop
        Set-Content -LiteralPath (Join-Path $NetworkPath "install.ps1")   -Value $InstallPs1Content   -Encoding ASCII -ErrorAction Stop
        Set-Content -LiteralPath (Join-Path $NetworkPath "uninstall.ps1") -Value $UninstallPs1Content -Encoding ASCII -ErrorAction Stop
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
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

function New-MECMApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$NetworkPath,
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][string]$Publisher
    )

    $fsLocationBeforeCm = Get-Location
    try {
        if (-not (Test-IsAdmin)) {
            throw "Script must be run with admin privileges."
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) {
            throw "Failed to connect to CM site."
        }
        $existingApp = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existingApp) {
            $deploymentTypes = Get-CMDeploymentType -ApplicationName $AppName -ErrorAction SilentlyContinue
            if ($deploymentTypes -and $deploymentTypes.Count -gt 0) {
                Write-Warning "Application '$AppName' already exists with $($deploymentTypes.Count) deployment type(s). Skipping creation."
                return
            }
            Write-Warning "Application '$AppName' already exists. Skipping creation."
            return
        }

        $cmApp = New-CMApplication -Name $AppName -Publisher $Publisher -SoftwareVersion $Version -LocalizedApplicationName $AppName -Description $Comment -ErrorAction Stop

        Create-BatchAndPS1Files -NetworkPath $NetworkPath -FileName $FileName

        $detectionClause = New-CMDetectionClauseFile `
            -Path "$env:ProgramFiles\VMware\VMware Tools" `
            -FileName "vmtoolsd.exe" `
            -Value `
            -PropertyType Version `
            -ExpressionOperator GreaterEquals `
            -ExpectedValue $Version

        $deploymentTypeParams = @{
            ApplicationName          = $AppName
            DeploymentTypeName       = $AppName
            InstallCommand           = "install.bat"
            UninstallCommand         = "uninstall.bat"
            ContentLocation          = $NetworkPath
            InstallationBehaviorType = "InstallForSystem"
            LogonRequirementType     = "WhetherOrNotUserLoggedOn"
            EstimatedRuntimeMins     = $EstimatedRuntimeMins
            MaximumRuntimeMins       = $MaximumRuntimeMins
            AddDetectionClause       = @($detectionClause)
            ContentFallback          = $true
            SlowNetworkDeploymentMode = "Download"
            ErrorAction              = "Stop"
        }

        Add-CMScriptDeploymentType @deploymentTypeParams | Out-Null
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Created MECM application: $AppName"
    }
    finally {
        Set-Location $fsLocationBeforeCm -ErrorAction SilentlyContinue
    }
}

# --- Main Script ---
try {
    if (-not (Test-IsAdmin)) {
        Write-Error "This script must be run with admin privileges. Please run PowerShell as Administrator."
        exit 1
    }

    Set-Location $PSScriptRoot -ErrorAction Stop

    if ($GetLatestVersionOnly) {
        $Version = Get-LatestVmwareToolsVersion -Quiet
        Write-Output $Version
        return
    }

    $runAsUser = "{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME
    $machine   = $env:COMPUTERNAME

    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "VMware Tools Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host "RunAsUser                   : $runAsUser"
    Write-Host "Machine                     : $machine"
    Write-Host "SiteCode                    : $SiteCode"
    Write-Host "BaseDownloadRoot            : $BaseDownloadRoot"
    Write-Host "VmwareToolsRootNetworkPath  : $VmwareToolsRootNetworkPath"
    Write-Host "VmwareToolsReleasesUrl      : $VmwareToolsReleasesUrl"
    Write-Host "VmwareToolsLatestUrl        : $VmwareToolsLatestUrl"
    Write-Host ""

    $Version  = Get-LatestVmwareToolsVersion
    $FileName = Get-LatestVmwareToolsInstallerFileName

    $AppName     = "VMWare Tools $Version"
    $NetworkPath = Join-Path $VmwareToolsRootNetworkPath $AppName

    if (-not (Test-NetworkShareAccess -Path $VmwareToolsRootNetworkPath)) {
        Write-Error "Network share '$VmwareToolsRootNetworkPath' is inaccessible. Skipping '$AppName'."
        exit 1
    }

    if (-not (Test-Path -LiteralPath $NetworkPath)) {
        New-Item -ItemType Directory -Path $NetworkPath -Force -ErrorAction Stop | Out-Null
    }

    $DownloadFolderName = "VMware_Tools_${Version}_installers"
    $DownloadPath = Join-Path $BaseDownloadRoot $DownloadFolderName
    if (-not (Test-Path -LiteralPath $DownloadPath)) {
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
    }

    $OutputPath      = Join-Path $DownloadPath $FileName
    $NetworkFilePath = Join-Path $NetworkPath $FileName

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        $dlUrl = "{0}{1}" -f $VmwareToolsLatestUrl, $FileName
        curl.exe -L --fail --silent --show-error -o $OutputPath $dlUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $dlUrl" }
    }

    if (-not (Test-Path -LiteralPath $NetworkFilePath)) {
        Copy-Item -LiteralPath $OutputPath -Destination $NetworkPath -Force -ErrorAction Stop
    }

    Create-BatchAndPS1Files -NetworkPath $NetworkPath -FileName $FileName

    New-MECMApplication -AppName $AppName -Version $Version -NetworkPath $NetworkPath -FileName $FileName -Publisher $Publisher

    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
