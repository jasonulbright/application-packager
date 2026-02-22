<#
Vendor: Microsoft Corporation
App: Microsoft Edge Enterprise (x64)
CMName: Microsoft Edge Enterprise

.SYNOPSIS
    Packages Microsoft Edge Enterprise (x64) MSI for MECM.

.DESCRIPTION
    Downloads the latest Microsoft Edge Enterprise x64 MSI from the official
    Microsoft link, stages content to a versioned network location, and creates
    an MECM Application with file-based version detection.
    Detection uses msedge.exe version in the primary install path OR the staged
    EdgeUpdate path (OR connector).

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\Chredge\<Version>

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Edge version string and exits.

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
$EdgeStableMSIUrl = "https://go.microsoft.com/fwlink/?LinkID=2093437"
$EdgeVersionUrl   = "https://edgeupdates.microsoft.com/api/products?view=enterprise"

$VendorFolder = "Microsoft"
$AppFolder    = "Chredge"
$MsiFileName  = "MicrosoftEdgeEnterpriseX64.msi"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\Edge"

$EstimatedRuntimeMins = 10
$MaximumRuntimeMins   = 20

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

function Get-LatestEdgeVersion {
    param([switch]$Quiet)

    if (-not $Quiet) {
        Write-Host "Edge version API             : $EdgeVersionUrl"
    }

    try {
        $json = (curl.exe -L --fail --silent --show-error $EdgeVersionUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Edge version info: $EdgeVersionUrl" }

        $response = ConvertFrom-Json $json
        $stableChannel = $response | Where-Object { $_.Product -eq "Stable" }
        $latestVersion = $stableChannel.Releases |
            Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } |
            Sort-Object { [version]$_.ProductVersion } -Descending |
            Select-Object -First 1 -ExpandProperty ProductVersion

        if (-not $latestVersion) {
            throw "Could not determine latest Microsoft Edge Stable version."
        }

        if (-not $Quiet) {
            Write-Host "Latest Edge version          : $latestVersion"
        }
        return $latestVersion
    }
    catch {
        Write-Error "Failed to get Edge version: $($_.Exception.Message)"
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
        if ($h.PSObject.Properties.Name -contains 'Revision') { $revs += [UInt32]$h.Revision; continue }
        if ($h.PSObject.Properties.Name -contains 'CIVersion') { $revs += [UInt32]$h.CIVersion; continue }
    }

    $revs = $revs | Sort-Object -Unique -Descending
    if ($revs.Count -le $KeepLatest) { return }

    foreach ($rev in ($revs | Select-Object -Skip $KeepLatest)) {
        Remove-CMApplicationRevisionHistory -Id $CI_ID -Revision $rev -Force -ErrorAction Stop
    }
}

function New-MECMEdgeApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$MsiFileName,
        [Parameter(Mandatory)][string]$Publisher
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
start /wait "" msiexec.exe /i "%~dp0$MsiFileName" /qn /norestart
exit /b 0
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal
start /wait "" msiexec.exe /x "%~dp0$MsiFileName" /qn /norestart
exit /b 0
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        # Primary detection: msedge.exe in Edge\Application
        $clause1 = New-CMDetectionClauseFile `
            -Path "C:\Program Files (x86)\Microsoft\Edge\Application" `
            -FileName "msedge.exe" `
            -Value `
            -PropertyType Version `
            -ExpressionOperator GreaterEquals `
            -ExpectedValue $SoftwareVersion

        # Staged detection: msedge.exe in EdgeUpdate\Install (OR)
        $clause2 = New-CMDetectionClauseFile `
            -Path "C:\Program Files (x86)\Microsoft\EdgeUpdate\Install" `
            -FileName "msedge.exe" `
            -Value `
            -PropertyType Version `
            -ExpressionOperator GreaterEquals `
            -ExpectedValue $SoftwareVersion

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
            -AddDetectionClause @($clause1, $clause2) `
            -DetectionClauseConnector @{"LogicalName"=$clause2.Setting.LogicalName; "Connector"="OR"} `
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

function Get-EdgeNetworkAppRoot {
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
        $v = Get-LatestEdgeVersion -Quiet
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
    Write-Host "Microsoft Edge Enterprise (x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "FileServerPath               : $FileServerPath"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host "EdgeVersionUrl               : $EdgeVersionUrl"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-EdgeNetworkAppRoot -FileServerPath $FileServerPath

    $version = Get-LatestEdgeVersion
    if (-not $version) {
        throw "Could not resolve Edge version."
    }

    $localMsi    = Join-Path $BaseDownloadRoot $MsiFileName
    $contentPath = Join-Path $networkAppRoot $version

    Initialize-Folder -Path $contentPath

    $netMsi = Join-Path $contentPath $MsiFileName

    Write-Host "Version                      : $version"
    Write-Host "Local MSI                    : $localMsi"
    Write-Host "ContentPath                  : $contentPath"
    Write-Host "Network MSI                  : $netMsi"
    Write-Host ""

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Host "Downloading MSI..."
        curl.exe -L --fail --silent --show-error -o $localMsi $EdgeStableMSIUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $EdgeStableMSIUrl" }
    }
    else {
        Write-Host "Local MSI exists. Skipping download."
    }

    if (-not (Test-Path -LiteralPath $netMsi)) {
        Write-Host "Copying MSI to network..."
        Copy-Item -LiteralPath $localMsi -Destination $netMsi -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network MSI exists. Skipping copy."
    }

    $appName   = "Microsoft Edge Enterprise - $version"
    $publisher = "Microsoft Corporation"

    Write-Host ""
    Write-Host "CM Application Name          : $appName"
    Write-Host "CM SoftwareVersion           : $version"
    Write-Host ""

    New-MECMEdgeApplication `
        -AppName $appName `
        -SoftwareVersion $version `
        -ContentPath $contentPath `
        -MsiFileName $MsiFileName `
        -Publisher $publisher

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
