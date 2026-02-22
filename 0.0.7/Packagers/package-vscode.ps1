<#
Vendor: Microsoft Corporation
App: Visual Studio Code (x64)
CMName: VS Code

.SYNOPSIS
    Packages Visual Studio Code (x64) for MECM.

.DESCRIPTION
    Downloads the latest VS Code x64 installer from the official update API,
    stages content to a versioned network location, and creates an MECM Application
    with file-based version detection.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\Visual Studio Code\<Version>

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available VS Code version string and exits.

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
$StaticDownloadUrl = "https://update.code.visualstudio.com/latest/win32-x64/stable"
$VersionApiUrl     = "https://update.code.visualstudio.com/api/releases/stable"

$VendorFolder = "Microsoft"
$AppFolder    = "Visual Studio Code"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\VSCode"

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

function Get-LatestVSCodeVersion {
    param([switch]$Quiet)

    if (-not $Quiet) {
        Write-Host "VS Code version API          : $VersionApiUrl"
    }

    try {
        $json = (curl.exe -L --fail --silent --show-error $VersionApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch version info: $VersionApiUrl" }

        $versions = ConvertFrom-Json $json
        if (-not $versions -or $versions.Count -eq 0) {
            throw "No versions found in API response."
        }

        $version = $versions[0]
        if (-not $Quiet) {
            Write-Host "Latest VS Code version       : $version"
        }
        return $version
    }
    catch {
        Write-Error "Failed to get VS Code version: $($_.Exception.Message)"
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

function New-MECMVSCodeApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$FileName,
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
start /wait "" "%~dp0$FileName" /VERYSILENT /NORESTART /FORCECLOSEAPPLICATIONS /MERGETASKS=!runcode
exit /b 0
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal
"C:\Program Files\Microsoft VS Code\unins000.exe" /SILENT
exit /b 0
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        $clause = New-CMDetectionClauseFile `
            -Path "$env:ProgramFiles\Microsoft VS Code" `
            -FileName "Code.exe" `
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

function Get-VSCodeNetworkAppRoot {
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
        $v = Get-LatestVSCodeVersion -Quiet
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
    Write-Host "VS Code (x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "FileServerPath               : $FileServerPath"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host "VersionApiUrl                : $VersionApiUrl"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-VSCodeNetworkAppRoot -FileServerPath $FileServerPath

    $version = Get-LatestVSCodeVersion
    if (-not $version) {
        throw "Could not resolve VS Code version."
    }

    $fileName    = "VSCodeSetup-x64-${version}.exe"
    $localExe    = Join-Path $BaseDownloadRoot $fileName
    $contentPath = Join-Path $networkAppRoot $version

    Initialize-Folder -Path $contentPath

    $netExe = Join-Path $contentPath $fileName

    Write-Host "Version                      : $version"
    Write-Host "Local installer              : $localExe"
    Write-Host "ContentPath                  : $contentPath"
    Write-Host "Network installer            : $netExe"
    Write-Host ""

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Host "Downloading installer..."
        curl.exe -L --fail --silent --show-error -o $localExe $StaticDownloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $StaticDownloadUrl" }
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

    $appName   = "VS Code $version"
    $publisher = "Microsoft Corporation"

    Write-Host ""
    Write-Host "CM Application Name          : $appName"
    Write-Host "CM SoftwareVersion           : $version"
    Write-Host ""

    New-MECMVSCodeApplication `
        -AppName $appName `
        -SoftwareVersion $version `
        -ContentPath $contentPath `
        -FileName $fileName `
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
