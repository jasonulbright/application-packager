<#
Vendor: Adobe Inc.
App: Adobe Acrobat (Reader) DC (x64)
CMName: Adobe Acrobat (Reader) DC

.SYNOPSIS
    Packages the latest Adobe Acrobat (Reader) DC (x64) for MECM.

.DESCRIPTION
    Parses Adobe's official release notes page to determine the current Acrobat DC
    version, constructs the enterprise installer URL, downloads the x64 MUI EXE,
    stages content to a versioned network location, and creates an MECM Application
    with file version-based detection.

    Install:   AcroRdrDCx64{version}_MUI.exe /sAll /rs /rps /msi /qn /norestart
    Uninstall: PowerShell registry lookup (uninstall.ps1)
    Detection: Acrobat.exe file version >= packaged version

    GetLatestVersionOnly fetches only the Adobe release notes page (small HTML)
    to read the current version — no installer download is performed.

    Adobe Acrobat version notation:
      Release notes use format NN.NNN.NNNNN (e.g., 25.001.21223)
      Download URL uses the same parts concatenated (e.g., 2500121223)

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Adobe\Acrobat Reader DC\<Version>

.PARAMETER GetLatestVersionOnly
    Parses Adobe's release notes page for the current version, outputs the version
    string, and exits. No download or MECM changes are made.

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
$AdobeReleaseNotesUrl = "https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html"
$AdobeDownloadBase    = "https://ardownload3.adobe.com/pub/adobe/acrobat/win/AcrobatDC"

$VendorFolder = "Adobe"
$AppFolder    = "Acrobat Reader DC"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\AdobeReader"

$EstimatedRuntimeMins = 15
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

function Get-AdobeAcrobatVersion {
    param([switch]$Quiet)

    if (-not $Quiet) {
        Write-Host "Release notes URL            : $AdobeReleaseNotesUrl"
    }

    try {
        $html = (curl.exe -L --fail --silent --show-error $AdobeReleaseNotesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Adobe release notes: $AdobeReleaseNotesUrl" }

        $verMatch = [regex]::Match($html, '\b(\d{2}\.\d{3}\.\d{5})\b')
        if (-not $verMatch.Success) { throw "Could not parse Acrobat DC version from release notes page." }

        $version = $verMatch.Groups[1].Value

        if (-not $Quiet) {
            Write-Host "Latest Acrobat DC version    : $version"
        }
        return $version
    }
    catch {
        Write-Error "Failed to get Acrobat DC version: $($_.Exception.Message)"
        return $null
    }
}

function Get-AdobeInstallerInfo {
    param([Parameter(Mandatory)][string]$Version)

    $parts      = $Version -split '\.'
    $urlVersion = "$($parts[0])$($parts[1])$($parts[2])"
    $fileName   = "AcroRdrDCx64${urlVersion}_MUI.exe"
    $url        = "$AdobeDownloadBase/$urlVersion/$fileName"

    return [PSCustomObject]@{
        UrlVersion  = $urlVersion
        FileName    = $fileName
        DownloadUrl = $url
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

function New-MECMAdobeReaderApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$DetectionVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$InstallerFileName,
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
        $uninstallPs1Path = Join-Path $ContentPath "uninstall.ps1"

        if (-not (Test-Path -LiteralPath $installBatPath)) {
            $installBat = @"
@echo off
setlocal
start /wait "" "%~dp0$InstallerFileName" /sAll /rs /rps /msi /qn /norestart
exit /b 0
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallPs1Path)) {
            $uninstallPs1 = @'
$app = Get-ChildItem `
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
    -ErrorAction SilentlyContinue |
    Get-ItemProperty -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like 'Adobe Acrobat*' -and $_.UninstallString -match 'msiexec' } |
    Sort-Object DisplayVersion -Descending |
    Select-Object -First 1
if ($app) {
    Start-Process msiexec.exe -ArgumentList "/x `"$($app.PSChildName)`" /qn /norestart" -Wait -NoNewWindow
}
'@
            Set-Content -LiteralPath $uninstallPs1Path -Value $uninstallPs1 -Encoding UTF8 -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal
PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
exit /b %ERRORLEVEL%
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        $clause = New-CMDetectionClauseFile `
            -Path "$env:ProgramFiles\Adobe\Acrobat DC\Acrobat" `
            -FileName "Acrobat.exe" `
            -Value `
            -PropertyType Version `
            -ExpressionOperator GreaterEquals `
            -ExpectedValue $DetectionVersion `
            -Is64Bit

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

function Get-AdobeReaderNetworkAppRoot {
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
        $v = Get-AdobeAcrobatVersion -Quiet
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
    Write-Host "Adobe Acrobat (Reader) DC (x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "FileServerPath               : $FileServerPath"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host "AdobeReleaseNotesUrl         : $AdobeReleaseNotesUrl"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-AdobeReaderNetworkAppRoot -FileServerPath $FileServerPath

    $version = Get-AdobeAcrobatVersion
    if (-not $version) {
        throw "Could not resolve Acrobat DC version."
    }

    $dlInfo      = Get-AdobeInstallerInfo -Version $version
    $fileName    = $dlInfo.FileName
    $downloadUrl = $dlInfo.DownloadUrl
    $contentPath = Join-Path $networkAppRoot $version

    Initialize-Folder -Path $contentPath

    $localExe = Join-Path $BaseDownloadRoot $fileName
    $netExe   = Join-Path $contentPath $fileName

    Write-Host "Version                      : $version"
    Write-Host "URL version                  : $($dlInfo.UrlVersion)"
    Write-Host "Installer file               : $fileName"
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

    $exeFileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).FileVersion
    if ([string]::IsNullOrWhiteSpace($exeFileVersion)) {
        Write-Warning "Could not read FileVersion from EXE; using release notes version for detection."
        $exeFileVersion = $version
    }
    Write-Host "EXE FileVersion (detection)  : $exeFileVersion"

    if (-not (Test-Path -LiteralPath $netExe)) {
        Write-Host "Copying installer to network..."
        Copy-Item -LiteralPath $localExe -Destination $netExe -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network installer exists. Skipping copy."
    }

    $appName   = "Adobe Acrobat (Reader) DC $version"
    $publisher = "Adobe Inc."

    Write-Host ""
    Write-Host "CM Application Name          : $appName"
    Write-Host "CM SoftwareVersion           : $version"
    Write-Host "CM Detection Version         : $exeFileVersion"
    Write-Host ""

    New-MECMAdobeReaderApplication `
        -AppName $appName `
        -SoftwareVersion $version `
        -DetectionVersion $exeFileVersion `
        -ContentPath $contentPath `
        -InstallerFileName $fileName `
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
