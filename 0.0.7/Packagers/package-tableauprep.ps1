<#
Vendor: Tableau
App: Tableau Prep Builder (x64)
CMName: Tableau Prep Builder

.SYNOPSIS
    Packages Tableau Prep Builder (x64) for MECM.

.DESCRIPTION
    Downloads the latest Tableau Prep Builder x64 installer from the official
    Tableau download server, stages content to a versioned network location,
    temporarily installs the product to extract registry metadata and file
    versions, and creates an MECM Application with file-based version detection.
    Detection uses tableau-prep-builder.exe version from the registry-discovered
    install location.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Tableau\Prep\<Version>

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Tableau Prep Builder version string and exits.

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
    [string]$LogPath,
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

. "$PSScriptRoot\AppPackagerCommon.ps1"
Initialize-Logging -LogPath $LogPath

# --- Configuration ---
$BaseDownloadUrl = "https://downloads.tableau.com/tssoftware/"
$ReleaseNotesUrl = "https://www.tableau.com/support/releases"

$VendorFolder = "Tableau"
$AppFolder    = "Prep"

$InstallerPrefix        = "TableauPrep-64bit"
$DetectionFileName      = "tableau-prep-builder.exe"
$RegistryPrefix         = "Tableau 20"
$DisplayNameMustContain = "Prep"

$InstallArgs   = "/install /quiet /norestart ACCEPTEULA=1 SENDTELEMETRY=0"
$UninstallArgs = "/uninstall /quiet /norestart"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\TableauPrep"

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
        Write-Log "Admin check failed: $($_.Exception.Message)" -Level WARN
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
        Write-Log "Connected to CM site: $SiteCode"
        return $true
    }
    catch {
        Write-Log "Failed to connect to CM site: $($_.Exception.Message)" -Level ERROR
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
            Write-Log "Network path does not exist or is inaccessible: $Path" -Level ERROR
            return $false
        }

        try {
            $tmp = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
            Set-Content -LiteralPath $tmp -Value "test" -Encoding ASCII -ErrorAction Stop
            Remove-Item -LiteralPath $tmp -ErrorAction Stop
            return $true
        }
        catch {
            Write-Log "Network share is not writable: $Path ($($_.Exception.Message))" -Level ERROR
            return $false
        }
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

function Get-LatestTableauVersion {
    param([switch]$Quiet)

    Write-Log "Tableau release notes URL     : $ReleaseNotesUrl" -Quiet:$Quiet

    try {
        $HtmlContent = (curl.exe -L --fail --silent --show-error $ReleaseNotesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch release notes: $ReleaseNotesUrl" }

        $versionPattern = '\b(20\d{2})[.-](\d+)(?:[.-](\d+))?\b'
        $regexMatches = [regex]::Matches($HtmlContent, $versionPattern)

        if ($regexMatches.Count -eq 0) {
            throw "No version matches found in release notes."
        }

        $versions = foreach ($m in $regexMatches) {
            $year  = $m.Groups[1].Value
            $minor = $m.Groups[2].Value
            $patch = $m.Groups[3].Value
            if (-not $patch) { $patch = "0" }
            "{0}.{1}.{2}" -f $year, $minor, $patch
        }

        $versions = $versions | Select-Object -Unique

        $latest = $versions | Sort-Object -Descending -Property @{
                Expression = { [int](($_ -split '\.')[0]) }
            }, @{
                Expression = { [int](($_ -split '\.')[1]) }
            }, @{
                Expression = { [int](($_ -split '\.')[2]) }
            } | Select-Object -First 1

        if (-not $latest) {
            throw "Could not determine latest Tableau version."
        }

        Write-Log "Latest Tableau version        : $latest" -Quiet:$Quiet
        return $latest
    }
    catch {
        Write-Log "Failed to get Tableau version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-InstalledAppRegistryInfo {
    param(
        [Parameter(Mandatory)][string]$DisplayNamePrefix,
        [Parameter(Mandatory)][string]$DisplayNameMustContain
    )

    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $uninstallPaths) {
        $apps = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DisplayName -and
                $_.DisplayName -like "$DisplayNamePrefix*" -and
                $_.DisplayName -like "*$DisplayNameMustContain*"
            } |
            Sort-Object -Property DisplayVersion -Descending

        if ($apps -and $apps.Count -gt 0) {
            return $apps | Select-Object -First 1
        }
    }

    return $null
}

function Get-FileVersion {
    param([Parameter(Mandatory)][string]$FilePath)

    if (-not (Test-Path -LiteralPath $FilePath)) { return $null }
    try {
        return (Get-Item -LiteralPath $FilePath).VersionInfo.FileVersion
    }
    catch {
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

function New-MECMTableauPrepApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$InstallerFileName,
        [Parameter(Mandatory)][string]$Publisher,
        [Parameter(Mandatory)][string]$DetectionExePath,
        [Parameter(Mandatory)][string]$DetectionExeVersion
    )

    $orig = Get-Location

    try {
        if (-not (Test-IsAdmin)) { throw "Run PowerShell as Administrator." }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $existing = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Log "Application already exists: $AppName" -Level WARN
            return
        }

        Write-Log "Creating CM Application      : $AppName"
        $cmApp = New-CMApplication `
            -Name $AppName `
            -Publisher $Publisher `
            -SoftwareVersion $SoftwareVersion `
            -Description $Comment `
            -AutoInstall $true `
            -ErrorAction Stop

        Write-Log "Application CI_ID            : $($cmApp.CI_ID)"

        Set-Location C: -ErrorAction Stop

        $installBatPath   = Join-Path $ContentPath "install.bat"
        $installPs1Path   = Join-Path $ContentPath "install.ps1"
        $uninstallBatPath = Join-Path $ContentPath "uninstall.bat"
        $uninstallPs1Path = Join-Path $ContentPath "uninstall.ps1"

        if (-not (Test-Path -LiteralPath $installBatPath)) {
            $installBat = @"
@echo off
PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"
exit /b %ERRORLEVEL%
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $installPs1Path)) {
            $installPs1 = @"
`$proc = Start-Process "`$PSScriptRoot\$InstallerFileName" -ArgumentList "$InstallArgs" -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@
            Set-Content -LiteralPath $installPs1Path -Value $installPs1 -Encoding UTF8 -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
exit /b %ERRORLEVEL%
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallPs1Path)) {
            $uninstallPs1 = @"
`$proc = Start-Process "`$PSScriptRoot\$InstallerFileName" -ArgumentList "$UninstallArgs" -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@
            Set-Content -LiteralPath $uninstallPs1Path -Value $uninstallPs1 -Encoding UTF8 -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        $clause = New-CMDetectionClauseFile `
            -Path (Split-Path -Path $DetectionExePath -Parent) `
            -FileName (Split-Path -Path $DetectionExePath -Leaf) `
            -Value `
            -PropertyType Version `
            -ExpressionOperator GreaterEquals `
            -ExpectedValue $DetectionExeVersion

        Write-Log "Adding Script Deployment Type: $dtName"
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

        Write-Log "Created MECM application     : $AppName"
    }
    finally {
        Set-Location $orig -ErrorAction SilentlyContinue
    }
}

function Get-TableauPrepNetworkAppRoot {
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
        $v = Get-LatestTableauVersion -Quiet
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

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Tableau Prep Builder (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleaseNotesUrl              : $ReleaseNotesUrl"
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-TableauPrepNetworkAppRoot -FileServerPath $FileServerPath

    $version = Get-LatestTableauVersion
    if (-not $version) {
        throw "Could not resolve Tableau Prep Builder version."
    }

    $installerFileName = "${InstallerPrefix}-${version}.exe"
    $localExe    = Join-Path $BaseDownloadRoot $installerFileName
    $contentPath = Join-Path $networkAppRoot $version

    Initialize-Folder -Path $contentPath

    $netExe = Join-Path $contentPath $installerFileName

    Write-Log "Version                      : $version"
    Write-Log "Local installer              : $localExe"
    Write-Log "ContentPath                  : $contentPath"
    Write-Log "Network installer            : $netExe"
    Write-Log ""

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading installer..."
        $downloadUrl = "${BaseDownloadUrl}${installerFileName}"
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    if (-not (Test-Path -LiteralPath $netExe)) {
        Write-Log "Copying installer to network..."
        Copy-Item -LiteralPath $localExe -Destination $netExe -Force -ErrorAction Stop
    }
    else {
        Write-Log "Network installer exists. Skipping copy."
    }

    # --- Temporary install for metadata extraction ---
    Write-Log ""
    Write-Log "Installing temporarily for metadata extraction..."
    Start-Process -FilePath $localExe -ArgumentList $InstallArgs -Wait -NoNewWindow

    $regInfo = Get-InstalledAppRegistryInfo -DisplayNamePrefix $RegistryPrefix -DisplayNameMustContain $DisplayNameMustContain
    if (-not $regInfo) {
        Write-Log "Could not find installed application in registry: $RegistryPrefix (*$DisplayNameMustContain*)" -Level ERROR
        Write-Log "Uninstalling after failed metadata extraction..."
        Start-Process -FilePath $localExe -ArgumentList $UninstallArgs -Wait -NoNewWindow
        exit 1
    }

    $displayName     = $regInfo.DisplayName
    $displayVersion  = $regInfo.DisplayVersion
    $publisher       = $regInfo.Publisher
    $installLocation = $regInfo.InstallLocation

    if (-not $installLocation) {
        $installLocation = "C:\Program Files\Tableau"
    }

    $detectionExe        = Join-Path $installLocation $DetectionFileName
    $detectionExeVersion = Get-FileVersion -FilePath $detectionExe

    if (-not $detectionExeVersion) {
        Write-Log "Could not determine file version for: $detectionExe" -Level WARN
        $detectionExeVersion = $displayVersion
    }

    Write-Log "Uninstalling after metadata extraction..."
    Start-Process -FilePath $localExe -ArgumentList $UninstallArgs -Wait -NoNewWindow

    if (-not $publisher) { $publisher = "Tableau" }

    Write-Log ""
    Write-Log "CM Application Name          : $displayName"
    Write-Log "CM SoftwareVersion           : $displayVersion"
    Write-Log "Detection exe                : $detectionExe"
    Write-Log "Detection exe version        : $detectionExeVersion"
    Write-Log ""

    New-MECMTableauPrepApplication `
        -AppName $displayName `
        -SoftwareVersion $displayVersion `
        -ContentPath $contentPath `
        -InstallerFileName $installerFileName `
        -Publisher $publisher `
        -DetectionExePath $detectionExe `
        -DetectionExeVersion $detectionExeVersion

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
