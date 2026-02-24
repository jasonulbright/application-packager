<#
Vendor: Adobe Inc.
App: Adobe Acrobat (Reader) DC (x64)
CMName: Adobe Acrobat (Reader) DC
VendorUrl: https://www.adobe.com/acrobat/pdf-reader.html

.SYNOPSIS
    Packages the latest Adobe Acrobat (Reader) DC (x64) for MECM.

.DESCRIPTION
    Parses Adobe's official release notes page to determine the current Acrobat DC
    version, constructs the enterprise installer URL, downloads the x64 MUI EXE,
    stages content to a versioned local folder with file-based detection metadata,
    and creates an MECM Application with file version-based detection.
    Detection uses Acrobat.exe file version >= packaged version in the Program
    Files install path.

    Adobe Acrobat version notation:
      Release notes use format NN.NNN.NNNNN (e.g., 25.001.21223)
      Download URL uses the same parts concatenated (e.g., 2500121223)

    Supports two-phase operation:
      -StageOnly    Download, read FileVersion, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Adobe\Acrobat Reader DC\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\AdobeReader).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, read FileVersion, generate
    content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

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
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$AdobeReleaseNotesUrl = "https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html"
$AdobeDownloadBase    = "https://ardownload3.adobe.com/pub/adobe/acrobat/win/AcrobatDC"

$VendorFolder = "Adobe"
$AppFolder    = "Acrobat Reader DC"

$BaseDownloadRoot = Join-Path $DownloadRoot "AdobeReader"

# --- Functions ---


function Get-AdobeAcrobatVersion {
    param([switch]$Quiet)

    Write-Log "Release notes URL            : $AdobeReleaseNotesUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $AdobeReleaseNotesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Adobe release notes: $AdobeReleaseNotesUrl" }

        $verMatch = [regex]::Match($html, '\b(\d{2}\.\d{3}\.\d{5})\b')
        if (-not $verMatch.Success) { throw "Could not parse Acrobat DC version from release notes page." }

        $version = $verMatch.Groups[1].Value

        Write-Log "Latest Acrobat DC version    : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Acrobat DC version: $($_.Exception.Message)" -Level ERROR
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


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAdobeReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Adobe Acrobat (Reader) DC (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-AdobeAcrobatVersion
    if (-not $version) { throw "Could not resolve Acrobat DC version." }

    $dlInfo            = Get-AdobeInstallerInfo -Version $version
    $installerFileName = $dlInfo.FileName
    $downloadUrl       = $dlInfo.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "URL version                  : $($dlInfo.UrlVersion)"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # --- Read FileVersion from EXE for detection ---
    $exeFileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).FileVersion
    if ([string]::IsNullOrWhiteSpace($exeFileVersion)) {
        Write-Log "Could not read FileVersion from EXE; using release notes version for detection." -Level WARN
        $exeFileVersion = $version
    }
    Write-Log "EXE FileVersion (detection)  : $exeFileVersion"

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/sAll'', ''/rs'', ''/rps'', ''/msi'', ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        '$app = Get-ChildItem `',
        '    ''HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'',',
        '    ''HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'' `',
        '    -ErrorAction SilentlyContinue |',
        '    Get-ItemProperty -ErrorAction SilentlyContinue |',
        '    Where-Object { $_.DisplayName -like ''Adobe Acrobat*'' -and $_.UninstallString -match ''msiexec'' } |',
        '    Sort-Object DisplayVersion -Descending |',
        '    Select-Object -First 1',
        'if ($app) {',
        '    $proc = Start-Process msiexec.exe -ArgumentList @(''/x'', $app.PSChildName, ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        '    exit $proc.ExitCode',
        '}'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $detectionPath = "{0}\Adobe\Acrobat DC\Acrobat" -f $env:ProgramFiles

    $appName   = "Adobe Acrobat (Reader) DC $version"
    $publisher = "Adobe Inc."

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : Acrobat.exe"
    Write-Log "Detection version            : $exeFileVersion"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "Acrobat.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $exeFileVersion
            Is64Bit       = $true
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageAdobeReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Adobe Acrobat (Reader) DC (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }
    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()

    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    # --- Read manifest ---
    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
    Write-Log "Detection Version            : $($manifest.Detection.ExpectedValue)"
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network ---
    $localFiles = Get-ChildItem -Path $localContentPath -File -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $dest = Join-Path $networkContentPath $f.Name
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $($f.Name)"
        }
        else {
            Write-Log "Already on network           : $($f.Name)"
        }
    }

    # --- MECM application ---
    New-MECMApplicationFromManifest `
        -Manifest $manifest `
        -SiteCode $SiteCode `
        -Comment $Comment `
        -NetworkContentPath $networkContentPath `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins
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

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Adobe Acrobat (Reader) DC (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "AdobeReleaseNotesUrl         : $AdobeReleaseNotesUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageAdobeReader
    }
    elseif ($PackageOnly) {
        Invoke-PackageAdobeReader
    }
    else {
        Invoke-StageAdobeReader
        Invoke-PackageAdobeReader
    }

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
