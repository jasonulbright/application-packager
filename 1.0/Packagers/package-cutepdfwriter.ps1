<#
Vendor: Acro Software Inc.
App: CutePDF Writer
CMName: CutePDF Writer
VendorUrl: https://www.cutepdf.com/
ReleaseNotesUrl: https://www.cutepdf.com/Products/CutePDF/writer.asp
DownloadPageUrl: https://www.cutepdf.com/Products/CutePDF/writer.asp

.SYNOPSIS
    Packages CutePDF Writer for MECM.

.DESCRIPTION
    Downloads the latest CutePDF Writer EXE and its required Ghostscript converter
    from the CutePDF website, stages content to a versioned local folder with
    ARP-based detection metadata, and creates an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download both installers, read version from EXE, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    CutePDF Writer is a virtual PDF printer driver. It requires the Ghostscript-based
    converter (converter.exe) to function. The install wrapper installs the converter
    first, then CutePDF Writer.

    CutePDF Writer installs as a 32-bit application on 64-bit systems. The ARP key
    is in the Wow6432Node hive.

    GetLatestVersionOnly downloads the EXE and reads its file version.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Acro Software\CutePDF Writer\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Downloads the CutePDF Writer EXE, reads the file version, outputs the version
    string, and exits. No MECM changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
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
$WriterDownloadUrl    = "http://www.cutepdf.com/download/CuteWriter.exe"
$ConverterDownloadUrl = "http://www.cutepdf.com/download/converter.exe"
$WriterFileName       = "CuteWriter.exe"
$ConverterFileName    = "converter.exe"

$VendorFolder = "Acro Software"
$AppFolder    = "CutePDF Writer"

$BaseDownloadRoot = Join-Path $DownloadRoot "CutePDFWriter"

# --- Functions ---


function Get-LatestCutePDFVersion {
    param([switch]$Quiet)

    Write-Log "CutePDF Writer URL           : $WriterDownloadUrl" -Quiet:$Quiet

    try {
        Initialize-Folder -Path $BaseDownloadRoot
        $localExe = Join-Path $BaseDownloadRoot $WriterFileName

        if (-not (Test-Path -LiteralPath $localExe)) {
            Write-Log "Downloading CutePDF Writer EXE to read version..." -Quiet:$Quiet
            Invoke-DownloadWithRetry -Url $WriterDownloadUrl -OutFile $localExe
        }

        $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe)
        $version = $fileInfo.FileVersion
        if ([string]::IsNullOrWhiteSpace($version)) {
            $version = $fileInfo.ProductVersion
        }
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Could not read file version from CuteWriter.exe."
        }
        $version = $version.Trim()

        Write-Log "Latest CutePDF Writer version: $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get CutePDF version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCutePDF {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CutePDF Writer - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download CuteWriter.exe ---
    $localWriter = Join-Path $BaseDownloadRoot $WriterFileName
    Write-Log "Writer download URL          : $WriterDownloadUrl"
    Write-Log "Local writer path            : $localWriter"
    Write-Log ""

    if (-not (Test-Path -LiteralPath $localWriter)) {
        Write-Log "Downloading CutePDF Writer..."
        Invoke-DownloadWithRetry -Url $WriterDownloadUrl -OutFile $localWriter
    }
    else {
        Write-Log "Local writer exists. Skipping download."
    }

    # --- Read version from EXE ---
    $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localWriter)
    $version = $fileInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($version)) {
        $version = $fileInfo.ProductVersion
    }
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Could not read file version from CuteWriter.exe."
    }
    $version = $version.Trim()

    Write-Log "File version                 : $version"
    Write-Log ""

    # --- Download converter.exe (Ghostscript dependency) ---
    $localConverter = Join-Path $BaseDownloadRoot $ConverterFileName
    Write-Log "Converter download URL       : $ConverterDownloadUrl"
    Write-Log "Local converter path         : $localConverter"

    if (-not (Test-Path -LiteralPath $localConverter)) {
        Write-Log "Downloading Ghostscript converter..."
        Invoke-DownloadWithRetry -Url $ConverterDownloadUrl -OutFile $localConverter
    }
    else {
        Write-Log "Local converter exists. Skipping download."
    }
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    # Copy both files to staged folder
    foreach ($file in @($localWriter, $localConverter)) {
        $fileName = Split-Path $file -Leaf
        $stagedFile = Join-Path $localContentPath $fileName
        if (-not (Test-Path -LiteralPath $stagedFile)) {
            Copy-Item -LiteralPath $file -Destination $stagedFile -Force -ErrorAction Stop
            Write-Log "Copied to staged folder      : $fileName"
        }
        else {
            Write-Log "Staged file exists           : $fileName"
        }
    }

    # --- Generate content wrappers ---
    # Custom install wrapper: converter first, then CuteWriter
    $installPs1 = @(
        "# Install Ghostscript converter (required dependency)"
        "`$converterPath = Join-Path `$PSScriptRoot '$ConverterFileName'"
        "`$proc = Start-Process -FilePath `$converterPath -ArgumentList @('/S') -Wait -PassThru -NoNewWindow"
        "if (`$proc.ExitCode -ne 0) { exit `$proc.ExitCode }"
        ""
        "# Install CutePDF Writer"
        "`$writerPath = Join-Path `$PSScriptRoot '$WriterFileName'"
        "`$proc = Start-Process -FilePath `$writerPath -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/SP-') -Wait -PassThru -NoNewWindow"
        "exit `$proc.ExitCode"
    ) -join "`r`n"

    # CutePDF installs to Program Files (x86) on 64-bit systems
    $uninstallPs1 = @(
        "`$uninstaller = `"C:\Program Files (x86)\CutePDF Writer\unInstcpw64.exe`""
        "if (-not (Test-Path -LiteralPath `$uninstaller)) {"
        "    `$uninstaller = `"C:\Program Files\CutePDF Writer\unInstcpw.exe`""
        "}"
        "if (-not (Test-Path -LiteralPath `$uninstaller)) { Write-Error 'CutePDF uninstaller not found.'; exit 1 }"
        "`$proc = Start-Process -FilePath `$uninstaller -ArgumentList @('/uninstall', '/s') -Wait -PassThru -NoNewWindow"
        "exit `$proc.ExitCode"
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    # CutePDF Writer uses Wow6432Node ARP key on 64-bit systems
    $arpRegistryKey = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\CutePDF Writer Installation"

    $appName   = "CutePDF Writer $version"
    $publisher = "Acro Software Inc."

    Write-Log "ARP Registry Key             : $arpRegistryKey"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $WriterFileName
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $version
            Is64Bit             = $false
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

function Invoke-PackageCutePDF {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CutePDF Writer - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }
    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()

    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

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
        $v = Get-LatestCutePDFVersion -Quiet
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
    Write-Log "CutePDF Writer Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "WriterDownloadUrl            : $WriterDownloadUrl"
    Write-Log "ConverterDownloadUrl         : $ConverterDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCutePDF
    }
    elseif ($PackageOnly) {
        Invoke-PackageCutePDF
    }
    else {
        Invoke-StageCutePDF
        Invoke-PackageCutePDF
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
