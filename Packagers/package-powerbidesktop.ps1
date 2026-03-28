<#
Vendor: Microsoft
App: Microsoft Power BI Desktop (x64)
CMName: Microsoft Power BI Desktop
VendorUrl: https://www.microsoft.com/download/details.aspx?id=58494
CPE: cpe:2.3:a:microsoft:power_bi_desktop:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/power-bi/fundamentals/desktop-latest-update
DownloadPageUrl: https://www.microsoft.com/en-us/download/details.aspx?id=58494

.SYNOPSIS
    Packages Microsoft Power BI Desktop (x64) for MECM.

.DESCRIPTION
    Parses the Microsoft Download Center page to retrieve the current Power BI
    Desktop x64 installer URL and version, downloads the EXE, stages content to
    a versioned local folder with file-based detection metadata, and creates an
    MECM Application with file version-based detection.
    Detection uses PBIDesktop.exe file version GreaterEquals packaged version.

    NOTE: Power BI Desktop is released monthly. The installer URL embedded in the
    download page changes with each release; there is no stable permalink for the
    x64 EXE. Installer is always re-downloaded.

    GetLatestVersionOnly fetches only the Microsoft Download Center page (small
    HTML) to read the current version - no installer download is performed.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\Power BI Desktop\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\PowerBI).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Parses the Microsoft Download Center page for the current version, outputs
    the version string, and exits. No download or MECM changes are made.

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
$DownloadPageUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=58494"
$ExeFileName     = "PBIDesktopSetup_x64.exe"

$VendorFolder = "Microsoft"
$AppFolder    = "Power BI Desktop"

$BaseDownloadRoot = Join-Path $DownloadRoot "PowerBI"

# --- Functions ---


function Get-PowerBIDownloadInfo {
    param([switch]$Quiet)

    Write-Log "Download page URL            : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch download page: $DownloadPageUrl" }

        $urlMatch = [regex]::Match($html, '"url"\s*:\s*"(https://download\.microsoft\.com/download/[^"]+PBIDesktopSetup_x64\.exe)"')
        if (-not $urlMatch.Success) { throw "Could not locate PBIDesktopSetup_x64.exe URL in download page." }

        $verMatch = [regex]::Match($html, '"[Vv]ersion"\s*:\s*"(\d+\.\d+\.\d+(?:\.\d+)?)"')
        if (-not $verMatch.Success) { throw "Could not parse version from download page." }

        $dlUrl   = $urlMatch.Groups[1].Value
        $version = $verMatch.Groups[1].Value

        Write-Log "Download URL                 : $dlUrl" -Quiet:$Quiet
        Write-Log "Latest Power BI version      : $version" -Quiet:$Quiet

        return [PSCustomObject]@{
            Version     = $version
            DownloadUrl = $dlUrl
        }
    }
    catch {
        Write-Log "Failed to get Power BI download info: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePowerBI {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Microsoft Power BI Desktop (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version and download URL ---
    $info = Get-PowerBIDownloadInfo
    if (-not $info) { throw "Could not resolve Power BI Desktop download info." }

    $version     = $info.Version
    $downloadUrl = $info.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $ExeFileName"
    Write-Log ""

    # --- Download (always re-download, URL changes each release) ---
    $localExe = Join-Path $BaseDownloadRoot $ExeFileName
    Write-Log "Local installer path         : $localExe"
    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $ExeFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $ExeFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''-quiet'', ''ACCEPT_EULA=1'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    # Uninstall: registry lookup for product code, then msiexec /x
    $uninstallContent = (
        '$app = Get-ChildItem `',
        '    ''HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'',',
        '    ''HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'' `',
        '    -ErrorAction SilentlyContinue |',
        '    Get-ItemProperty -ErrorAction SilentlyContinue |',
        '    Where-Object { $_.DisplayName -like ''Microsoft Power BI Desktop*'' } |',
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
    $detectionPath = "{0}\Microsoft Power BI Desktop\bin" -f $env:ProgramFiles

    $appName   = "Microsoft Power BI Desktop $version"
    $publisher = "Microsoft Corporation"

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : PBIDesktop.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $ExeFileName
        InstallerType   = "EXE"
        InstallArgs     = "-quiet ACCEPT_EULA=1"
        UninstallArgs   = "/qn /norestart"
        RunningProcess  = @("PBIDesktop")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "PBIDesktop.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
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

function Invoke-PackagePowerBI {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Microsoft Power BI Desktop (x64) - PACKAGE phase"
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
        $info = Get-PowerBIDownloadInfo -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
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
    Write-Log "Microsoft Power BI Desktop (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadPageUrl              : $DownloadPageUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StagePowerBI
    }
    elseif ($PackageOnly) {
        Invoke-PackagePowerBI
    }
    else {
        Invoke-StagePowerBI
        Invoke-PackagePowerBI
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
