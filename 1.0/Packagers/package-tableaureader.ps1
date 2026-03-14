<#
Vendor: Salesforce (Tableau)
App: Tableau Reader (x64)
CMName: Tableau Reader
VendorUrl: https://www.tableau.com/products/reader
CPE: cpe:2.3:a:tableau:reader:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.tableau.com/support/releases
DownloadPageUrl: https://www.tableau.com/support/releases

.SYNOPSIS
    Packages Tableau Reader (x64) for MECM.

.DESCRIPTION
    Downloads the latest Tableau Reader x64 installer from the Tableau CDN,
    stages content to a versioned local folder with file-based version detection
    metadata, and creates an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download installer, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The latest version is resolved by scraping the Tableau release notes page.
    Tableau Reader is free (no license required).

    Detection uses the DisplayVersion registry value under the WOW6432Node
    uninstall key (ProductCode extracted dynamically from the Burn manifest).

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Tableau\Tableau Reader (x64)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 30

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 60

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Scrapes the Tableau release notes page for the latest version, outputs the
    version string, and exits. No MECM changes are made.

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
    [int]$EstimatedRuntimeMins = 30,
    [int]$MaximumRuntimeMins = 60,
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
$BaseDownloadUrl = "https://downloads.tableau.com/tssoftware/"
$ReleaseNotesUrl = "https://www.tableau.com/support/releases"

$VendorFolder = "Tableau"
$AppFolder    = "Tableau Reader (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "TableauReader"

# --- Functions ---


function Get-LatestTableauVersion {
    param([switch]$Quiet)

    Write-Log "Tableau release notes URL     : $ReleaseNotesUrl" -Quiet:$Quiet

    try {
        $htmlContent = (curl.exe -L --fail --silent --show-error $ReleaseNotesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch release notes: $ReleaseNotesUrl" }

        $versionPattern = '\b(20\d{2})[.-](\d+)(?:[.-](\d+))?\b'
        $regexMatches = [regex]::Matches($htmlContent, $versionPattern)

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


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTableauReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Tableau Reader (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get latest version ---
    $version = Get-LatestTableauVersion
    if (-not $version) { throw "Could not determine latest Tableau version." }

    # URL uses hyphens with 64bit: TableauReader-64bit-2025-3-3.exe
    $versionHyphenated = $version -replace '\.', '-'
    $installerFileName = "TableauReader-64bit-${versionHyphenated}.exe"
    $downloadUrl       = "${BaseDownloadUrl}${installerFileName}"

    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Tableau Reader..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

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
    $installContent = @(
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/install'', ''/quiet'', ''/norestart'', ''ACCEPTEULA=1'', ''SENDTELEMETRY=0'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = @(
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/uninstall'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Extract ProductCode and MSI version from Burn manifest ---
    $extractPath = Join-Path $BaseDownloadRoot "_extract"
    if (Test-Path -LiteralPath $extractPath) { Remove-Item -LiteralPath $extractPath -Recurse -Force }
    $sevenZip = "C:\Program Files\7-Zip\7z.exe"
    Write-Log "Extracting Burn manifest from installer..."
    $proc = Start-Process -FilePath $sevenZip -ArgumentList @('x', "`"$localExe`"", "-o`"$extractPath`"", '-y') -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -gt 1) { throw "7z extraction failed with exit code $($proc.ExitCode)" }

    $manifestXml = Join-Path $extractPath "u23"
    if (-not (Test-Path -LiteralPath $manifestXml)) { throw "Burn manifest (u23) not found in extracted payload" }
    $xmlContent = [System.IO.File]::ReadAllText($manifestXml)
    $xml = [xml]$xmlContent

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('ba', 'http://schemas.microsoft.com/wix/2010/BootstrapperApplicationData')
    $pkgNode = $xml.SelectSingleNode("//ba:WixPackageProperties[@Package='Tableau']", $ns)
    if (-not $pkgNode) { throw "Could not find Tableau package node in Burn manifest" }

    $productCode = $pkgNode.GetAttribute('ProductCode')
    $msiVersion  = $pkgNode.GetAttribute('Version')
    Write-Log "ProductCode (from manifest)  : $productCode"
    Write-Log "MSI Version (from manifest)  : $msiVersion"

    Remove-Item -LiteralPath $extractPath -Recurse -Force -ErrorAction SilentlyContinue

    # --- Write stage manifest ---
    $appName   = "Tableau Reader $version (x64)"
    $publisher = "Salesforce (Tableau)"
    $arpKey = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$productCode"

    Write-Log "Detection registry key       : $arpKey"
    Write-Log "Detection DisplayVersion     : $msiVersion"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $msiVersion
            Operator            = "IsEquals"
            Is64Bit             = $false
        }
    }

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageTableauReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Tableau Reader (x64) - PACKAGE phase"
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
    Write-Log "Tableau Reader (x64) Auto-Packager starting"
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

    if ($StageOnly) {
        Invoke-StageTableauReader
    }
    elseif ($PackageOnly) {
        Invoke-PackageTableauReader
    }
    else {
        Invoke-StageTableauReader
        Invoke-PackageTableauReader
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
