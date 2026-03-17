<#
Vendor: Cloud Software Group
App: XenServer VM Tools (x64)
CMName: XenServer VM Tools
VendorUrl: https://www.xenserver.com/
CPE: cpe:2.3:a:cloud_software_group:xenserver_vm_tools:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.xenserver.com/en-us/xenserver/8/whats-new
DownloadPageUrl: https://www.xenserver.com/downloads

.SYNOPSIS
    Packages XenServer VM Tools (Windows x64) MSI for MECM.

.DESCRIPTION
    Scrapes the XenServer downloads page for the latest Windows VM Tools
    (Management Agent) MSI URL, downloads it, stages content to a versioned
    local folder with ARP detection metadata, and creates an MECM Application
    with registry-based detection.

    Detection uses MSI ProductCode-based ARP registry key with DisplayVersion.

    Install parameters include ALLOWAUTOUPDATE=NO and IDENTIFYAUTOUPDATE=NO
    to prevent the agent from self-updating outside of MECM control.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (Package phase)
    - Local administrator
    - Write access to FileServerPath (Package phase)
    - Internet access (Stage phase)
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
$DownloadPageUrl  = "https://www.xenserver.com/downloads"

$VendorFolder     = "Cloud Software Group"
$AppFolder        = "XenServer VM Tools (x64)"
$BaseDownloadRoot = Join-Path $DownloadRoot "XenServerVMTools-x64"
$MsiFileName      = "managementagentx64.msi"

# --- Functions ---


function Resolve-VMToolsMsiUrl {
    param([switch]$Quiet)

    Write-Log "XenServer downloads page     : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch XenServer downloads page." }

        # Match: https://downloads.xenserver.com/vm-tools-windows/<version>/managementagentx64-<version>.msi
        $rx = [regex]'href\s*=\s*"(?<url>https://downloads\.xenserver\.com/vm-tools-windows/(?<ver>[^/]+)/managementagentx64[^"]*\.msi)"'
        $m = $rx.Match($html)

        if (-not $m.Success) {
            throw "Could not locate VM Tools Windows MSI link on the downloads page."
        }

        $url = $m.Groups["url"].Value
        Write-Log "Resolved MSI URL             : $url" -Quiet:$Quiet
        return $url
    }
    catch {
        Write-Log "Failed to resolve VM Tools MSI URL: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageVMTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "XenServer VM Tools (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    $msiUrl = Resolve-VMToolsMsiUrl
    if (-not $msiUrl) { throw "Could not resolve VM Tools MSI download URL." }

    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"
    Write-Log ""
    Write-Log "Downloading MSI..."
    Invoke-DownloadWithRetry -Url $msiUrl -OutFile $localMsi

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]
    $manufacturer      = $props["Manufacturer"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productName))       { throw "MSI ProductName missing." }
    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion           : $productVersionRaw"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $productVersionRaw
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $MsiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Derive ARP detection from MSI properties ---
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode
    $arpEntry = [pscustomobject]@{
        RegistryKeyRelative = $arpRegistryKey
        DisplayName         = $productName
        DisplayVersion      = $productVersionRaw
        Is64Bit             = $true
    }
    Write-Log "ARP detection derived from MSI properties (no temp install needed)."

    Write-Log ""
    Write-Log "ARP DisplayName              : $($arpEntry.DisplayName)"
    Write-Log "ARP DisplayVersion           : $($arpEntry.DisplayVersion)"
    Write-Log "ARP RegistryKey              : $($arpEntry.RegistryKeyRelative)"
    Write-Log "ARP Is64Bit                  : $($arpEntry.Is64Bit)"
    Write-Log ""

    # --- Generate content wrappers ---
    # ALLOWAUTOUPDATE=NO IDENTIFYAUTOUPDATE=NO prevents self-update outside MECM
    $installPs1 = @(
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/i'', "`"$msiPath`"", ''ALLOWAUTOUPDATE=NO'', ''IDENTIFYAUTOUPDATE=NO'', ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallPs1 = @(
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/x'', "`"$msiPath`"", ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $publisher = if (-not [string]::IsNullOrWhiteSpace($manufacturer)) { $manufacturer } else { "Cloud Software Group" }
    $appName = "XenServer VM Tools - $productVersionRaw (x64)"

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpEntry.RegistryKeyRelative
            ValueName           = "DisplayVersion"
            DisplayName         = $arpEntry.DisplayName
            DisplayVersion      = $arpEntry.DisplayVersion
            Is64Bit             = $arpEntry.Is64Bit
        }
    }

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageVMTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "XenServer VM Tools (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    if (-not (Test-Path -LiteralPath $localMsi)) {
        throw "Local MSI not found - run Stage phase first: $localMsi"
    }

    $props = Get-MsiPropertyMap -MsiPath $localMsi
    if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) {
        throw "Cannot read ProductVersion from cached MSI."
    }

    $productVersionRaw = $props["ProductVersion"]
    $localContentPath  = Join-Path $BaseDownloadRoot $productVersionRaw
    $manifestPath      = Join-Path $localContentPath "stage-manifest.json"

    # --- Read manifest ---
    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.DisplayVersion)"
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
        Initialize-Folder -Path $BaseDownloadRoot

        $msiUrl = Resolve-VMToolsMsiUrl -Quiet
        if (-not $msiUrl) { exit 1 }

        $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
        Invoke-DownloadWithRetry -Url $msiUrl -OutFile $localMsi -Quiet

        $props = Get-MsiPropertyMap -MsiPath $localMsi
        if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) { exit 1 }

        Write-Output $props["ProductVersion"]
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
    Write-Log "XenServer VM Tools (x64) Auto-Packager starting"
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
        Invoke-StageVMTools
    }
    elseif ($PackageOnly) {
        Invoke-PackageVMTools
    }
    else {
        Invoke-StageVMTools
        Invoke-PackageVMTools
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
