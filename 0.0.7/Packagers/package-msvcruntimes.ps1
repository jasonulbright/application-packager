<#
Vendor: Microsoft
App: Microsoft Visual C++ 2015-2022 Redistributable - 14
CMName: Microsoft Visual C++ 2015-2022 Redistributable - 14

.SYNOPSIS
    Downloads the latest Microsoft Visual C++ 2015-2022 Redistributables (x86 + x64) and creates an MECM application.

.DESCRIPTION
    - Determines latest supported VC++ runtime version from Microsoft Learn lifecycle FAQ table; falls back to live installer file version when needed
    - Downloads vc_redist.x86.exe + vc_redist.x64.exe (permalinks)
    - Stages content to a versioned network folder
    - Creates static install.bat / uninstall.bat (with logging)
    - Creates MECM Application + Script Installer deployment type
    - Detection uses BOTH registry values (AND):
        HKLM\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X86  Version == v<ver>.00
        HKLM\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X64  Version == v<ver>.00
    - Deployment type content options:
        - Allow clients to use fallback source location for content
        - Download content from DP and run locally
    - Restart behavior: force device restart

.PARAMETER GetLatestVersionOnly
    Outputs only the latest version string and exits.
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Configuration ---
$NetworkRootPath   = Join-Path $FileServerPath "Applications\Microsoft\VC++ 2015-2022 Redistributable"
$BaseDownloadRoot  = Join-Path $env:USERPROFILE "Downloads\_AutoPackager"

$LifecycleFaqUrl   = "https://learn.microsoft.com/en-us/lifecycle/faq/visual-c-faq#what-versions-of-visual-c---redistributable--msvc-runtime-libraries--and-msvc-build-tools-are-supported-"
$VcRedistInfoUrl   = "https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170"

$UrlX86            = "https://aka.ms/vc14/vc_redist.x86.exe"
$UrlX64            = "https://aka.ms/vc14/vc_redist.x64.exe"

$Publisher         = "Microsoft Corporation"

$EstimatedRuntimeMins = 15
$MaximumRuntimeMins   = 45

# Registry detection (stable)
$RegKeyX86      = "SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X86"
$RegKeyX64      = "SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X64"
$RegValueName   = "Version"

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
        Set-Location "${SiteCode}:" -ErrorAction Stop
        Write-Host "Connected to CM site: $SiteCode"
        return $true
    }
    catch {
        Write-Error "Failed to connect to CM site: $($_.Exception.Message)"
        return $false
    }
}


function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)

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

function Get-ExeFileVersion {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Quiet
    )

    $vi = (Get-Item -LiteralPath $Path).VersionInfo
    if (-not $vi) { throw "Could not read VersionInfo from: $Path" }

    $fv = $vi.FileVersion
    $pv = $vi.ProductVersion

    if (-not [string]::IsNullOrWhiteSpace($fv)) { $fv = $fv.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($pv)) { $pv = $pv.Trim() }

    if (-not $Quiet) {
        Write-Host "EXE VersionInfo FileVersion   : $fv"
        Write-Host "EXE VersionInfo ProductVersion: $pv"
    }

    if ($fv -match '^\d+\.\d+\.\d+\.\d+$') { return $fv }
    if ($pv -match '^\d+\.\d+\.\d+\.\d+$') { return $pv }

    throw "Could not determine quad version from VersionInfo for: $Path"
}

function Get-ShortVersionFromQuad {
    param([Parameter(Mandatory)][string]$QuadVersion)
    $parts = $QuadVersion -split '\.'
    if ($parts.Count -lt 3) { throw "Unexpected version format: $QuadVersion" }
    return ("{0}.{1}.{2}" -f $parts[0], $parts[1], $parts[2])
}

function Convert-ShortVersionToRegExpected {
    param([Parameter(Mandatory)][string]$ShortVersion)
    return ("v{0}.00" -f $ShortVersion)
}

function Get-LatestMsvcRuntimeVersionFromLifecycleFaq {
    param([switch]$Quiet)

    if (-not $Quiet) { Write-Host "Fetching lifecycle version table from: $LifecycleFaqUrl" }
    try {
        $content = (curl.exe -L --fail --silent --show-error $LifecycleFaqUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch lifecycle page: $LifecycleFaqUrl" }

        # Expected cell values look like: 14.44.35211 or 14.50.35710
        $matches = [regex]::Matches($content, '(?<!\d)(14\.\d{2}\.\d{5})(?!\d)')
        if ($matches.Count -lt 1) {
            throw "Could not find runtime versions in lifecycle table HTML."
        }

        $versions = @()
        foreach ($m in $matches) {
            try { $versions += [version]$m.Groups[1].Value } catch {}
        }

        if ($versions.Count -lt 1) {
            throw "Could not parse lifecycle runtime versions into [version]."
        }

        $latest = ($versions | Sort-Object -Descending | Select-Object -First 1).ToString()
        if (-not $Quiet) { Write-Host "Lifecycle latest runtime version: $latest" }
        return $latest
    }
    catch {
        if (-not $Quiet) { Write-Warning "Lifecycle table lookup failed: $($_.Exception.Message)" }
        return $null
    }
}

function Get-LatestMsvcRuntimeVersionFromInstaller {
    param(
        [Parameter(Mandatory)][string]$LocalX86Path,
        [Parameter(Mandatory)][string]$LocalX64Path,
        [switch]$Quiet
    )

    $verX86Quad = Get-ExeFileVersion -Path $LocalX86Path -Quiet:$Quiet
    $verX64Quad = Get-ExeFileVersion -Path $LocalX64Path -Quiet:$Quiet

    $verX86 = Get-ShortVersionFromQuad -QuadVersion $verX86Quad
    $verX64 = Get-ShortVersionFromQuad -QuadVersion $verX64Quad

    if (-not $Quiet) {
        Write-Host "Installer-derived short version (x86): $verX86"
        Write-Host "Installer-derived short version (x64): $verX64"
    }

    $v1 = [version]$verX86
    $v2 = [version]$verX64
    $latest = ( @($v1,$v2) | Sort-Object -Descending | Select-Object -First 1 ).ToString()

    if (-not $Quiet) { Write-Host "Installer-derived short version (max): $latest" }
    return $latest
}

function Resolve-LatestMsvcVersion {
    param(
        [Parameter(Mandatory)][string]$LocalX86Path,
        [Parameter(Mandatory)][string]$LocalX64Path,
        [switch]$Quiet
    )

    $lifecycle = Get-LatestMsvcRuntimeVersionFromLifecycleFaq -Quiet:$Quiet
    $installer = Get-LatestMsvcRuntimeVersionFromInstaller -LocalX86Path $LocalX86Path -LocalX64Path $LocalX64Path -Quiet:$Quiet

    if ($null -eq $lifecycle) {
        if (-not $Quiet) { Write-Host "Using installer-derived version (lifecycle not available)." }
        return $installer
    }

    $lv = [version]$lifecycle
    $iv = [version]$installer

    if ($iv -gt $lv) {
        if (-not $Quiet) { Write-Warning "Installer version ($installer) is newer than lifecycle table ($lifecycle). Using installer-derived version." }
        return $installer
    }

    if ($iv -lt $lv) {
        if (-not $Quiet) { Write-Warning "Installer version ($installer) is older than lifecycle table ($lifecycle). Using lifecycle version." }
        return $lifecycle
    }

    if (-not $Quiet) { Write-Host "Lifecycle and installer versions match: $lifecycle" }
    return $lifecycle
}


function Download-IfMissing {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$OutFile,
        [switch]$Quiet
    )

    if (Test-Path -LiteralPath $OutFile) {
        if (-not $Quiet) { Write-Host "Local file exists. Skipping download: $OutFile" }
        return
    }

    if (-not $Quiet) {
        Write-Host "Downloading                  : $Url"
        Write-Host "To                           : $OutFile"
    }
    curl.exe -L --fail --silent --show-error -o $OutFile $Url
    if ($LASTEXITCODE -ne 0) { throw "Download failed: $Url" }
}


function Create-StaticBatchFiles {
    param([Parameter(Mandatory)][string]$TargetFolder)

    $installPath   = Join-Path $TargetFolder "install.bat"
    $uninstallPath = Join-Path $TargetFolder "uninstall.bat"

    if (-not (Test-Path -LiteralPath $installPath)) {
        $install = @'
"%~dp0vc_redist.x86.exe" /install /quiet /norestart /log "%~dp0x86.install.log"
"%~dp0vc_redist.x64.exe" /install /quiet /norestart /log "%~dp0x64.install.log"
'@
        Set-Content -LiteralPath $installPath -Value $install -Encoding ASCII -ErrorAction Stop
    }

    if (-not (Test-Path -LiteralPath $uninstallPath)) {
        $uninstall = @'
"%~dp0vc_redist.x86.exe" /uninstall /quiet /norestart /log "%~dp0x86.uninstall.log"
"%~dp0vc_redist.x64.exe" /uninstall /quiet /norestart /log "%~dp0x64.uninstall.log"
'@
        Set-Content -LiteralPath $uninstallPath -Value $uninstall -Encoding ASCII -ErrorAction Stop
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

# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    try {
        Ensure-Folder -Path $BaseDownloadRoot

        $localX86 = Join-Path $BaseDownloadRoot "vc_redist.x86.exe"
        $localX64 = Join-Path $BaseDownloadRoot "vc_redist.x64.exe"

        Download-IfMissing -Url $UrlX86 -OutFile $localX86 -Quiet
        Download-IfMissing -Url $UrlX64 -OutFile $localX64 -Quiet

        $latest = Resolve-LatestMsvcVersion -LocalX86Path $localX86 -LocalX64Path $localX64 -Quiet
        if ([string]::IsNullOrWhiteSpace($latest)) { exit 1 }

        Write-Output $latest
        exit 0
    }
    catch {
        exit 1
    }
}

# --- Main ---
$originalLocation = Get-Location

try {
    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "MSVC Redistributables Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                   : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Host "Machine                     : $env:COMPUTERNAME"
    Write-Host "Start location              : $originalLocation"
    Write-Host "SiteCode                    : $SiteCode"
    Write-Host "NetworkRootPath             : $NetworkRootPath"
    Write-Host "BaseDownloadRoot            : $BaseDownloadRoot"
    Write-Host "LifecycleFaqUrl             : $LifecycleFaqUrl"
    Write-Host "VcRedistInfoUrl             : $VcRedistInfoUrl"
    Write-Host "UrlX86                      : $UrlX86"
    Write-Host "UrlX64                      : $UrlX64"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Ensure-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $NetworkRootPath)) {
        throw "Network root path not accessible: $NetworkRootPath"
    }

    # Download latest installers (permalinks)
    $localX86 = Join-Path $BaseDownloadRoot "vc_redist.x86.exe"
    $localX64 = Join-Path $BaseDownloadRoot "vc_redist.x64.exe"

        Download-IfMissing -Url $UrlX86 -OutFile $localX86 -Quiet
        Download-IfMissing -Url $UrlX64 -OutFile $localX64 -Quiet

    # Select version (lifecycle authoritative; installer used for verification / fallback)
        $latest = Resolve-LatestMsvcVersion -LocalX86Path $localX86 -LocalX64Path $localX64 -Quiet
    $ShortVersion = $latest
    if ([string]::IsNullOrWhiteSpace($ShortVersion)) {
        throw "Could not determine latest VC++ runtime version."
    }

    $RegExpected = Convert-ShortVersionToRegExpected -ShortVersion $ShortVersion

    Write-Host ""
    Write-Host "Selected runtime version     : $ShortVersion"
    Write-Host "Registry expected value      : $RegExpected"
    Write-Host ""

    # Target content folder
    $TargetNetworkFolder = Join-Path $NetworkRootPath $ShortVersion
    Ensure-Folder -Path $TargetNetworkFolder

    # Stage files (filesystem context only)
    $netX86 = Join-Path $TargetNetworkFolder "vc_redist.x86.exe"
    $netX64 = Join-Path $TargetNetworkFolder "vc_redist.x64.exe"

    if (-not (Test-Path -LiteralPath $netX86)) {
        Copy-Item -LiteralPath $localX86 -Destination $netX86 -Force -ErrorAction Stop
        Write-Host "Staged                        : $netX86"
    }
    else {
        Write-Host "Exists (skip)                 : $netX86"
    }

    if (-not (Test-Path -LiteralPath $netX64)) {
        Copy-Item -LiteralPath $localX64 -Destination $netX64 -Force -ErrorAction Stop
        Write-Host "Staged                        : $netX64"
    }
    else {
        Write-Host "Exists (skip)                 : $netX64"
    }

    Create-StaticBatchFiles -TargetFolder $TargetNetworkFolder

    $AppName = "Microsoft Visual C++ 2015-2022 Redistributable (x86+x64) - $ShortVersion"

    Write-Host ""
    Write-Host "CM Application Name          : $AppName"
    Write-Host "Content folder              : $TargetNetworkFolder"
    Write-Host ""

    # MECM creation (site drive only during CM cmdlets)

    if (-not (Connect-CMSite -SiteCode $SiteCode)) {
        throw "CM site connection failed."
    }

    try {
        $clauseX86 = New-CMDetectionClauseRegistryKeyValue `
            -Hive LocalMachine `
            -KeyName $RegKeyX86 `
            -ValueName $RegValueName `
            -PropertyType String `
            -Value `
            -ExpectedValue $RegExpected `
            -ExpressionOperator IsEquals

        $clauseX64 = New-CMDetectionClauseRegistryKeyValue `
            -Hive LocalMachine `
            -KeyName $RegKeyX64 `
            -ValueName $RegValueName `
            -PropertyType String `
            -Value `
            -ExpectedValue $RegExpected `
            -ExpressionOperator IsEquals

        $detectionClauses = @()
        $detectionClauses += $clauseX86, $clauseX64

        $existing = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($null -ne $existing) {
            Write-Warning "Application already exists: $AppName"
            return
        }

        Write-Host "Creating CM Application      : $AppName"
        $cmApp = New-CMApplication -Name $AppName -Publisher $Publisher -SoftwareVersion $ShortVersion -Description $Comment -ErrorAction Stop

        $dtParams = @{
            ApplicationName           = $AppName
            DeploymentTypeName        = "Script Installer"
            InstallCommand            = "install.bat"
            UninstallCommand          = "uninstall.bat"
            ContentLocation           = $TargetNetworkFolder

            InstallationBehaviorType  = "InstallForSystem"
            LogonRequirementType      = "WhetherOrNotUserLoggedOn"

            EstimatedRuntimeMins      = $EstimatedRuntimeMins
            MaximumRuntimeMins        = $MaximumRuntimeMins

            ContentFallback           = $true
            SlowNetworkDeploymentMode = "Download"

            PostExecutionBehavior     = "ForceReboot"

            AddDetectionClause        = $detectionClauses

            ScriptLanguage            = "PowerShell"
            ErrorAction               = "Stop"
        }

        Write-Host "Adding Script DeploymentType : Script Installer"
        Add-CMScriptDeploymentType @dtParams | Out-Null
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host ""
        Write-Host "SUCCESS: Created MECM application + DT:"
        Write-Host "  App : $AppName"
        Write-Host "  Ver : $ShortVersion"
        Write-Host "  Det : HKLM\$RegKeyX86\$RegValueName == $RegExpected  AND  HKLM\$RegKeyX64\$RegValueName == $RegExpected"
        Write-Host "  DT  : ContentFallback=On; SlowNetwork=Download; PostExecutionBehavior=ForceReboot"
        Write-Host ""
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored location after MECM work to: $originalLocation"
    }
}
catch {
    try { Set-Location $originalLocation -ErrorAction SilentlyContinue } catch {}
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    throw
}