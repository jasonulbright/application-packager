<#
.SYNOPSIS
    Shared module for AppPackager packager scripts.

.DESCRIPTION
    Import this module at the top of every packager script to get:
      - TLS 1.2 enforcement
      - Structured logging (Write-Log, Initialize-Logging)
      - Download with retry (Invoke-DownloadWithRetry)
      - Admin check (Test-IsAdmin)
      - ConfigMgr site connection (Connect-CMSite)
      - Folder initialization (Initialize-Folder)
      - Network share access test (Test-NetworkShareAccess)
      - Content wrapper generation (Write-ContentWrappers, New-MsiWrapperContent)
      - MECM application creation (New-MECMApplicationFromManifest)
      - CM revision history cleanup (Remove-CMApplicationRevisionHistoryByCIId)

.EXAMPLE
    Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
    Initialize-Logging -LogPath $LogPath

    Write-Log "Starting packager..."
    Write-Log "Something went wrong" -Level ERROR
    Invoke-DownloadWithRetry -Url $url -OutFile $file
#>

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

$script:__AppPackagerLogPath = $null

function Initialize-Logging {
    param([string]$LogPath)

    $script:__AppPackagerLogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.

    .DESCRIPTION
        INFO  -> Write-Host (stdout)
        WARN  -> Write-Host (stdout)
        ERROR -> Write-Host (stdout) + $host.UI.WriteErrorLine (stderr)

        -Quiet suppresses all console output but still writes to the log file.
    #>
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted

        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__AppPackagerLogPath) {
        Add-Content -LiteralPath $script:__AppPackagerLogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Download with retry
# ---------------------------------------------------------------------------

function Invoke-DownloadWithRetry {
    <#
    .SYNOPSIS
        Downloads a file via curl.exe with a single retry on failure.

    .DESCRIPTION
        Wraps curl.exe file-download calls (curl.exe -L --fail --silent --show-error -o <file> <url>)
        with retry logic. Throws on final failure.

        Does NOT wrap scraping/variable-capture calls or URL-resolution calls.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Url,

        [Parameter(Mandatory)]
        [string]$OutFile,

        [string[]]$ExtraCurlArgs = @(),

        [int]$RetryCount = 1,

        [int]$RetryDelaySec = 5,

        [switch]$Quiet
    )

    $maxAttempts = 1 + $RetryCount

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if ($attempt -gt 1) {
            Write-Log ("Retrying download (attempt {0} of {1}) after {2}s delay..." -f $attempt, $maxAttempts, $RetryDelaySec) -Level WARN -Quiet:$Quiet
            Start-Sleep -Seconds $RetryDelaySec
        }

        $allArgs = @('-L', '--fail', '--silent', '--show-error') + $ExtraCurlArgs + @('-o', $OutFile, $Url)
        & curl.exe @allArgs 2>$null
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            return
        }

        if ($attempt -lt $maxAttempts) {
            Write-Log ("Download attempt {0} failed (curl exit code {1}). Will retry." -f $attempt, $exitCode) -Level WARN -Quiet:$Quiet
        }
    }

    $msg = "Download failed after $maxAttempts attempt(s): $Url"
    Write-Log $msg -Level ERROR -Quiet:$Quiet
    throw $msg
}

# ---------------------------------------------------------------------------
# TLS 1.2 enforcement
# ---------------------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------------------------------------------------------------------------
# Environment & pre-flight checks
# ---------------------------------------------------------------------------

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

# ---------------------------------------------------------------------------
# MECM helpers
# ---------------------------------------------------------------------------

function Get-MsiPropertyMap {
    param([Parameter(Mandatory)][string]$MsiPath)

    $installer = $null
    $db = $null
    $view = $null
    $record = $null

    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))

        $wanted = @("ProductName", "ProductVersion", "Manufacturer", "ProductCode")
        $map = @{}

        foreach ($p in $wanted) {
            $sql  = "SELECT `Value` FROM `Property` WHERE `Property`='$p'"
            $view = $db.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $db, @($sql))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)

            if ($null -ne $record) {
                $val = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
                $map[$p] = $val
            }
            else {
                $map[$p] = $null
            }
        }

        return $map
    }
    finally {
        foreach ($o in @($record, $view, $db, $installer)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ---------------------------------------------------------------------------
# ARP (Add/Remove Programs) registry discovery
# ---------------------------------------------------------------------------

function Find-UninstallEntry {
    <#
    .SYNOPSIS
        Searches the ARP uninstall registry keys for a product by DisplayName.

    .DESCRIPTION
        Searches both native and WOW6432Node uninstall registry paths for entries
        matching the given DisplayName pattern. Returns the registry key path
        (relative, ready for New-CMDetectionClauseRegistryKeyValue), DisplayVersion,
        Publisher, and uninstall strings.

        Supports retry/polling for installers that register asynchronously.
    #>
    param(
        [Parameter(Mandatory)][string]$DisplayNamePattern,

        [string]$ExpectedVersion,

        [int]$MaxRetries = 1,

        [int]$RetryDelaySec = 0
    )

    $uninstallRoots = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"; Is64Bit = $true },
        @{ Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; Is64Bit = $false }
    )

    $found = $null

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        if ($attempt -gt 1) {
            Write-Log ("Registry poll attempt {0}/{1} - sleeping {2}s..." -f $attempt, $MaxRetries, $RetryDelaySec) -Level WARN
            Start-Sleep -Seconds $RetryDelaySec
        }

        $candidates = @()

        foreach ($root in $uninstallRoots) {
            $keys = Get-ChildItem -Path $root.Path -ErrorAction SilentlyContinue
            if (-not $keys) { continue }

            foreach ($k in $keys) {
                $props = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
                $dn = $props.DisplayName
                if ([string]::IsNullOrWhiteSpace($dn)) { continue }

                if ($dn -like $DisplayNamePattern) {
                    $regRelative = ($root.Path -replace '^HKLM:\\', '') + '\' + $k.PSChildName

                    $candidates += [pscustomobject]@{
                        RegistryKeyRelative  = $regRelative
                        DisplayName          = $dn
                        DisplayVersion       = $props.DisplayVersion
                        Publisher            = $props.Publisher
                        UninstallString      = $props.UninstallString
                        QuietUninstallString = $props.QuietUninstallString
                        Is64Bit              = $root.Is64Bit
                    }
                }
            }
        }

        if ($candidates.Count -gt 0) {
            if ($ExpectedVersion) {
                $match = $candidates | Where-Object { $_.DisplayVersion -eq $ExpectedVersion } | Select-Object -First 1
                if ($match) { $found = $match; break }
            }

            $found = $candidates | Select-Object -First 1
            break
        }
    }

    return $found
}

# ---------------------------------------------------------------------------
# Stage manifest
# ---------------------------------------------------------------------------

function Write-StageManifest {
    <#
    .SYNOPSIS
        Writes a stage-manifest.json file.

    .DESCRIPTION
        Serializes ManifestData to JSON with schema metadata.

        Schema v2 adds optional fields for PSADT/deployment tool integration:
          InstallerType     "MSI" or "EXE"
          InstallArgs       Silent install arguments
          UninstallArgs     Silent uninstall arguments
          UninstallCommand  Full uninstall command (for EXE products)
          ProductCode       MSI ProductCode GUID (for MSI products)
          RunningProcess    Array of process names to close before install
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$ManifestData
    )

    $ManifestData['SchemaVersion'] = 2
    $ManifestData['StagedAt'] = (Get-Date -Format 'o')

    $json = $ManifestData | ConvertTo-Json -Depth 6
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8 -ErrorAction Stop
    Write-Log "Wrote stage manifest         : $Path"
}

function Read-StageManifest {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Stage manifest not found: $Path"
    }

    $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
    $manifest = $json | ConvertFrom-Json

    if (-not $manifest.SchemaVersion) {
        throw "Invalid stage manifest (missing SchemaVersion): $Path"
    }

    Write-Log "Read stage manifest          : $Path"
    return $manifest
}

# ---------------------------------------------------------------------------
# MECM helpers (continued)
# ---------------------------------------------------------------------------

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

# ---------------------------------------------------------------------------
# Network path helpers
# ---------------------------------------------------------------------------

function Get-NetworkAppRoot {
    <#
    .SYNOPSIS
        Creates and returns the network content root for an application.

    .DESCRIPTION
        Builds the path <FileServerPath>\Applications\<VendorFolder>\<AppFolder>,
        creating each level if it does not exist. Returns the final path.
    #>
    param(
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$VendorFolder,
        [Parameter(Mandatory)][string]$AppFolder
    )

    $appsRoot   = Join-Path $FileServerPath "Applications"
    $vendorPath = Join-Path $appsRoot $VendorFolder
    $appPath    = Join-Path $vendorPath $AppFolder

    Initialize-Folder -Path $appsRoot
    Initialize-Folder -Path $vendorPath
    Initialize-Folder -Path $appPath

    return $appPath
}

# ---------------------------------------------------------------------------
# Content wrapper generation
# ---------------------------------------------------------------------------

function Write-ContentWrappers {
    <#
    .SYNOPSIS
        Creates install/uninstall .bat and .ps1 wrapper files in a content folder.

    .DESCRIPTION
        Writes four files to OutputPath: install.bat, install.ps1, uninstall.bat,
        uninstall.ps1. The .bat files are thin shims that call the corresponding
        .ps1. The .ps1 content is passed as strings by the caller.

        Skips files that already exist (logs a message). All files are written
        with -Encoding ASCII to avoid BOM issues.

    .PARAMETER InstallBatExitCode
        Exit code expression for install.bat. Default: '%ERRORLEVEL%'.
        Use '3010' for products that always require reboot (e.g. VMware Tools).

    .PARAMETER UninstallBatExitCode
        Exit code expression for uninstall.bat. Default: '%ERRORLEVEL%'.
    #>
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$InstallPs1Content,
        [Parameter(Mandatory)][string]$UninstallPs1Content,
        [string]$InstallBatExitCode   = '%ERRORLEVEL%',
        [string]$UninstallBatExitCode = '%ERRORLEVEL%'
    )

    $installBatPath   = Join-Path $OutputPath "install.bat"
    $installPs1Path   = Join-Path $OutputPath "install.ps1"
    $uninstallBatPath = Join-Path $OutputPath "uninstall.bat"
    $uninstallPs1Path = Join-Path $OutputPath "uninstall.ps1"

    # .bat wrapper template: @echo off, call PowerShell, propagate exit code
    $installBat = (
        '@echo off',
        'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"',
        ('exit /b {0}' -f $InstallBatExitCode)
    ) -join "`r`n"

    $uninstallBat = (
        '@echo off',
        'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"',
        ('exit /b {0}' -f $UninstallBatExitCode)
    ) -join "`r`n"

    $files = @(
        @{ Path = $installBatPath;   Content = $installBat;          Label = 'install.bat' },
        @{ Path = $installPs1Path;   Content = $InstallPs1Content;   Label = 'install.ps1' },
        @{ Path = $uninstallBatPath; Content = $uninstallBat;        Label = 'uninstall.bat' },
        @{ Path = $uninstallPs1Path; Content = $UninstallPs1Content; Label = 'uninstall.ps1' }
    )

    foreach ($f in $files) {
        if (-not (Test-Path -LiteralPath $f.Path)) {
            Set-Content -LiteralPath $f.Path -Value $f.Content -Encoding ASCII -ErrorAction Stop
            Write-Log "Created wrapper              : $($f.Label)"
        }
        else {
            Write-Log "Wrapper exists, skipped      : $($f.Label)"
        }
    }
}

function New-MsiWrapperContent {
    <#
    .SYNOPSIS
        Returns install and uninstall .ps1 content strings for an MSI product.

    .DESCRIPTION
        Generates PowerShell script content that uses Start-Process with
        array-based ArgumentList (avoiding quoting issues). Returns a hashtable
        with Install and Uninstall keys.
    #>
    param([Parameter(Mandatory)][string]$MsiFileName)

    $install = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/i'', "`"$msiPath`"", ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstall = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/x'', "`"$msiPath`"", ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    return @{
        Install   = $install
        Uninstall = $uninstall
    }
}

function New-ExeWrapperContent {
    <#
    .SYNOPSIS
        Returns install and uninstall .ps1 content strings for an EXE product.

    .DESCRIPTION
        Generates PowerShell script content that uses Start-Process with
        array-based ArgumentList for the installer EXE. Returns a hashtable
        with Install and Uninstall keys.

        For products where uninstall uses a different command (e.g. registry
        lookup, msiexec), the caller should build uninstall content directly
        and pass it to Write-ContentWrappers.
    #>
    param(
        [Parameter(Mandatory)][string]$InstallerFileName,
        [Parameter(Mandatory)][string]$InstallArgs,
        [Parameter(Mandatory)][string]$UninstallCommand,
        [string]$UninstallArgs = ''
    )

    $install = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
        ('$proc = Start-Process -FilePath $exePath -ArgumentList @({0}) -Wait -PassThru -NoNewWindow' -f $InstallArgs),
        'exit $proc.ExitCode'
    ) -join "`r`n"

    if ($UninstallArgs -ne '') {
        $uninstall = (
            ('$proc = Start-Process -FilePath ''{0}'' -ArgumentList @({1}) -Wait -PassThru -NoNewWindow' -f $UninstallCommand, $UninstallArgs),
            'exit $proc.ExitCode'
        ) -join "`r`n"
    }
    else {
        $uninstall = (
            ('$proc = Start-Process -FilePath ''{0}'' -Wait -PassThru -NoNewWindow' -f $UninstallCommand),
            'exit $proc.ExitCode'
        ) -join "`r`n"
    }

    return @{
        Install   = $install
        Uninstall = $uninstall
    }
}

# ---------------------------------------------------------------------------
# MECM application creation from manifest
# ---------------------------------------------------------------------------

function New-SingleDetectionClause {
    <#
    .SYNOPSIS
        Builds a single CM detection clause object from a manifest detection block.
    .DESCRIPTION
        Internal helper for New-MECMApplicationFromManifest. Supports
        RegistryKeyValue, RegistryKey, and File detection types.
        Must be called while the current location is a filesystem drive
        (not the CM PSDrive).
    #>
    param([Parameter(Mandatory)][pscustomobject]$Det)

    $type = if ($Det.Type) { $Det.Type } else { 'RegistryKeyValue' }

    switch ($type) {
        'RegistryKeyValue' {
            $operator = if ($Det.Operator) { $Det.Operator } else { 'IsEquals' }
            $expected = if ($Det.ExpectedValue) { $Det.ExpectedValue } else { $Det.DisplayVersion }
            $valName  = if ($Det.ValueName) { $Det.ValueName } else { 'DisplayVersion' }
            $propType = if ($Det.PropertyType) { $Det.PropertyType } else { 'String' }

            $p = @{
                Hive               = 'LocalMachine'
                KeyName            = $Det.RegistryKeyRelative
                ValueName          = $valName
                PropertyType       = $propType
                Value              = $true
                ExpressionOperator = $operator
                ExpectedValue      = $expected
            }
            if ($null -ne $Det.Is64Bit) { $p['Is64Bit'] = [bool]$Det.Is64Bit }

            return (New-CMDetectionClauseRegistryKeyValue @p)
        }
        'RegistryKey' {
            $p = @{
                Hive      = 'LocalMachine'
                KeyName   = $Det.RegistryKeyRelative
                Existence = $true
            }
            if ($null -ne $Det.Is64Bit) { $p['Is64Bit'] = [bool]$Det.Is64Bit }

            return (New-CMDetectionClauseRegistryKey @p)
        }
        'File' {
            if ($Det.PropertyType -eq 'Existence') {
                $p = @{
                    Path      = $Det.FilePath
                    FileName  = $Det.FileName
                    Existence = $true
                }
            }
            else {
                $op = if ($Det.Operator) { $Det.Operator } else { 'GreaterEquals' }
                $p = @{
                    Path               = $Det.FilePath
                    FileName           = $Det.FileName
                    PropertyType       = $Det.PropertyType
                    Value              = $true
                    ExpressionOperator = $op
                    ExpectedValue      = $Det.ExpectedValue
                }
            }
            if ($null -ne $Det.Is64Bit) { $p['Is64Bit'] = [bool]$Det.Is64Bit }

            return (New-CMDetectionClauseFile @p)
        }
        default { throw "Unsupported detection clause type: $type" }
    }
}


function New-MECMApplicationFromManifest {
    <#
    .SYNOPSIS
        Creates an MECM application with Script deployment type from a stage manifest.

    .DESCRIPTION
        Reads a stage manifest object and creates a CM Application with a single
        Script deployment type. Supports all detection methods:

          RegistryKeyValue  Single registry value comparison (IsEquals or GreaterEquals)
          RegistryKey       Registry key existence check
          File              File existence or version comparison
          Script            PowerShell script-based detection
          Compound          Multiple clauses joined by AND or OR

        Handles CM site connection, duplicate app check, New-CMApplication with
        -AutoInstall $true, detection clause creation, Add-CMScriptDeploymentType
        with gold standard parameters, optional PostExecutionBehavior, and
        revision history cleanup.

        Backward compatible: manifests without a Detection.Type field default
        to RegistryKeyValue; DisplayVersion is accepted as an alias for
        ExpectedValue.

    .OUTPUTS
        [UInt32] CI_ID of the created application, or $null if the app already exists.
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Manifest,
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$Comment,
        [Parameter(Mandatory)][string]$NetworkContentPath,
        [int]$EstimatedRuntimeMins = 15,
        [int]$MaximumRuntimeMins = 30
    )

    $orig = Get-Location

    try {
        if (-not (Connect-CMSite -SiteCode $SiteCode)) {
            throw "CM site connection failed."
        }

        $appName = $Manifest.AppName

        $existing = Get-CMApplication -Name $appName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Log "Application already exists: $appName" -Level WARN
            return $null
        }

        Write-Log "Creating CM Application      : $appName"
        $cmAppParams = @{
            Name             = $appName
            Publisher        = $Manifest.Publisher
            SoftwareVersion  = $Manifest.SoftwareVersion
            Description      = $Comment
            AutoInstall      = $true
            ErrorAction      = 'Stop'
        }
        # Set Software Center display name if provided (omits channel/arch details)
        if ($Manifest.DisplayName) {
            $cmAppParams['LocalizedApplicationName'] = $Manifest.DisplayName
            Write-Log "Software Center name         : $($Manifest.DisplayName)"
        }
        $cmApp = New-CMApplication @cmAppParams

        Write-Log "Application CI_ID            : $($cmApp.CI_ID)"

        # Determine detection type (backward compat: missing Type = RegistryKeyValue)
        $detType = if ($Manifest.Detection.Type) { $Manifest.Detection.Type } else { 'RegistryKeyValue' }

        # Common deployment type parameters (splatted)
        $dtName = $appName
        $dtParams = @{
            ApplicationName           = $appName
            DeploymentTypeName        = $dtName
            ContentLocation           = $NetworkContentPath
            InstallCommand            = 'install.bat'
            UninstallCommand          = 'uninstall.bat'
            InstallationBehaviorType  = 'InstallForSystem'
            LogonRequirementType      = 'WhetherOrNotUserLoggedOn'
            EstimatedRuntimeMins      = $EstimatedRuntimeMins
            MaximumRuntimeMins        = $MaximumRuntimeMins
            ContentFallback           = $true
            SlowNetworkDeploymentMode = 'Download'
            ErrorAction               = 'Stop'
        }

        if ($Manifest.PostExecutionBehavior) {
            $dtParams['PostExecutionBehavior'] = $Manifest.PostExecutionBehavior
        }

        if ($Manifest.InstallationBehaviorType) {
            $dtParams['InstallationBehaviorType'] = $Manifest.InstallationBehaviorType
        }
        if ($Manifest.LogonRequirementType) {
            $dtParams['LogonRequirementType'] = $Manifest.LogonRequirementType
        }
        if ($Manifest.RequireUserInteraction -eq $true) {
            $dtParams['RequireUserInteraction'] = $true
        }

        if ($detType -eq 'Script') {
            # Script-based detection: pass script text, no clause objects needed
            $lang = if ($Manifest.Detection.ScriptLanguage) { $Manifest.Detection.ScriptLanguage } else { 'PowerShell' }
            $dtParams['ScriptLanguage'] = $lang
            $dtParams['ScriptText']     = $Manifest.Detection.ScriptText
        }
        else {
            # Clause-based detection: leave CM PSDrive to create clause objects
            # (CM PSDrive context can interfere with parameter binding)
            Set-Location C: -ErrorAction Stop

            if ($detType -eq 'Compound') {
                $clauses = @()
                foreach ($c in $Manifest.Detection.Clauses) {
                    $clauses += New-SingleDetectionClause -Det $c
                }
                $dtParams['AddDetectionClause'] = $clauses

                # OR connector: specify OR for each clause beyond the first
                # AND is the default and needs no explicit connector
                if ($Manifest.Detection.Connector -eq 'Or' -and $clauses.Count -ge 2) {
                    $connectors = @()
                    for ($i = 1; $i -lt $clauses.Count; $i++) {
                        $connectors += @{
                            LogicalName = $clauses[$i].Setting.LogicalName
                            Connector   = 'OR'
                        }
                    }
                    $dtParams['DetectionClauseConnector'] = $connectors
                }
            }
            else {
                # Single clause: RegistryKeyValue, RegistryKey, or File
                $clause = New-SingleDetectionClause -Det $Manifest.Detection
                $dtParams['AddDetectionClause'] = @($clause)
            }

            # Reconnect to CM site for Add-CMScriptDeploymentType
            if (-not (Connect-CMSite -SiteCode $SiteCode)) {
                throw "CM site reconnection failed."
            }
        }

        Write-Log "Adding Script Deployment Type : $dtName"
        Add-CMScriptDeploymentType @dtParams | Out-Null

        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Log ""
        Write-Log "Created MECM application     : $appName"

        return [UInt32]$cmApp.CI_ID
    }
    finally {
        Set-Location $orig -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Packager preferences
# ---------------------------------------------------------------------------

function Get-PackagerPreferences {
    <#
    .SYNOPSIS
        Reads packager-preferences.json from the Packagers folder.
    #>
    $prefsPath = Join-Path $PSScriptRoot "packager-preferences.json"
    if (-not (Test-Path -LiteralPath $prefsPath)) {
        Write-Log "Preferences file not found: $prefsPath" -Level WARN
        return $null
    }
    $json = Get-Content -LiteralPath $prefsPath -Raw -Encoding UTF8 -ErrorAction Stop
    return ($json | ConvertFrom-Json)
}

# ---------------------------------------------------------------------------
# ODT config XML generation
# ---------------------------------------------------------------------------

function New-OdtConfigXml {
    <#
    .SYNOPSIS
        Generates a full ODT configuration XML string for download or install.

    .DESCRIPTION
        Builds the XML matching the production ODT template with all properties,
        excluded apps, AppSettings, logging, etc. Used by all M365 packager
        scripts for both download.xml and install.xml.

    .PARAMETER OfficeClientEdition
        Architecture: "32" or "64".

    .PARAMETER Version
        Full M365 version string (e.g. "16.0.19127.20532").

    .PARAMETER ProductIds
        Array of product IDs (e.g. @('O365ProPlusRetail') or
        @('O365ProPlusRetail', 'VisioProRetail')).

    .PARAMETER SourcePath
        SourcePath attribute for the Add element. For download: local content
        folder path. For install: ".".

    .PARAMETER Channel
        ODT channel name. Valid values: MonthlyEnterprise, Current.
        Default: MonthlyEnterprise.

    .PARAMETER CompanyName
        Value for the AppSettings Company name. Omit or pass empty to skip
        the AppSettings block entirely.
    #>
    param(
        [Parameter(Mandatory)][string]$OfficeClientEdition,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string[]]$ProductIds,
        [string]$SourcePath,
        [ValidateSet('MonthlyEnterprise','Current')]
        [string]$Channel = 'MonthlyEnterprise',
        [string]$CompanyName
    )

    $addAttrs = @()
    if ($SourcePath) {
        $addAttrs += 'SourcePath="{0}"' -f $SourcePath
    }
    $addAttrs += 'OfficeClientEdition="{0}"' -f $OfficeClientEdition
    $addAttrs += 'Channel="{0}"' -f $Channel
    $addAttrs += 'OfficeMgmtCOM="TRUE"'
    $addAttrs += 'Version="{0}"' -f $Version
    $addAttrs += 'MigrateArch="TRUE"'

    $lines = @('<Configuration>')
    $lines += '  <Add {0}>' -f ($addAttrs -join ' ')

    foreach ($prodId in $ProductIds) {
        $lines += '    <Product ID="{0}">' -f $prodId
        $lines += '      <Language ID="en-us" />'
        $lines += '      <ExcludeApp ID="Groove" />'
        $lines += '      <ExcludeApp ID="Lync" />'
        $lines += '      <ExcludeApp ID="OneDrive" />'
        $lines += '      <ExcludeApp ID="Teams" />'
        $lines += '      <ExcludeApp ID="Bing" />'
        $lines += '    </Product>'
    }

    $lines += '  </Add>'
    $lines += '  <Property Name="SharedComputerLicensing" Value="1" />'
    $lines += '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
    $lines += '  <Property Name="DeviceBasedLicensing" Value="0" />'
    $lines += '  <Property Name="PinIconsToTaskbar" Value="FALSE" />'
    $lines += '  <Property Name="SCLCacheOverride" Value="0" />'
    $lines += '  <RemoveMSI />'
    if ($CompanyName) {
        $lines += '  <AppSettings>'
        $lines += '    <Setup Name="Company" Value="{0}" />' -f $CompanyName
        $lines += '  </AppSettings>'
    }
    $lines += '  <Display Level="None" AcceptEULA="TRUE" />'
    $lines += '  <Logging Level="Standard" Path="%programdata%\Appdeploy\Office2016" />'
    $lines += '</Configuration>'

    return ($lines -join "`r`n")
}

# ---------------------------------------------------------------------------
# Java vendor release helpers
# ---------------------------------------------------------------------------

function Get-LatestTemurinRelease {
    <#
    .SYNOPSIS
        Queries the Eclipse Adoptium API for the latest Temurin release.
    .DESCRIPTION
        Returns a hashtable with Version, DownloadUrl, and FileName for the
        latest Temurin JRE or JDK MSI installer. The -LTS suffix is stripped
        from the version string.
    .PARAMETER FeatureVersion
        Major Java version (8, 11, 17, 21, 25).
    .PARAMETER ImageType
        'jre' or 'jdk'.
    .PARAMETER Architecture
        'x64' or 'x86'. Defaults to 'x64'.
    .PARAMETER Quiet
        Suppress log output (for GetLatestVersionOnly mode).
    #>
    param(
        [Parameter(Mandatory)][int]$FeatureVersion,
        [Parameter(Mandatory)][ValidateSet('jre','jdk')][string]$ImageType,
        [ValidateSet('x64','x86')][string]$Architecture = 'x64',
        [switch]$Quiet
    )

    $apiUrl = "https://api.adoptium.net/v3/assets/latest/$FeatureVersion/hotspot?architecture=$Architecture&image_type=$ImageType&os=windows"
    Write-Log "Adoptium API URL             : $apiUrl" -Quiet:$Quiet

    try {
        $json = (& curl.exe -L --fail --silent --show-error $apiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Adoptium API." }

        $data = ConvertFrom-Json $json

        $asset = $data | Where-Object { $_.binary.installer.name -match '\.msi$' } | Select-Object -First 1
        if (-not $asset) { throw "No MSI installer found for Temurin $ImageType $FeatureVersion ($Architecture)." }

        $downloadUrl = $asset.binary.installer.link
        $fileName    = $asset.binary.installer.name
        $rawVersion  = $asset.version.semver

        if ([string]::IsNullOrWhiteSpace($rawVersion)) { throw "version.semver is empty in Adoptium API response." }

        $version = $rawVersion -replace '[\.\-]\d*\.?LTS$', ''

        Write-Log ("Temurin {0} {1} version      : {2}" -f $ImageType, $FeatureVersion, $version) -Quiet:$Quiet

        return @{
            Version     = $version
            DownloadUrl = $downloadUrl
            FileName    = $fileName
        }
    }
    catch {
        Write-Log ("Failed to get Temurin release: {0}" -f $_.Exception.Message) -Level ERROR
        return $null
    }
}


function Get-LatestCorrettoRelease {
    <#
    .SYNOPSIS
        Queries the GitHub API for the latest Amazon Corretto JDK release.
    .DESCRIPTION
        Returns a hashtable with Version (4-part normalized), DownloadUrl, and
        FileName. Corretto uses 5-part versioning; the 5th part (Corretto patch)
        is stripped to produce a 4-part version compatible with the GUI regex
        and the .NET [version] type.
    .PARAMETER FeatureVersion
        Major Java version (8, 11, 17, 21, 25).
    .PARAMETER Architecture
        'x64' or 'x86'. Defaults to 'x64'.
    .PARAMETER Quiet
        Suppress log output (for GetLatestVersionOnly mode).
    #>
    param(
        [Parameter(Mandatory)][int]$FeatureVersion,
        [ValidateSet('x64','x86')][string]$Architecture = 'x64',
        [switch]$Quiet
    )

    $apiUrl = "https://api.github.com/repos/corretto/corretto-$FeatureVersion/releases/latest"
    Write-Log "Corretto GitHub API URL      : $apiUrl" -Quiet:$Quiet

    try {
        $json = (& curl.exe -L --fail --silent --show-error -A "PowerShell" $apiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Corretto GitHub API." }

        $release = ConvertFrom-Json $json

        $tagVersion = $release.tag_name
        if ([string]::IsNullOrWhiteSpace($tagVersion)) { throw "tag_name is empty in Corretto release response." }

        # Construct MSI filename from known pattern
        # v8: amazon-corretto-{TAG}-windows-{ARCH}-jdk.msi
        # v11+: amazon-corretto-{TAG}-windows-{ARCH}.msi
        if ($FeatureVersion -le 8) {
            $fileName = "amazon-corretto-$tagVersion-windows-$Architecture-jdk.msi"
        }
        else {
            $fileName = "amazon-corretto-$tagVersion-windows-$Architecture.msi"
        }
        $downloadUrl = "https://corretto.aws/downloads/resources/$tagVersion/$fileName"

        # Normalize 5-part version to 4 parts (strip Corretto patch)
        $parts = $tagVersion -split '\.'
        if ($parts.Count -ge 5) {
            $version = ($parts[0..3] -join '.')
        }
        else {
            $version = $tagVersion
        }

        Write-Log ("Corretto {0} raw version     : {1}" -f $FeatureVersion, $tagVersion) -Quiet:$Quiet
        Write-Log ("Corretto {0} normalized      : {1}" -f $FeatureVersion, $version) -Quiet:$Quiet

        return @{
            Version     = $version
            DownloadUrl = $downloadUrl
            FileName    = $fileName
        }
    }
    catch {
        Write-Log ("Failed to get Corretto release: {0}" -f $_.Exception.Message) -Level ERROR
        return $null
    }
}

# ---------------------------------------------------------------------------
# Module export (belt-and-suspenders with .psd1 FunctionsToExport)
# ---------------------------------------------------------------------------

Export-ModuleMember -Function *
