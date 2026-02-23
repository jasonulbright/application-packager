<#
.SYNOPSIS
    WinForms front-end for application packager scripts (metadata-driven, no network on launch).

.DESCRIPTION
    Displays locally-discovered packager scripts in a DataGridView with checkboxes and status indicators.
    On launch, the tool performs LOCAL-ONLY operations:
      - Enumerates packager scripts in the PackagersRoot folder
      - Parses metadata tags from each script header:
          # Vendor:
          # App:
          # CMName:   (optional; defaults to App)
      - Populates the grid with Vendor/Application and placeholders for Current/Latest/Status

    No network operations are performed on launch. Network operations are only performed after explicit user action:
      - Check Latest: runs selected packagers with -GetLatestVersionOnly and updates Latest/Status
      - Check MECM: queries MECM for selected products (requires existing CM PSDrive session)
      - Stage Packages: downloads installers, discovers ARP metadata, writes stage manifests
      - Package Apps: reads stage manifests, copies content to network, creates MECM applications

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM"). The PSDrive is assumed to already exist in the session.
    Used only when MECM actions are invoked by the user.

.PARAMETER PackagersRoot
    Local folder containing packager scripts (e.g., .\Packagers).
    Scripts supported:
      - package-*.ps1
      - package-*.notps1   (treated as PowerShell script content; used for email/security systems)

.EXAMPLE
    .\start-apppackager.ps1

.EXAMPLE
    .\start-apppackager.ps1 -SiteCode "MCM" -PackagersRoot "D:\CM\Packagers"

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8.2
      - Windows Forms (System.Windows.Forms)

    Startup behavior:
      - No MECM queries on launch
      - No internet queries on launch
      - No network share access on launch

    ScriptName : start-apppackager.ps1
    Purpose    : WinForms front-end for packager scripts (metadata-driven selection + actions)
    Owner      : CM Engineering
    Version    : 0.3.0
    Updated    : 2026-02-22
#>


param(
    [string]$SiteCode = "MCM",
    [string]$PackagersRoot = (Join-Path $PSScriptRoot "Packagers")
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }
# -----------------------------
# Helpers
# -----------------------------
function Get-ConfigStorePath {
    Join-Path $PSScriptRoot "AppPackager.configurations.json"
}

function Read-ConfigStore {
    $path = Get-ConfigStorePath
    if (-not (Test-Path -LiteralPath $path)) { return @() }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $data = $raw | ConvertFrom-Json -ErrorAction Stop

        if ($data -is [System.Collections.IEnumerable]) { return @($data) }
        return @($data)
    }
    catch {
        return @()
    }
}

function Write-ConfigStore {
    param([Parameter(Mandatory)][object[]]$Configs)

    $path = Get-ConfigStorePath
    $json = ($Configs | ConvertTo-Json -Depth 6)
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
}

function Set-ConfigurationInputs {
    param(
        [Parameter(Mandatory)][pscustomobject]$Config,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TxtComment,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TxtFSPath,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TxtSiteCode,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TxtDownloadRoot,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TxtEstRuntime,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TxtMaxRuntime
    )

    if ($null -ne $Config.WOComment)            { $TxtComment.Text      = [string]$Config.WOComment }
    if ($null -ne $Config.FileShareRoot)         { $TxtFSPath.Text       = [string]$Config.FileShareRoot }
    if ($null -ne $Config.SiteCode)              { $TxtSiteCode.Text     = [string]$Config.SiteCode }
    if ($null -ne $Config.DownloadRoot)           { $TxtDownloadRoot.Text = [string]$Config.DownloadRoot }
    if ($null -ne $Config.EstimatedRuntimeMins)   { $TxtEstRuntime.Text   = [string]$Config.EstimatedRuntimeMins }
    if ($null -ne $Config.MaximumRuntimeMins)     { $TxtMaxRuntime.Text   = [string]$Config.MaximumRuntimeMins }
}

function Save-Configuration {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$WOComment,
        [Parameter(Mandatory)][string]$FileShareRoot,
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$DownloadRoot,
        [Parameter(Mandatory)][int]$EstimatedRuntimeMins,
        [Parameter(Mandatory)][int]$MaximumRuntimeMins
    )

    $configs = @(Read-ConfigStore)

    $existing = $configs | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
    if ($existing) {
        $existing.WOComment            = $WOComment
        $existing.FileShareRoot        = $FileShareRoot
        $existing.SiteCode             = $SiteCode
        $existing.DownloadRoot          = $DownloadRoot
        $existing.EstimatedRuntimeMins  = $EstimatedRuntimeMins
        $existing.MaximumRuntimeMins    = $MaximumRuntimeMins
    }
    else {
        $configs += [pscustomobject]@{
            Name                  = $Name
            WOComment             = $WOComment
            FileShareRoot         = $FileShareRoot
            SiteCode              = $SiteCode
            DownloadRoot           = $DownloadRoot
            EstimatedRuntimeMins   = $EstimatedRuntimeMins
            MaximumRuntimeMins     = $MaximumRuntimeMins
        }
    }

    Write-ConfigStore -Configs $configs
    return $configs
}

function Select-OnlyUpdateAvailable {
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable
    )

    foreach ($row in $DataTable.Rows) {
        $row["Selected"] = ($row["Status"] -eq "Update available")
    }
}

function Get-WindowStatePath {
    Join-Path $PSScriptRoot "AppPackager.windowstate.json"
}

function Save-WindowState {
    param([Parameter(Mandatory)][System.Windows.Forms.Form]$Form)

    $state = @{}

    if ($Form.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
        $state.Left   = $Form.Left
        $state.Top    = $Form.Top
        $state.Width  = $Form.Width
        $state.Height = $Form.Height
    }
    else {
        $state.Left   = $Form.RestoreBounds.Left
        $state.Top    = $Form.RestoreBounds.Top
        $state.Width  = $Form.RestoreBounds.Width
        $state.Height = $Form.RestoreBounds.Height
    }
    $state.Maximized = ($Form.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)

    try {
        $json = $state | ConvertTo-Json
        Set-Content -LiteralPath (Get-WindowStatePath) -Value $json -Encoding UTF8
    }
    catch { }
}

function Restore-WindowState {
    param([Parameter(Mandatory)][System.Windows.Forms.Form]$Form)

    $path = Get-WindowStatePath
    if (-not (Test-Path -LiteralPath $path)) { return }

    try {
        $state = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

        $pt = New-Object System.Drawing.Point ([int]$state.Left), ([int]$state.Top)
        $screen = [System.Windows.Forms.Screen]::FromPoint($pt)
        $wa = $screen.WorkingArea

        $left   = [Math]::Max($wa.Left, [Math]::Min([int]$state.Left, ($wa.Right - 200)))
        $top    = [Math]::Max($wa.Top,  [Math]::Min([int]$state.Top,  ($wa.Bottom - 100)))
        $width  = [Math]::Max($Form.MinimumSize.Width,  [Math]::Min([int]$state.Width,  $wa.Width))
        $height = [Math]::Max($Form.MinimumSize.Height, [Math]::Min([int]$state.Height, $wa.Height))

        $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
        $Form.Location = New-Object System.Drawing.Point $left, $top
        $Form.Size     = New-Object System.Drawing.Size $width, $height

        if ($state.Maximized -eq $true) {
            $Form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
        }
    }
    catch { }
}

function New-Mdl2IconBitmap {
    param(
        [Parameter(Mandatory)][int]$CodePoint,
        [int]$Size = 20
    )

    $bmp = New-Object System.Drawing.Bitmap $Size, $Size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $g.Clear([System.Drawing.Color]::Transparent)

    $font = New-Object System.Drawing.Font("Segoe MDL2 Assets", ($Size - 4), [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60, 60, 60))

    $glyph = [char]$CodePoint
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center

    $rect = New-Object System.Drawing.RectangleF 0, 0, $Size, $Size
    $g.DrawString($glyph, $font, $brush, $rect, $sf)

    $brush.Dispose()
    $font.Dispose()
    $sf.Dispose()
    $g.Dispose()

    return $bmp
}
function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory=$true)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory=$true)][System.Drawing.Color]$BackColor
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand

    $hover = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 18),
        [Math]::Max(0, $BackColor.G - 18),
        [Math]::Max(0, $BackColor.B - 18)
    )
    $down = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 36),
        [Math]::Max(0, $BackColor.G - 36),
        [Math]::Max(0, $BackColor.B - 36)
    )

    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}


function Enable-DoubleBuffer {
    param([Parameter(Mandatory=$true)][System.Windows.Forms.Control]$Control)

    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function Add-LogLine {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox,
        [Parameter(Mandatory)][string]$Message
    )
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message

    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) {
        $TextBox.Text = $line
    }
    else {
        $TextBox.AppendText([Environment]::NewLine + $line)
    }

    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
}

function Get-PackagerMetadata {
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $meta = [ordered]@{
        Vendor = $null
        App    = $null
        CMName = $null
    }

    $lines = Get-Content -LiteralPath $Path -TotalCount 200 -ErrorAction Stop

    foreach ($line in $lines) {
        $l = $line.TrimStart([char]0xFEFF)

        if (-not $meta.Vendor -and $l -match '^\s*(?:#\s*)?Vendor\s*:\s*(.+?)\s*$') { $meta.Vendor = $Matches[1].Trim(); continue }
        if (-not $meta.App    -and $l -match '^\s*(?:#\s*)?App\s*:\s*(.+?)\s*$')    { $meta.App    = $Matches[1].Trim(); continue }
        if (-not $meta.CMName -and $l -match '^\s*(?:#\s*)?CMName\s*:\s*(.+?)\s*$') { $meta.CMName = $Matches[1].Trim(); continue }

        if (-not $meta.App    -and $l -match '^\s*(?:#\s*)?Application\s*:\s*(.+?)\s*$') { $meta.App = $Matches[1].Trim(); continue }
    }

    if (-not $meta.CMName) { $meta.CMName = $meta.App }

    return [pscustomobject]@{
        Vendor      = $meta.Vendor
        Application = $meta.App
        CMName      = $meta.CMName
        Script      = (Split-Path -Leaf $Path)
        FullPath    = $Path
    }
}

function Get-Packagers {
    param(
        [Parameter(Mandatory)][string]$Root
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        return @()
    }

    $files = Get-ChildItem -LiteralPath $Root -File -ErrorAction Stop |
        Where-Object { $_.Name -match '^package-.*\.(?:ps1|notps1)$' } |
        Sort-Object Name

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($f in $files) {
        try {
            $m = Get-PackagerMetadata -Path $f.FullName

            $status = "Ready"
            if ($f.Extension -ieq ".notps1") {
                $status = "Not runnable (.notps1)"
            }
            if (-not $m.Vendor -or -not $m.Application) {
                $status = "Missing metadata (Vendor/App)"
            }

            $items.Add([pscustomobject]@{
                Selected      = $false
                Vendor        = $m.Vendor
                Application   = $m.Application
                CMName        = $m.CMName
                Script        = $m.Script
                FullPath      = $m.FullPath
                CurrentVersion= ""
                LatestVersion = ""
                Status        = $status
            })
        }
        catch {
            $items.Add([pscustomobject]@{
                Selected      = $false
                Vendor        = ""
                Application   = ""
                CMName        = ""
                Script        = $f.Name
                FullPath      = $f.FullName
                CurrentVersion= ""
                LatestVersion = ""
                Status        = ("Read error: " + $_.Exception.Message)
            })
        }
    }
    return $items
}

function Test-PackagerSupportsFileServerPath {
    param([Parameter(Mandatory)][string]$PackagerPath)

    try {
        $head = Get-Content -LiteralPath $PackagerPath -TotalCount 120 -ErrorAction Stop | Out-String
        return ($head -match '\$FileServerPath')
    }
    catch {
        return $false
    }
}

function Get-PackagerFolderInfo {
    param([Parameter(Mandatory)][string]$ScriptPath)

    $info = @{ DownloadSubfolder = $null; VendorFolder = $null; AppFolder = $null }
    try {
        $lines = Get-Content -LiteralPath $ScriptPath -TotalCount 120 -ErrorAction Stop
        foreach ($line in $lines) {
            if (-not $info.DownloadSubfolder -and $line -match '\$BaseDownloadRoot\s*=\s*Join-Path\s+\$DownloadRoot\s+"([^"]+)"') {
                $info.DownloadSubfolder = $matches[1]
            }
            if (-not $info.VendorFolder -and $line -match '^\s*\$VendorFolder\s*=\s*"([^"]+)"') {
                $info.VendorFolder = $matches[1]
            }
            if (-not $info.AppFolder -and $line -match '^\s*\$AppFolder\s*=\s*"([^"]+)"') {
                $info.AppFolder = $matches[1]
            }
        }
    }
    catch { }

    return $info
}

function Invoke-PackagerGetLatestVersion {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$FileServerPath = $null,
        [string]$DownloadRoot = $null
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $argsBase = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -SiteCode "{1}" -GetLatestVersionOnly' -f $PackagerPath, $SiteCode)
    if ($FileServerPath -and (Test-PackagerSupportsFileServerPath -PackagerPath $PackagerPath)) {
        $argsBase = ($argsBase + (' -FileServerPath "{0}"' -f $FileServerPath))
    }
    if ($DownloadRoot) {
        $argsBase = ($argsBase + (' -DownloadRoot "{0}"' -f $DownloadRoot))
    }
    $psi.Arguments = $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi

    $null = $p.Start()
    $stderr = $p.StandardError.ReadToEnd()
    $stdout = $p.StandardOutput.ReadToEnd()
    $p.WaitForExit()

    if ($p.ExitCode -ne 0) {
        $msg = $stderr
        if ([string]::IsNullOrWhiteSpace($msg)) { $msg = $stdout }
        if ([string]::IsNullOrWhiteSpace($msg)) { $msg = "Packager returned exit code $($p.ExitCode)." }
        throw $msg.Trim()
    }

    $lines = @($stdout -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not $lines -or $lines.Count -lt 1) {
        throw "No version output received."
    }

    $version = ([string]$lines[0]).Trim()

    if ($version -notmatch '^\d+(\.\d+){1,3}$') {
        throw ("Unexpected version string: '{0}'" -f $version)
    }

    return $version
}

function Compare-SemVer {
    param(
        [Parameter(Mandatory)][string]$A,
        [Parameter(Mandatory)][string]$B
    )
    try {
        $va = [version]$A
        $vb = [version]$B
        return $va.CompareTo($vb)
    }
    catch {
        return 0
    }
}

function Get-MecmCurrentVersionByCMName {
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$CMName
    )

    if (-not (Get-Command -Name Get-CMApplication -ErrorAction SilentlyContinue)) {
        try {
            if ($env:SMS_ADMIN_UI_PATH) {
                $cmModule = Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) "ConfigurationManager.psd1"
                if (Test-Path -LiteralPath $cmModule) {
                    Import-Module $cmModule -Force -ErrorAction Stop
                }
            }
        } catch { }
    }
    if (-not (Get-Command -Name Get-CMApplication -ErrorAction SilentlyContinue)) {
        throw "ConfigMgr PowerShell cmdlets not available in this session."
    }

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
    }
    catch {
        throw ("Failed to connect to CM site PSDrive '{0}:'" -f $SiteCode)
    }

    $apps = @(Get-CMApplication -Name $CMName -ErrorAction SilentlyContinue)

    if (-not $apps -or $apps.Count -eq 0) {
        $apps = @(Get-CMApplication -Name ("{0}*" -f $CMName) -ErrorAction SilentlyContinue)
    }

    if (-not $apps -or $apps.Count -eq 0) {
        return [pscustomobject]@{
            Found          = $false
            DisplayName    = $null
            SoftwareVersion= $null
            MatchCount     = 0
        }
    }

    $exact = $apps | Where-Object { $_.LocalizedDisplayName -eq $CMName -or $_.Name -eq $CMName }
    if ($exact -and $exact.Count -gt 0) {
        $chosen = $exact | Select-Object -First 1
    }
    else {
        $parsable = @()
        $nonParsable = @()

        foreach ($a in $apps) {
            try {
                $null = [version]$a.SoftwareVersion
                $parsable += $a
            }
            catch {
                $nonParsable += $a
            }
        }

        if ($parsable.Count -gt 0) {
            $chosen = $parsable | Sort-Object { [version]$_.SoftwareVersion } -Descending | Select-Object -First 1
        }
        else {
            $chosen = $nonParsable | Sort-Object Name -Descending | Select-Object -First 1
        }
    }

    return [pscustomobject]@{
        Found           = $true
        DisplayName     = $chosen.LocalizedDisplayName
        SoftwareVersion = $chosen.SoftwareVersion
        MatchCount      = $apps.Count
    }
}

function Invoke-ProcessWithStreaming {
    param(
        [Parameter(Mandatory)][System.Diagnostics.ProcessStartInfo]$StartInfo,
        [Parameter(Mandatory)][string]$OutLog,
        [Parameter(Mandatory)][string]$ErrLog,
        [string]$StructuredLog = '',
        [System.Windows.Forms.TextBox]$LogTextBox = $null
    )

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $StartInfo

    $null = $p.Start()

    $outLines = New-Object System.Collections.Generic.List[string]

    # Read stderr asynchronously (collected, not streamed)
    $errTask = $p.StandardError.ReadToEndAsync()

    # Read stdout line-by-line for real-time display
    $reader   = $p.StandardOutput
    $lineTask = $reader.ReadLineAsync()

    while ($true) {
        if ($lineTask.IsCompleted) {
            $line = $lineTask.Result
            if ($null -eq $line) { break }

            $outLines.Add($line)

            if ($LogTextBox) {
                $displayLine = $line -replace '^\[[\d: -]+\] \[\w+\s*\] ', ''
                if ($displayLine.Trim()) {
                    Add-LogLine -TextBox $LogTextBox -Message ("  {0}" -f $displayLine)
                }
            }

            $lineTask = $reader.ReadLineAsync()
        }

        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 50
    }

    $p.WaitForExit()

    $stdout = ($outLines -join "`r`n")
    $stderr = $errTask.Result

    Set-Content -LiteralPath $OutLog -Value $stdout -Encoding UTF8
    Set-Content -LiteralPath $ErrLog -Value $stderr -Encoding UTF8

    return [pscustomobject]@{
        ExitCode      = $p.ExitCode
        OutLog        = $OutLog
        ErrLog        = $ErrLog
        StructuredLog = $StructuredLog
        StdErr        = $stderr
    }
}

function Invoke-PackagerRun {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$Comment,
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$LogFolder,
        [string]$DownloadRoot = $null,
        [int]$EstimatedRuntimeMins = 0,
        [int]$MaximumRuntimeMins = 0,
        [System.Windows.Forms.TextBox]$LogTextBox = $null
    )

    if (-not (Test-Path -LiteralPath $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
    }

    $stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $base  = [IO.Path]::GetFileNameWithoutExtension($PackagerPath)
    $outLog         = Join-Path $LogFolder ("{0}-{1}.out.log" -f $base, $stamp)
    $errLog         = Join-Path $LogFolder ("{0}-{1}.err.log" -f $base, $stamp)
    $structuredLog  = Join-Path $LogFolder ("{0}-{1}.structured.log" -f $base, $stamp)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $argsBase = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -SiteCode "{1}" -Comment "{2}" -LogPath "{3}"' -f $PackagerPath, $SiteCode, $Comment, $structuredLog)
    if (Test-PackagerSupportsFileServerPath -PackagerPath $PackagerPath) {
        $argsBase = ($argsBase + (' -FileServerPath "{0}"' -f $FileServerPath))
    }
    if ($DownloadRoot) {
        $argsBase = ($argsBase + (' -DownloadRoot "{0}"' -f $DownloadRoot))
    }
    if ($EstimatedRuntimeMins -gt 0) {
        $argsBase = ($argsBase + (' -EstimatedRuntimeMins {0}' -f $EstimatedRuntimeMins))
    }
    if ($MaximumRuntimeMins -gt 0) {
        $argsBase = ($argsBase + (' -MaximumRuntimeMins {0}' -f $MaximumRuntimeMins))
    }
    $psi.Arguments = $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true

    return Invoke-ProcessWithStreaming -StartInfo $psi -OutLog $outLog -ErrLog $errLog -StructuredLog $structuredLog -LogTextBox $LogTextBox
}

function Invoke-PackagerStage {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$LogFolder,
        [string]$DownloadRoot = $null,
        [System.Windows.Forms.TextBox]$LogTextBox = $null
    )

    if (-not (Test-Path -LiteralPath $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
    }

    $stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $base  = [IO.Path]::GetFileNameWithoutExtension($PackagerPath)
    $outLog         = Join-Path $LogFolder ("{0}-stage-{1}.out.log" -f $base, $stamp)
    $errLog         = Join-Path $LogFolder ("{0}-stage-{1}.err.log" -f $base, $stamp)
    $structuredLog  = Join-Path $LogFolder ("{0}-stage-{1}.structured.log" -f $base, $stamp)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $argsBase = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -StageOnly -LogPath "{1}"' -f $PackagerPath, $structuredLog)
    if ($DownloadRoot) {
        $argsBase = ($argsBase + (' -DownloadRoot "{0}"' -f $DownloadRoot))
    }
    $psi.Arguments = $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true

    return Invoke-ProcessWithStreaming -StartInfo $psi -OutLog $outLog -ErrLog $errLog -StructuredLog $structuredLog -LogTextBox $LogTextBox
}

function Invoke-PackagerPackage {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$Comment,
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$LogFolder,
        [string]$DownloadRoot = $null,
        [int]$EstimatedRuntimeMins = 0,
        [int]$MaximumRuntimeMins = 0,
        [System.Windows.Forms.TextBox]$LogTextBox = $null
    )

    if (-not (Test-Path -LiteralPath $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
    }

    $stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $base  = [IO.Path]::GetFileNameWithoutExtension($PackagerPath)
    $outLog         = Join-Path $LogFolder ("{0}-package-{1}.out.log" -f $base, $stamp)
    $errLog         = Join-Path $LogFolder ("{0}-package-{1}.err.log" -f $base, $stamp)
    $structuredLog  = Join-Path $LogFolder ("{0}-package-{1}.structured.log" -f $base, $stamp)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $argsBase = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -PackageOnly -SiteCode "{1}" -Comment "{2}" -LogPath "{3}"' -f $PackagerPath, $SiteCode, $Comment, $structuredLog)
    if (Test-PackagerSupportsFileServerPath -PackagerPath $PackagerPath) {
        $argsBase = ($argsBase + (' -FileServerPath "{0}"' -f $FileServerPath))
    }
    if ($DownloadRoot) {
        $argsBase = ($argsBase + (' -DownloadRoot "{0}"' -f $DownloadRoot))
    }
    if ($EstimatedRuntimeMins -gt 0) {
        $argsBase = ($argsBase + (' -EstimatedRuntimeMins {0}' -f $EstimatedRuntimeMins))
    }
    if ($MaximumRuntimeMins -gt 0) {
        $argsBase = ($argsBase + (' -MaximumRuntimeMins {0}' -f $MaximumRuntimeMins))
    }
    $psi.Arguments = $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true

    return Invoke-ProcessWithStreaming -StartInfo $psi -OutLog $outLog -ErrLog $errLog -StructuredLog $structuredLog -LogTextBox $LogTextBox
}

# -----------------------------
# UI
# -----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "AppPackager"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1120, 720)
$form.MinimumSize = New-Object System.Drawing.Size(980, 640)

$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.BackColor = [System.Drawing.Color]::White
$icoPath = Join-Path $PSScriptRoot "apppackager.ico"
$logoPath = Join-Path $PSScriptRoot "apppackager-logo.jpg"
if (-not (Test-Path -LiteralPath $logoPath)) {
    $logoPath = Join-Path $PSScriptRoot "apppackager-logo.png"
}

# Form icon: prefer .ico (multi-resolution), fall back to bitmap conversion
try {
    if (Test-Path -LiteralPath $icoPath) {
        $form.Icon = New-Object System.Drawing.Icon $icoPath
    }
    elseif (Test-Path -LiteralPath $logoPath) {
        $bmp = New-Object System.Drawing.Bitmap $logoPath
        $form.Icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    }
    else {
        $form.Icon = [System.Drawing.SystemIcons]::Application
    }
}
catch {
    $form.Icon = [System.Drawing.SystemIcons]::Application
}

# Header logo
$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Size = New-Object System.Drawing.Size(196, 196)
$picLogo.Location = New-Object System.Drawing.Point(20, 20)
$picLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$picLogo.BackColor = [System.Drawing.Color]::Transparent

if (Test-Path -LiteralPath $logoPath) {
    try { $picLogo.Image = [System.Drawing.Image]::FromFile($logoPath) } catch { }
}

$form.Controls.Add($picLogo)

# -----------------------------
# Info label (replaces former 5-line descriptor panel)
# -----------------------------
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.AutoSize  = $true
$lblInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblInfo.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
$lblInfo.BackColor = $form.BackColor
$lblInfo.Text      = "No network or MECM actions occur until a button is selected."
$lblInfo.Anchor    = "Top,Left,Right"
$form.Controls.Add($lblInfo)

# -----------------------------
# Row 1 inputs — WO/Comment, File Share Root, Site Code
# -----------------------------

# Work Order / Comment
$lblComment = New-Object System.Windows.Forms.Label
$lblComment.Text = "WO / Comment:"
$lblComment.AutoSize = $true
$lblComment.BackColor = $form.BackColor
$lblComment.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblComment.Anchor = "Top,Left,Right"
$form.Controls.Add($lblComment)

$txtComment = New-Object System.Windows.Forms.TextBox
$txtComment.Text = "WO#00000001234567"
$txtComment.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtComment.Width = 220
$txtComment.MaxLength = 64
$txtComment.Anchor = "Top,Left,Right"
$form.Controls.Add($txtComment)

# File share root
$lblFSPath = New-Object System.Windows.Forms.Label
$lblFSPath.Text = "File Share Root:"
$lblFSPath.AutoSize = $true
$lblFSPath.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblFSPath.BackColor = $form.BackColor
$lblFSPath.Anchor = "Top,Right"
$form.Controls.Add($lblFSPath)

$txtFSPath = New-Object System.Windows.Forms.TextBox
$txtFSPath.Text = "\\fileserver\sccm$"
$txtFSPath.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtFSPath.Width = 200
$txtFSPath.MaxLength = 200
$txtFSPath.Anchor = "Top,Right"
$form.Controls.Add($txtFSPath)

# Site code
$lblSiteCode = New-Object System.Windows.Forms.Label
$lblSiteCode.Text = "Site Code:"
$lblSiteCode.AutoSize = $true
$lblSiteCode.BackColor = $form.BackColor
$lblSiteCode.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSiteCode.Anchor = "Top,Right"
$form.Controls.Add($lblSiteCode)

$txtSiteCode = New-Object System.Windows.Forms.TextBox
$txtSiteCode.Text = $SiteCode
$txtSiteCode.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtSiteCode.Width = 80
$txtSiteCode.MaxLength = 5
$txtSiteCode.Anchor = "Top,Right"
$form.Controls.Add($txtSiteCode)

# -----------------------------
# Row 2 inputs — Download Root, Est Runtime, Max Runtime
# -----------------------------

# Download Root
$lblDownloadRoot = New-Object System.Windows.Forms.Label
$lblDownloadRoot.Text = "Download Root:"
$lblDownloadRoot.AutoSize = $true
$lblDownloadRoot.BackColor = $form.BackColor
$lblDownloadRoot.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblDownloadRoot.Anchor = "Top,Left,Right"
$form.Controls.Add($lblDownloadRoot)

$txtDownloadRoot = New-Object System.Windows.Forms.TextBox
$txtDownloadRoot.Text = "C:\temp\ap"
$txtDownloadRoot.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtDownloadRoot.Width = 200
$txtDownloadRoot.MaxLength = 200
$txtDownloadRoot.Anchor = "Top,Left,Right"
$form.Controls.Add($txtDownloadRoot)

# Estimated Runtime (mins)
$lblEstRuntime = New-Object System.Windows.Forms.Label
$lblEstRuntime.Text = "Est. Runtime:"
$lblEstRuntime.AutoSize = $true
$lblEstRuntime.BackColor = $form.BackColor
$lblEstRuntime.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblEstRuntime.Anchor = "Top,Right"
$form.Controls.Add($lblEstRuntime)

$txtEstRuntime = New-Object System.Windows.Forms.TextBox
$txtEstRuntime.Text = "15"
$txtEstRuntime.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtEstRuntime.Width = 50
$txtEstRuntime.MaxLength = 4
$txtEstRuntime.Anchor = "Top,Right"
$form.Controls.Add($txtEstRuntime)

# Maximum Runtime (mins)
$lblMaxRuntime = New-Object System.Windows.Forms.Label
$lblMaxRuntime.Text = "Max Runtime:"
$lblMaxRuntime.AutoSize = $true
$lblMaxRuntime.BackColor = $form.BackColor
$lblMaxRuntime.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblMaxRuntime.Anchor = "Top,Right"
$form.Controls.Add($lblMaxRuntime)

$txtMaxRuntime = New-Object System.Windows.Forms.TextBox
$txtMaxRuntime.Text = "30"
$txtMaxRuntime.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtMaxRuntime.Width = 50
$txtMaxRuntime.MaxLength = 4
$txtMaxRuntime.Anchor = "Top,Right"
$form.Controls.Add($txtMaxRuntime)

# "mins" suffix labels
$lblEstMins = New-Object System.Windows.Forms.Label
$lblEstMins.Text = "mins"
$lblEstMins.AutoSize = $true
$lblEstMins.BackColor = $form.BackColor
$lblEstMins.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblEstMins.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
$lblEstMins.Anchor = "Top,Right"
$form.Controls.Add($lblEstMins)

$lblMaxMins = New-Object System.Windows.Forms.Label
$lblMaxMins.Text = "mins"
$lblMaxMins.AutoSize = $true
$lblMaxMins.BackColor = $form.BackColor
$lblMaxMins.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblMaxMins.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
$lblMaxMins.Anchor = "Top,Right"
$form.Controls.Add($lblMaxMins)

# -----------------------------
# Configuration row (between logo area and grid)
# -----------------------------
$lblLoadConfig = New-Object System.Windows.Forms.Label
$lblLoadConfig.Text = "Load Configuration:"
$lblLoadConfig.AutoSize = $true
$lblLoadConfig.BackColor = $form.BackColor
$lblLoadConfig.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblLoadConfig.Anchor = "Top,Right"
$form.Controls.Add($lblLoadConfig)

$cmbLoadConfig = New-Object System.Windows.Forms.ComboBox
$cmbLoadConfig.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbLoadConfig.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$cmbLoadConfig.Width = 220
$cmbLoadConfig.Anchor = "Top,Right"
$form.Controls.Add($cmbLoadConfig)

$lblSaveConfig = New-Object System.Windows.Forms.Label
$lblSaveConfig.Text = "Save Configuration:"
$lblSaveConfig.AutoSize = $true
$lblSaveConfig.BackColor = $form.BackColor
$lblSaveConfig.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSaveConfig.Anchor = "Top,Right"
$form.Controls.Add($lblSaveConfig)

$txtSaveConfigName = New-Object System.Windows.Forms.TextBox
$txtSaveConfigName.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtSaveConfigName.Width = 200
$txtSaveConfigName.MaxLength = 64
$txtSaveConfigName.Anchor = "Top,Right"
$txtSaveConfigName.Text = "Configuration Name"
$form.Controls.Add($txtSaveConfigName)

$btnSaveConfig = New-Object System.Windows.Forms.Button
$btnSaveConfig.Width = 44
$btnSaveConfig.Height = 26
$btnSaveConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSaveConfig.FlatAppearance.BorderSize = 0
$btnSaveConfig.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34) # green
$btnSaveConfig.ForeColor = [System.Drawing.Color]::White
$btnSaveConfig.UseVisualStyleBackColor = $false
$btnSaveConfig.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSaveConfig.Text = [char]::ConvertFromUtf32(0x2713)  # checkmark
$btnSaveConfig.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$btnSaveConfig.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$btnSaveConfig.Anchor = "Top,Right"
$form.Controls.Add($btnSaveConfig)

# Keep a script-scoped cache for the combo binding
$script:ConfigCache = @()

function Update-ConfigDropdown {
    param([string]$SelectName = $null)

    $script:ConfigCache = @(Read-ConfigStore | Sort-Object Name)

    $cmbLoadConfig.BeginUpdate()
    $cmbLoadConfig.Items.Clear()
    foreach ($c in $script:ConfigCache) {
        [void]$cmbLoadConfig.Items.Add($c.Name)
    }
    $cmbLoadConfig.EndUpdate()

    if ($SelectName) {
        $idx = $cmbLoadConfig.Items.IndexOf($SelectName)
        if ($idx -ge 0) { $cmbLoadConfig.SelectedIndex = $idx }
    }
    elseif ($cmbLoadConfig.Items.Count -gt 0 -and $cmbLoadConfig.SelectedIndex -lt 0) {
        $cmbLoadConfig.SelectedIndex = 0
    }
}

$cmbLoadConfig.Add_SelectedIndexChanged({
    $name = [string]$cmbLoadConfig.SelectedItem
    if ([string]::IsNullOrWhiteSpace($name)) { return }

    $cfg = $script:ConfigCache | Where-Object { $_.Name -eq $name } | Select-Object -First 1
    if ($cfg) {
        Set-ConfigurationInputs -Config $cfg `
            -TxtComment $txtComment -TxtFSPath $txtFSPath -TxtSiteCode $txtSiteCode `
            -TxtDownloadRoot $txtDownloadRoot -TxtEstRuntime $txtEstRuntime -TxtMaxRuntime $txtMaxRuntime
    }
})

$btnSaveConfig.Add_Click({
    $name = $txtSaveConfigName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name) -or $name -eq "Configuration Name") { return }

    $estVal = 15; $maxVal = 30
    if ([int]::TryParse($txtEstRuntime.Text.Trim(), [ref]$estVal)) { } else { $estVal = 15 }
    if ([int]::TryParse($txtMaxRuntime.Text.Trim(), [ref]$maxVal)) { } else { $maxVal = 30 }

    $null = Save-Configuration -Name $name `
        -WOComment $txtComment.Text -FileShareRoot $txtFSPath.Text -SiteCode $txtSiteCode.Text `
        -DownloadRoot $txtDownloadRoot.Text -EstimatedRuntimeMins $estVal -MaximumRuntimeMins $maxVal
    Update-ConfigDropdown -SelectName $name
})

# Initial population
Update-ConfigDropdown

# Data grid
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 70)
$grid.Size = New-Object System.Drawing.Size(1060, 430)
$grid.Anchor = "Top,Left,Right,Bottom"
$grid.ReadOnly = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AutoGenerateColumns = $false
$grid.ScrollBars = "Both"
$grid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
$grid.AutoSizeRowsMode = "None"
$grid.ColumnHeadersHeightSizeMode = "DisableResizing"
$grid.ColumnHeadersHeight = 34
$grid.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$grid.GridColor = [System.Drawing.Color]::FromArgb(220,220,220)
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
$grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $grid.ColumnHeadersDefaultCellStyle.BackColor
$grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = $grid.ColumnHeadersDefaultCellStyle.ForeColor
$grid.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
$grid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
$grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(225, 235, 245)
$grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$grid.RowTemplate.Height = 32
$grid.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None

$colSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSel.HeaderText = ""
$colSel.Width = 44
$colSel.Frozen = $true
$colSel.Name = "Selected"
$grid.Columns.Add($colSel) | Out-Null

$colVendor = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colVendor.HeaderText = "Vendor"
$colVendor.Width = 170
$colVendor.ReadOnly = $true
$colVendor.Name = "Vendor"
$grid.Columns.Add($colVendor) | Out-Null

$colApp = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colApp.HeaderText = "Application"
$colApp.Width = 360
$colApp.ReadOnly = $true
$colApp.Name = "Application"
$grid.Columns.Add($colApp) | Out-Null

$colCurrent = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCurrent.HeaderText = "Current (MECM)"
$colCurrent.Width = 160
$colCurrent.ReadOnly = $true
$colCurrent.Name = "CurrentVersion"
$grid.Columns.Add($colCurrent) | Out-Null

$colLatest = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colLatest.HeaderText = "Latest"
$colLatest.Width = 140
$colLatest.ReadOnly = $true
$colLatest.Name = "LatestVersion"
$grid.Columns.Add($colLatest) | Out-Null

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.HeaderText = "Status"
$colStatus.Width = 180
$colStatus.ReadOnly = $true
$colStatus.Name = "Status"
$grid.Columns.Add($colStatus) | Out-Null

$colCMName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCMName.HeaderText = "CMName (Debug)"
$colCMName.Width = 240
$colCMName.ReadOnly = $true
$colCMName.Name = "CMName"
$colCMName.Visible = $false
$grid.Columns.Add($colCMName) | Out-Null

$colScript = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colScript.HeaderText = "Script (Debug)"
$colScript.Width = 220
$colScript.ReadOnly = $true
$colScript.Name = "Script"
$colScript.Visible = $false
$grid.Columns.Add($colScript) | Out-Null

$form.Controls.Add($grid)

Enable-DoubleBuffer -Control $form
Enable-DoubleBuffer -Control $grid

# -----------------------------
# Select-All header checkbox
# -----------------------------
$chkSelectAll = New-Object System.Windows.Forms.CheckBox
$chkSelectAll.Size = New-Object System.Drawing.Size(16,16)
$chkSelectAll.BackColor = [System.Drawing.Color]::Transparent
$chkSelectAll.Checked = $false
$form.Controls.Add($chkSelectAll)
$chkSelectAll.BringToFront()

function Set-SelectAllCheckboxPosition {
    try {
        $headerRect = $grid.GetCellDisplayRectangle(0, -1, $true)

        $x = $grid.Left + $headerRect.X + [int](($headerRect.Width - $chkSelectAll.Width) / 2)
        $y = $grid.Top  + $headerRect.Y + [int](($headerRect.Height - $chkSelectAll.Height) / 2)

        $chkSelectAll.Location = New-Object System.Drawing.Point($x, $y)
        $chkSelectAll.Visible = $true
    }
    catch {
        $chkSelectAll.Visible = $false
    }
}

$grid.Add_Scroll({ Set-SelectAllCheckboxPosition })
$grid.Add_ColumnWidthChanged({ Set-SelectAllCheckboxPosition })
$grid.Add_SizeChanged({ Set-SelectAllCheckboxPosition })
$form.Add_Shown({ Set-SelectAllCheckboxPosition })

$chkSelectAll.Add_CheckedChanged({

    if ($script:SuppressSelectAll) {
        return
    }

    $grid.EndEdit()

    $target = $chkSelectAll.Checked
    foreach ($row in $dt.Rows) {
        $row["Selected"] = $target
    }
})

$grid.Add_CellValueChanged({
    param($s, $e)

    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
        $s.EndEdit()

        $allSelected = $true
        $anySelected = $false

        foreach ($row in $dt.Rows) {
            if ($row["Selected"] -eq $true) {
                $anySelected = $true
            }
            else {
                $allSelected = $false
            }
        }

        $script:SuppressSelectAll = $true
        try {
            $chkSelectAll.Checked = ($anySelected -and $allSelected)
        }
        finally {
            $script:SuppressSelectAll = $false
        }
    }
})

$grid.Add_CurrentCellDirtyStateChanged({
    if ($grid.IsCurrentCellDirty) {
        $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

# -----------------------------
# Right-click context menu
# -----------------------------
$gridContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$script:ContextMenuRowIndex = -1

$menuOpenLogFolder     = New-Object System.Windows.Forms.ToolStripMenuItem "Open Log Folder"
$menuOpenStagedFolder  = New-Object System.Windows.Forms.ToolStripMenuItem "Open Staged Folder"
$menuOpenNetworkShare  = New-Object System.Windows.Forms.ToolStripMenuItem "Open Network Share"
$menuSep1              = New-Object System.Windows.Forms.ToolStripSeparator
$menuCopyLatestVersion = New-Object System.Windows.Forms.ToolStripMenuItem "Copy Latest Version"

$gridContextMenu.Items.AddRange(@($menuOpenLogFolder, $menuOpenStagedFolder, $menuOpenNetworkShare, $menuSep1, $menuCopyLatestVersion))

$grid.Add_CellMouseClick({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0) {
        $script:ContextMenuRowIndex = $e.RowIndex
        $grid.ClearSelection()
        $grid.Rows[$e.RowIndex].Selected = $true

        $row = $dt.Rows[$e.RowIndex]
        $menuCopyLatestVersion.Enabled = (-not [string]::IsNullOrWhiteSpace([string]$row["LatestVersion"]))

        $gridContextMenu.Show($grid, $grid.PointToClient([System.Windows.Forms.Cursor]::Position))
    }
})

$menuOpenLogFolder.Add_Click({
    $logFolder = Join-Path $PSScriptRoot "Logs"
    if (-not (Test-Path -LiteralPath $logFolder)) {
        New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
    }
    Start-Process "explorer.exe" -ArgumentList $logFolder
})

$menuOpenStagedFolder.Add_Click({
    if ($script:ContextMenuRowIndex -lt 0) { return }
    $row = $dt.Rows[$script:ContextMenuRowIndex]
    $dlRoot = $txtDownloadRoot.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($dlRoot)) {
        Add-LogLine -TextBox $txtLog -Message "Download Root is not set."
        return
    }

    $info = Get-PackagerFolderInfo -ScriptPath ([string]$row["FullPath"])
    if ($info.DownloadSubfolder) {
        $targetPath = Join-Path $dlRoot $info.DownloadSubfolder

        # Try version-specific subfolder first
        $version = [string]$row["LatestVersion"]
        if (-not [string]::IsNullOrWhiteSpace($version)) {
            $versionPath = Join-Path $targetPath $version
            if (Test-Path -LiteralPath $versionPath) {
                Start-Process "explorer.exe" -ArgumentList $versionPath
                return
            }
        }

        if (Test-Path -LiteralPath $targetPath) {
            Start-Process "explorer.exe" -ArgumentList $targetPath
            return
        }
    }

    if (Test-Path -LiteralPath $dlRoot) {
        Start-Process "explorer.exe" -ArgumentList $dlRoot
    }
    else {
        Add-LogLine -TextBox $txtLog -Message ("Folder not found: {0}" -f $dlRoot)
    }
})

$menuOpenNetworkShare.Add_Click({
    if ($script:ContextMenuRowIndex -lt 0) { return }
    $row = $dt.Rows[$script:ContextMenuRowIndex]
    $fsPath = $txtFSPath.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($fsPath)) {
        Add-LogLine -TextBox $txtLog -Message "File Share Root is not set."
        return
    }

    $info = Get-PackagerFolderInfo -ScriptPath ([string]$row["FullPath"])
    if ($info.VendorFolder -and $info.AppFolder) {
        $targetPath = Join-Path (Join-Path (Join-Path $fsPath "Applications") $info.VendorFolder) $info.AppFolder
        if (Test-Path -LiteralPath $targetPath) {
            Start-Process "explorer.exe" -ArgumentList $targetPath
            return
        }
    }

    $appsRoot = Join-Path $fsPath "Applications"
    if (Test-Path -LiteralPath $appsRoot) {
        Start-Process "explorer.exe" -ArgumentList $appsRoot
    }
    else {
        Add-LogLine -TextBox $txtLog -Message ("Network path not accessible: {0}" -f $appsRoot)
    }
})

$menuCopyLatestVersion.Add_Click({
    if ($script:ContextMenuRowIndex -lt 0) { return }
    $row = $dt.Rows[$script:ContextMenuRowIndex]
    $version = [string]$row["LatestVersion"]
    if (-not [string]::IsNullOrWhiteSpace($version)) {
        [System.Windows.Forms.Clipboard]::SetText($version)
        Add-LogLine -TextBox $txtLog -Message ("Copied version to clipboard: {0}" -f $version)
    }
})

# Debug toggle
$chkDebug = New-Object System.Windows.Forms.CheckBox
$chkDebug.Text = "Show Debug Columns"
$chkDebug.AutoSize = $true
$chkDebug.Location = New-Object System.Drawing.Point(20, 510)
$chkDebug.Anchor = "Top,Left"
$form.Controls.Add($chkDebug)

$btnSelectUpdates = New-Object System.Windows.Forms.Button
$btnSelectUpdates.Text = "Select Update Available"
$btnSelectUpdates.AutoSize = $true
$btnSelectUpdates.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelectUpdates.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
$btnSelectUpdates.FlatAppearance.BorderSize = 1
$btnSelectUpdates.BackColor = [System.Drawing.Color]::White
$btnSelectUpdates.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$btnSelectUpdates.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnSelectUpdates.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSelectUpdates.Height = 26
$btnSelectUpdates.Anchor = "Bottom,Left"
$form.Controls.Add($btnSelectUpdates)

$btnSelectUpdates.Add_Click({
    $grid.EndEdit()
    Select-OnlyUpdateAvailable -DataTable $dt
    $grid.Refresh()
    Add-LogLine -TextBox $txtLog -Message "Selected rows with 'Update available' status."
})

# Unicode emoji code points
$emojiSearch  = [char]::ConvertFromUtf32(0x1F50D)  # search
$emojiBox     = [char]::ConvertFromUtf32(0x1F5C4)  # file cabinet
$emojiDown    = [char]::ConvertFromUtf32(0x2B07)    # down arrow
$emojiPackage = [char]::ConvertFromUtf32(0x1F4E6)   # package

# Buttons row
$btnLatest = New-Object System.Windows.Forms.Button
$btnLatest.Text = "$emojiSearch  Check Latest"
$btnLatest.Size = New-Object System.Drawing.Size(240, 52)
$btnLatest.Location = New-Object System.Drawing.Point(20, 548)
$btnLatest.Anchor = "Bottom,Left"
$btnLatest.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$btnLatest.ImageAlign = "MiddleLeft"
$btnLatest.TextAlign = "MiddleCenter"
$btnLatest.Padding = 0
$form.Controls.Add($btnLatest)

$btnMecm = New-Object System.Windows.Forms.Button
$btnMecm.Text   = "$emojiBox  Check MECM"
$btnMecm.Size = New-Object System.Drawing.Size(240, 52)
$btnMecm.Location = New-Object System.Drawing.Point(280, 548)
$btnMecm.Anchor = "Bottom,Left"
$btnMecm.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$btnMecm.ImageAlign = "MiddleLeft"
$btnMecm.TextAlign = "MiddleCenter"
$btnMecm.Padding = 0
$form.Controls.Add($btnMecm)

$btnStage = New-Object System.Windows.Forms.Button
$btnStage.Text    = "$emojiDown  Stage Packages"
$btnStage.Size = New-Object System.Drawing.Size(240, 52)
$btnStage.Location = New-Object System.Drawing.Point(540, 548)
$btnStage.Anchor = "Bottom,Left"
$btnStage.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$btnStage.ImageAlign = "MiddleLeft"
$btnStage.TextAlign = "MiddleCenter"
$btnStage.Padding = 0
$form.Controls.Add($btnStage)

$btnPackage = New-Object System.Windows.Forms.Button
$btnPackage.Text    = "$emojiPackage  Package Apps"
$btnPackage.Size = New-Object System.Drawing.Size(240, 52)
$btnPackage.Location = New-Object System.Drawing.Point(800, 548)
$btnPackage.Anchor = "Bottom,Right"
$btnPackage.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$btnPackage.ImageAlign = "MiddleLeft"
$btnPackage.TextAlign = "MiddleCenter"
$btnPackage.Padding = 0
$form.Controls.Add($btnPackage)

Set-ModernButtonStyle -Button $btnLatest  -BackColor ([System.Drawing.Color]::FromArgb(0, 120, 212))
Set-ModernButtonStyle -Button $btnMecm    -BackColor ([System.Drawing.Color]::FromArgb(16, 124, 16))
Set-ModernButtonStyle -Button $btnStage   -BackColor ([System.Drawing.Color]::FromArgb(217, 95, 2))
Set-ModernButtonStyle -Button $btnPackage -BackColor ([System.Drawing.Color]::FromArgb(100, 60, 160))

# 3-line log output
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.ReadOnly = $true
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.WordWrap = $false
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 10)
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtLog.Anchor = "Bottom,Left,Right"
$txtLog.Location = New-Object System.Drawing.Point(20, 612)
$txtLog.Size = New-Object System.Drawing.Size(1060, 54)
$form.Controls.Add($txtLog)

# Status strip
$status = New-Object System.Windows.Forms.StatusStrip
$status.Dock = [System.Windows.Forms.DockStyle]::Bottom
$status.SizingGrip = $false
$status.BackColor = [System.Drawing.Color]::White
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready. No network actions are performed until you click a button."
$status.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($status)

# -----------------------------
# Layout
# -----------------------------
function Set-UILayout {

    $padding   = 20
    $gap       = 20
    $btnHeight = 52
    $logHeight = 54
    $headerGap = 12
    $rowHeight = 30                # vertical spacing between config rows

    # Right edge
    $right = ($form.ClientSize.Width - $padding)

    # Minimum left boundary (right of logo + gap)
    $minLeft = ($picLogo.Right + 20)

    # -------------------------------------------------------
    # Config row 0 (Load / Save configuration) — right-aligned
    # -------------------------------------------------------
    $configY = ($picLogo.Top + 24)

    $saveBtnW  = 44
    $saveNameW = 200
    $loadW     = 220
    $cfgGap    = 12

    $btnSaveConfig.SetBounds(($right - $saveBtnW), $configY, $saveBtnW, 26)
    $txtSaveConfigName.SetBounds(($btnSaveConfig.Left - $cfgGap - $saveNameW), $configY, $saveNameW, 24)
    $lblSaveConfig.Location = New-Object System.Drawing.Point(
        ($txtSaveConfigName.Left - 8 - $lblSaveConfig.PreferredWidth),
        ($configY + 3)
    )
    $cmbLoadConfig.SetBounds(($lblSaveConfig.Left - $gap - $loadW), $configY, $loadW, 24)
    $lblLoadConfig.Location = New-Object System.Drawing.Point(
        ($cmbLoadConfig.Left - 8 - $lblLoadConfig.PreferredWidth),
        ($configY + 3)
    )

    # -------------------------------------------------------
    # Row 1 — WO / Comment, File Share Root, Site Code
    # -------------------------------------------------------
    $row1Y = ($configY + $rowHeight + 6)

    $siteBoxWidth    = 80
    $fsBoxWidth      = 200
    $commentBoxWidth = 220

    # Site Code at far right
    $txtSiteCode.SetBounds(($right - $siteBoxWidth), $row1Y, $siteBoxWidth, 24)
    $lblSiteCode.Location = New-Object System.Drawing.Point(
        ($txtSiteCode.Left - 8 - $lblSiteCode.PreferredWidth),
        ($row1Y + 3)
    )

    # File Share Root to the left of Site Code
    $fsRight = ($lblSiteCode.Left - $headerGap)
    $minFsBoxWidth = 160
    for ($i = 0; $i -lt 5; $i++) {
        $fsLeft = ($fsRight - $fsBoxWidth)
        $fsLabelLeft = ($fsLeft - 8 - $lblFSPath.PreferredWidth)
        if ($fsLabelLeft -ge $minLeft) { break }
        $delta = ($minLeft - $fsLabelLeft)
        $fsBoxWidth = ($fsBoxWidth - $delta)
        if ($fsBoxWidth -lt $minFsBoxWidth) { $fsBoxWidth = $minFsBoxWidth; break }
    }
    $fsLeft = ($fsRight - $fsBoxWidth)
    $txtFSPath.SetBounds($fsLeft, $row1Y, $fsBoxWidth, 24)
    $lblFSPath.Location = New-Object System.Drawing.Point(
        ($txtFSPath.Left - 8 - $lblFSPath.PreferredWidth),
        ($row1Y + 3)
    )

    # WO / Comment to the left of File Share Root
    $commentRight = ($lblFSPath.Left - $headerGap)
    $commentLeft  = ($commentRight - $commentBoxWidth)
    $minCommentBoxWidth = 140
    if ($commentLeft -lt $minLeft) {
        $delta = ($minLeft - $commentLeft)
        $commentBoxWidth = ($commentBoxWidth - $delta)
        if ($commentBoxWidth -lt $minCommentBoxWidth) { $commentBoxWidth = $minCommentBoxWidth }
        $commentLeft = ($commentRight - $commentBoxWidth)
    }
    $txtComment.SetBounds($commentLeft, $row1Y, $commentBoxWidth, 24)
    $lblComment.Location = New-Object System.Drawing.Point(
        ($txtComment.Left - 8 - $lblComment.PreferredWidth),
        ($row1Y + 3)
    )

    # -------------------------------------------------------
    # Row 2 — Download Root, Est Runtime, Max Runtime
    # -------------------------------------------------------
    $row2Y = ($row1Y + $rowHeight + 4)

    $runtimeBoxW = 50
    $minsLabelW  = 30

    # Max Runtime at far right: [label] [box] mins
    $lblMaxMins.Location = New-Object System.Drawing.Point(($right - $minsLabelW), ($row2Y + 4))
    $txtMaxRuntime.SetBounds(($lblMaxMins.Left - 4 - $runtimeBoxW), $row2Y, $runtimeBoxW, 24)
    $lblMaxRuntime.Location = New-Object System.Drawing.Point(
        ($txtMaxRuntime.Left - 8 - $lblMaxRuntime.PreferredWidth),
        ($row2Y + 3)
    )

    # Est Runtime to the left of Max Runtime
    $lblEstMins.Location = New-Object System.Drawing.Point(($lblMaxRuntime.Left - $headerGap - $minsLabelW), ($row2Y + 4))
    $txtEstRuntime.SetBounds(($lblEstMins.Left - 4 - $runtimeBoxW), $row2Y, $runtimeBoxW, 24)
    $lblEstRuntime.Location = New-Object System.Drawing.Point(
        ($txtEstRuntime.Left - 8 - $lblEstRuntime.PreferredWidth),
        ($row2Y + 3)
    )

    # Download Root fills remaining space to the left
    $dlRight = ($lblEstRuntime.Left - $headerGap)
    $dlBoxWidth = 200
    $minDlBoxWidth = 120
    $dlLeft = ($dlRight - $dlBoxWidth)
    $dlLabelLeft = ($dlLeft - 8 - $lblDownloadRoot.PreferredWidth)
    if ($dlLabelLeft -lt $minLeft) {
        $delta = ($minLeft - $dlLabelLeft)
        $dlBoxWidth = ($dlBoxWidth - $delta)
        if ($dlBoxWidth -lt $minDlBoxWidth) { $dlBoxWidth = $minDlBoxWidth }
        $dlLeft = ($dlRight - $dlBoxWidth)
    }
    $txtDownloadRoot.SetBounds($dlLeft, $row2Y, $dlBoxWidth, 24)
    $lblDownloadRoot.Location = New-Object System.Drawing.Point(
        ($txtDownloadRoot.Left - 8 - $lblDownloadRoot.PreferredWidth),
        ($row2Y + 3)
    )

    # -------------------------------------------------------
    # Info label — below row 2
    # -------------------------------------------------------
    $infoY = ($row2Y + $rowHeight + 2)
    $lblInfo.Location = New-Object System.Drawing.Point($minLeft, $infoY)
    $lblInfo.MaximumSize = New-Object System.Drawing.Size(($right - $minLeft), 0)

    # -------------------------------------------------------
    # Grid and bottom controls
    # -------------------------------------------------------
    $gridTop = ([Math]::Max($picLogo.Bottom, ($infoY + 22)) + 8)

    $bottomY = $form.ClientSize.Height - $status.Height - $padding
    $txtLog.SetBounds($padding, ($bottomY - $logHeight), ($form.ClientSize.Width - (2 * $padding)), $logHeight)

    $btnWidth = [int](($form.ClientSize.Width - (2 * $padding) - (3 * $gap)) / 4)
    if ($btnWidth -lt 180) { $btnWidth = 180 }

    $btnY = $txtLog.Top - 10 - $btnHeight
    $btnLatest.SetBounds($padding, $btnY, $btnWidth, $btnHeight)
    $btnMecm.SetBounds(($padding + $btnWidth + $gap), $btnY, $btnWidth, $btnHeight)
    $btnStage.SetBounds(($padding + (2 * ($btnWidth + $gap))), $btnY, $btnWidth, $btnHeight)
    $btnPackage.SetBounds(($padding + (3 * ($btnWidth + $gap))), $btnY, $btnWidth, $btnHeight)

    $chkDebug.Location = New-Object System.Drawing.Point($padding, ($btnY - 34))
    $btnSelectUpdates.Location = New-Object System.Drawing.Point(($chkDebug.Right + 20), ($btnY - 36))

    $grid.SetBounds($padding, $gridTop, ($form.ClientSize.Width - (2 * $padding)), ($chkDebug.Top - 10 - $gridTop))

    Set-SelectAllCheckboxPosition
}

$form.Add_Shown({ Set-UILayout })
$form.Add_Resize({ Set-UILayout })

# -----------------------------
# Data model
# -----------------------------
$dt = New-Object System.Data.DataTable
[void]$dt.Columns.Add("Selected", [bool])
[void]$dt.Columns.Add("Vendor", [string])
[void]$dt.Columns.Add("Application", [string])
[void]$dt.Columns.Add("CurrentVersion", [string])
[void]$dt.Columns.Add("LatestVersion", [string])
[void]$dt.Columns.Add("Status", [string])
[void]$dt.Columns.Add("CMName", [string])
[void]$dt.Columns.Add("Script", [string])
[void]$dt.Columns.Add("FullPath", [string])

$grid.DataSource = $dt

$grid.Columns["Selected"].DataPropertyName = "Selected"
$grid.Columns["Vendor"].DataPropertyName = "Vendor"
$grid.Columns["Application"].DataPropertyName = "Application"
$grid.Columns["CurrentVersion"].DataPropertyName = "CurrentVersion"
$grid.Columns["LatestVersion"].DataPropertyName = "LatestVersion"
$grid.Columns["Status"].DataPropertyName = "Status"
$grid.Columns["CMName"].DataPropertyName = "CMName"
$grid.Columns["Script"].DataPropertyName = "Script"

$grid.Add_CellBeginEdit({
    param($s, $e)
    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
        $path = [string]$dt.Rows[$e.RowIndex]["FullPath"]
        if ($path -match "\.notps1$") {
            $e.Cancel = $true
        }
    }
})

$grid.Add_RowPrePaint({
    param($s, $e)
    try {
        $path = [string]$dt.Rows[$e.RowIndex]["FullPath"]
        if ($path -match "\.notps1$") {
            $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGray
        }
    }
    catch { }
})

# -----------------------------
# Events
# -----------------------------
$chkDebug.Add_CheckedChanged({
    $show = $chkDebug.Checked
    $grid.Columns["CMName"].Visible = $show
    $grid.Columns["Script"].Visible = $show
})

$btnLatest.Add_Click({
    $siteCodeValue = $txtSiteCode.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -TextBox $txtLog -Message "SiteCode is required."
        $statusLabel.Text = "SiteCode is required."
        return
    }

    $selectedRows = @()
    foreach ($r in $dt.Rows) {
        if ($r["Selected"] -eq $true) { $selectedRows += $r }
    }

    if (-not $selectedRows -or $selectedRows.Count -eq 0) {
        Add-LogLine -TextBox $txtLog -Message "No rows selected."
        return
    }

    $btnLatest.Enabled  = $false
    $btnMecm.Enabled    = $false
    $btnStage.Enabled   = $false
    $btnPackage.Enabled = $false
    $form.UseWaitCursor = $true

    try {
        $statusLabel.Text = "Checking latest versions for selected packagers..."

        foreach ($row in $selectedRows) {
            [System.Windows.Forms.Application]::DoEvents()

            $app     = [string]$row["Application"]
            $script  = [string]$row["Script"]
            $path    = [string]$row["FullPath"]

            Add-LogLine -TextBox $txtLog -Message ("Latest: {0} ({1})" -f $app, $script)
            $row["Status"] = "Checking latest..."

            try {
                $latest = Invoke-PackagerGetLatestVersion `
                    -PackagerPath $path `
                    -SiteCode $siteCodeValue `
                    -FileServerPath $txtFSPath.Text.Trim() `
                    -DownloadRoot $txtDownloadRoot.Text.Trim()
                $row["LatestVersion"] = $latest

                $current = [string]$row["CurrentVersion"]
                if (-not [string]::IsNullOrWhiteSpace($current)) {
                    $cmp = Compare-SemVer -A $current -B $latest
                    if ($cmp -lt 0) {
                        $row["Status"] = "Update available"
                    }
                    elseif ($cmp -eq 0) {
                        $row["Status"] = "Up to date"
                    }
                    else {
                        $row["Status"] = "Current newer"
                    }
                }
                else {
                    $row["Status"] = "Latest retrieved"
                }

                Add-LogLine -TextBox $txtLog -Message ("Latest version: {0}" -f $latest)
            }
            catch {
                $row["Status"] = "Error"
                Add-LogLine -TextBox $txtLog -Message ("Error: {0}" -f $_.Exception.Message)
            }
        }

        $statusLabel.Text = "Latest check complete."
    }
    finally {
        $form.UseWaitCursor = $false
        $btnLatest.Enabled  = $true
        $btnMecm.Enabled    = $true
        $btnStage.Enabled   = $true
        $btnPackage.Enabled = $true
    }
})

$btnMecm.Add_Click({
    $siteCodeValue = $txtSiteCode.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -TextBox $txtLog -Message "SiteCode is required."
        $statusLabel.Text = "SiteCode is required."
        return
    }

    $selectedRows = @()
    foreach ($r in $dt.Rows) {
        if ($r["Selected"] -eq $true) { $selectedRows += $r }
    }

    if (-not $selectedRows -or $selectedRows.Count -eq 0) {
        Add-LogLine -TextBox $txtLog -Message "No rows selected."
        return
    }

    $btnLatest.Enabled  = $false
    $btnMecm.Enabled    = $false
    $btnStage.Enabled   = $false
    $btnPackage.Enabled = $false
    $form.UseWaitCursor = $true

    try {
        $statusLabel.Text = "Querying MECM for selected products..."

        foreach ($row in $selectedRows) {
            [System.Windows.Forms.Application]::DoEvents()

            $app    = [string]$row["Application"]
            $cmName = [string]$row["CMName"]

            Add-LogLine -TextBox $txtLog -Message ("MECM: {0}" -f $app)
            $row["Status"] = "Querying MECM..."

            try {
                $res = Get-MecmCurrentVersionByCMName -SiteCode $siteCodeValue -CMName $cmName

                if (-not $res.Found) {
                    $row["CurrentVersion"] = ""
                    $row["Status"] = "Not found in MECM"
                    Add-LogLine -TextBox $txtLog -Message "Not found."
                    continue
                }

                $row["CurrentVersion"] = [string]$res.SoftwareVersion

                $latest = [string]$row["LatestVersion"]
                if (-not [string]::IsNullOrWhiteSpace($latest) -and -not [string]::IsNullOrWhiteSpace($res.SoftwareVersion)) {
                    $cmp = Compare-SemVer -A ([string]$res.SoftwareVersion) -B $latest
                    if ($cmp -lt 0)      { $row["Status"] = "Update available" }
                    elseif ($cmp -eq 0)  { $row["Status"] = "Up to date" }
                    else                 { $row["Status"] = "Current newer" }
                }
                else {
                    $row["Status"] = "MECM version retrieved"
                }

                if ($res.MatchCount -gt 1) {
                    Add-LogLine -TextBox $txtLog -Message ("Found {0} matches; using: {1} ({2})" -f $res.MatchCount, $res.DisplayName, $res.SoftwareVersion)
                }
                else {
                    Add-LogLine -TextBox $txtLog -Message ("Current version: {0}" -f $res.SoftwareVersion)
                }
            }
            catch {
                $row["Status"] = "Error"
                Add-LogLine -TextBox $txtLog -Message ("Error: {0}" -f $_.Exception.Message)
            }
        }

        Select-OnlyUpdateAvailable -DataTable $dt
        $grid.EndEdit()
        $grid.Refresh()

        $statusLabel.Text = "MECM query complete."
    }
    finally {
        $form.UseWaitCursor = $false
        $btnLatest.Enabled  = $true
        $btnMecm.Enabled    = $true
        $btnStage.Enabled   = $true
        $btnPackage.Enabled = $true
    }
})

$btnStage.Add_Click({
    $dlRootValue = $txtDownloadRoot.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($dlRootValue)) {
        Add-LogLine -TextBox $txtLog -Message "Download Root is required for staging."
        $statusLabel.Text = "Download Root is required."
        return
    }

    $selectedRows = @()
    foreach ($r in $dt.Rows) {
        if ($r["Selected"] -eq $true) { $selectedRows += $r }
    }

    if (-not $selectedRows -or $selectedRows.Count -eq 0) {
        Add-LogLine -TextBox $txtLog -Message "No rows selected."
        return
    }

    $btnLatest.Enabled  = $false
    $btnMecm.Enabled    = $false
    $btnStage.Enabled   = $false
    $btnPackage.Enabled = $false
    $form.UseWaitCursor = $true

    try {
        $logFolder = Join-Path $PSScriptRoot "Logs"
        $statusLabel.Text = "Staging selected packages..."

        foreach ($row in $selectedRows) {
            [System.Windows.Forms.Application]::DoEvents()

            $app    = [string]$row["Application"]
            $script = [string]$row["Script"]
            $path   = [string]$row["FullPath"]

            $row["Status"] = "Staging..."
            Add-LogLine -TextBox $txtLog -Message ("Stage: {0} ({1})" -f $app, $script)

            try {
                $res = Invoke-PackagerStage `
                    -PackagerPath $path `
                    -LogFolder $logFolder `
                    -DownloadRoot $dlRootValue `
                    -LogTextBox $txtLog

                if ($res.ExitCode -eq 0) {
                    $row["Status"] = "Staged"
                    Add-LogLine -TextBox $txtLog -Message ("Staged. Logs: {0}" -f (Split-Path -Leaf $res.OutLog))
                }
                else {
                    $row["Status"] = "Stage error"

                    $stderrLines = @($res.StdErr -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })

                    if ($stderrLines.Count -gt 0) {
                        $linesToShow = [Math]::Min($stderrLines.Count, 10)
                        for ($i = 0; $i -lt $linesToShow; $i++) {
                            Add-LogLine -TextBox $txtLog -Message ("  stderr: {0}" -f $stderrLines[$i])
                        }
                        if ($stderrLines.Count -gt $linesToShow) {
                            Add-LogLine -TextBox $txtLog -Message ("  ... and {0} more line(s) in: {1}" -f ($stderrLines.Count - $linesToShow), (Split-Path -Leaf $res.ErrLog))
                        }
                    }
                    else {
                        Add-LogLine -TextBox $txtLog -Message ("Error: Exit code {0}, no stderr output." -f $res.ExitCode)
                    }

                    Add-LogLine -TextBox $txtLog -Message ("Logs: {0}" -f (Split-Path -Leaf $res.OutLog))
                }
            }
            catch {
                $row["Status"] = "Stage error"
                Add-LogLine -TextBox $txtLog -Message ("Error: {0}" -f $_.Exception.Message)
            }
        }

        $statusLabel.Text = "Stage complete."
    }
    finally {
        $form.UseWaitCursor = $false
        $btnLatest.Enabled  = $true
        $btnMecm.Enabled    = $true
        $btnStage.Enabled   = $true
        $btnPackage.Enabled = $true
    }
})

$btnPackage.Add_Click({
    $siteCodeValue = $txtSiteCode.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -TextBox $txtLog -Message "SiteCode is required."
        $statusLabel.Text = "SiteCode is required."
        return
    }

    $commentValue = $txtComment.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($commentValue)) {
        Add-LogLine -TextBox $txtLog -Message "Work Order / Comment is required."
        $statusLabel.Text = "Work Order / Comment is required."
        return
    }
    $fsPathValue = $txtFSPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($fsPathValue)) {
        Add-LogLine -TextBox $txtLog -Message "File Share Root is required."
        $statusLabel.Text = "File Share Root is required."
        return
    }

    $dlRootValue = $txtDownloadRoot.Text.Trim()
    $estVal = 15; $maxVal = 30
    if (-not [int]::TryParse($txtEstRuntime.Text.Trim(), [ref]$estVal)) { $estVal = 15 }
    if (-not [int]::TryParse($txtMaxRuntime.Text.Trim(), [ref]$maxVal)) { $maxVal = 30 }

    $selectedRows = @()
    foreach ($r in $dt.Rows) {
        if ($r["Selected"] -eq $true) { $selectedRows += $r }
    }

    if (-not $selectedRows -or $selectedRows.Count -eq 0) {
        Add-LogLine -TextBox $txtLog -Message "No rows selected."
        return
    }

    $btnLatest.Enabled  = $false
    $btnMecm.Enabled    = $false
    $btnStage.Enabled   = $false
    $btnPackage.Enabled = $false
    $form.UseWaitCursor = $true

    try {
        $logFolder = Join-Path $PSScriptRoot "Logs"
        $statusLabel.Text = "Packaging selected applications..."

        foreach ($row in $selectedRows) {
            [System.Windows.Forms.Application]::DoEvents()

            $app    = [string]$row["Application"]
            $script = [string]$row["Script"]
            $path   = [string]$row["FullPath"]

            $row["Status"] = "Packaging..."
            Add-LogLine -TextBox $txtLog -Message ("Package: {0} ({1})" -f $app, $script)

            try {
                $res = Invoke-PackagerPackage `
                    -PackagerPath $path `
                    -SiteCode $siteCodeValue `
                    -Comment $commentValue `
                    -FileServerPath $fsPathValue `
                    -LogFolder $logFolder `
                    -DownloadRoot $dlRootValue `
                    -EstimatedRuntimeMins $estVal `
                    -MaximumRuntimeMins $maxVal `
                    -LogTextBox $txtLog

                if ($res.ExitCode -eq 0) {
                    $row["Status"] = "Packaged"
                    Add-LogLine -TextBox $txtLog -Message ("Packaged. Logs: {0}" -f (Split-Path -Leaf $res.OutLog))
                }
                else {
                    $row["Status"] = "Package error"

                    $stderrLines = @($res.StdErr -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })

                    if ($stderrLines.Count -gt 0) {
                        $linesToShow = [Math]::Min($stderrLines.Count, 10)
                        for ($i = 0; $i -lt $linesToShow; $i++) {
                            Add-LogLine -TextBox $txtLog -Message ("  stderr: {0}" -f $stderrLines[$i])
                        }
                        if ($stderrLines.Count -gt $linesToShow) {
                            Add-LogLine -TextBox $txtLog -Message ("  ... and {0} more line(s) in: {1}" -f ($stderrLines.Count - $linesToShow), (Split-Path -Leaf $res.ErrLog))
                        }
                    }
                    else {
                        Add-LogLine -TextBox $txtLog -Message ("Error: Exit code {0}, no stderr output." -f $res.ExitCode)
                    }

                    Add-LogLine -TextBox $txtLog -Message ("Logs: {0}" -f (Split-Path -Leaf $res.OutLog))
                }
            }
            catch {
                $row["Status"] = "Package error"
                Add-LogLine -TextBox $txtLog -Message ("Error: {0}" -f $_.Exception.Message)
            }
        }

        $statusLabel.Text = "Package complete."
    }
    finally {
        $form.UseWaitCursor = $false
        $btnLatest.Enabled  = $true
        $btnMecm.Enabled    = $true
        $btnStage.Enabled   = $true
        $btnPackage.Enabled = $true
    }
})

$form.Add_Shown({
    Add-LogLine -TextBox $txtLog -Message ("Loading packagers from: {0}" -f $PackagersRoot)

    $items = Get-Packagers -Root $PackagersRoot
    foreach ($m in $items) {
        $row = $dt.NewRow()
        $row["Selected"] = $false
        $row["Vendor"] = $m.Vendor
        $row["Application"] = $m.Application
        $row["CurrentVersion"] = ""
        $row["LatestVersion"] = ""
        $row["Status"] = $m.Status
        $row["CMName"] = $m.CMName
        $row["Script"] = $m.Script
        $row["FullPath"] = $m.FullPath
        $dt.Rows.Add($row) | Out-Null
    }

    $statusLabel.Text = ("Loaded {0} packager(s). Ready." -f $dt.Rows.Count)
})

Restore-WindowState -Form $form

$form.Add_FormClosing({
    Save-WindowState -Form $form
})

[void]$form.ShowDialog()
