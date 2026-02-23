#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for AppPackagerCommon shared module.

.DESCRIPTION
    Tests pure-logic and local-filesystem functions. Does NOT require MECM,
    network shares, real MSI files, or administrator elevation.

.EXAMPLE
    Invoke-Pester .\AppPackagerCommon.Tests.ps1
#>

BeforeAll {
    Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
}

# ============================================================================
# Write-Log / Initialize-Logging
# ============================================================================

Describe 'Write-Log' {
    It 'writes formatted message to log file' {
        $logFile = Join-Path $TestDrive 'test.log'
        Initialize-Logging -LogPath $logFile

        Write-Log 'Hello world' -Quiet

        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] Hello world'
    }

    It 'tags WARN messages correctly' {
        $logFile = Join-Path $TestDrive 'warn.log'
        Initialize-Logging -LogPath $logFile

        Write-Log 'Something odd' -Level WARN -Quiet

        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[WARN \] Something odd'
    }

    It 'tags ERROR messages correctly' {
        $logFile = Join-Path $TestDrive 'error.log'
        Initialize-Logging -LogPath $logFile

        Write-Log 'Failure' -Level ERROR -Quiet

        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[ERROR\] Failure'
    }

    It 'accepts empty string message' {
        $logFile = Join-Path $TestDrive 'empty.log'
        Initialize-Logging -LogPath $logFile

        { Write-Log '' -Quiet } | Should -Not -Throw

        $lines = Get-Content -LiteralPath $logFile
        # Header line + empty-message line
        $lines.Count | Should -BeGreaterOrEqual 2
    }
}

Describe 'Initialize-Logging' {
    It 'creates log file with header line' {
        $logFile = Join-Path $TestDrive 'init.log'
        Initialize-Logging -LogPath $logFile

        Test-Path -LiteralPath $logFile | Should -BeTrue
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] === Log initialized ==='
    }

    It 'creates parent directories if missing' {
        $logFile = Join-Path $TestDrive 'sub\dir\deep.log'
        Initialize-Logging -LogPath $logFile

        Test-Path -LiteralPath $logFile | Should -BeTrue
    }
}

# ============================================================================
# New-MsiWrapperContent
# ============================================================================

Describe 'New-MsiWrapperContent' {
    BeforeAll {
        $result = New-MsiWrapperContent -MsiFileName 'acme-widget.msi'
    }

    It 'returns a hashtable with Install and Uninstall keys' {
        $result | Should -BeOfType [hashtable]
        $result.Keys | Should -Contain 'Install'
        $result.Keys | Should -Contain 'Uninstall'
    }

    It 'install script references the MSI filename' {
        $result.Install | Should -Match 'acme-widget\.msi'
    }

    It 'install script uses msiexec /i with /qn /norestart' {
        $result.Install | Should -Match 'msiexec\.exe'
        $result.Install | Should -Match '/i'
        $result.Install | Should -Match '/qn'
        $result.Install | Should -Match '/norestart'
    }

    It 'uninstall script uses msiexec /x with /qn /norestart' {
        $result.Uninstall | Should -Match 'msiexec\.exe'
        $result.Uninstall | Should -Match '/x'
        $result.Uninstall | Should -Match '/qn'
        $result.Uninstall | Should -Match '/norestart'
    }

    It 'both scripts end with exit $proc.ExitCode' {
        $result.Install   | Should -Match 'exit \$proc\.ExitCode'
        $result.Uninstall | Should -Match 'exit \$proc\.ExitCode'
    }

    It 'uses array-form ArgumentList (not single-string)' {
        # Should contain @( indicating array splatting
        $result.Install | Should -Match '-ArgumentList @\('
    }
}

# ============================================================================
# New-ExeWrapperContent
# ============================================================================

Describe 'New-ExeWrapperContent' {
    Context 'with uninstall args' {
        BeforeAll {
            $result = New-ExeWrapperContent `
                -InstallerFileName 'setup.exe' `
                -InstallArgs "'/S', '/norestart'" `
                -UninstallCommand 'C:\Program Files\Acme\uninstall.exe' `
                -UninstallArgs "'/S'"
        }

        It 'returns a hashtable with Install and Uninstall keys' {
            $result | Should -BeOfType [hashtable]
            $result.Keys | Should -Contain 'Install'
            $result.Keys | Should -Contain 'Uninstall'
        }

        It 'install script references the installer filename' {
            $result.Install | Should -Match 'setup\.exe'
        }

        It 'install script includes install args' {
            $result.Install | Should -Match '/S'
        }

        It 'uninstall script references the uninstall command' {
            $result.Uninstall | Should -Match 'uninstall\.exe'
        }

        It 'uninstall script includes uninstall args' {
            $result.Uninstall | Should -Match '/S'
        }

        It 'both scripts end with exit $proc.ExitCode' {
            $result.Install   | Should -Match 'exit \$proc\.ExitCode'
            $result.Uninstall | Should -Match 'exit \$proc\.ExitCode'
        }
    }

    Context 'without uninstall args' {
        BeforeAll {
            $result = New-ExeWrapperContent `
                -InstallerFileName 'setup.exe' `
                -InstallArgs "'/S'" `
                -UninstallCommand 'C:\Program Files\Acme\uninstall.exe'
        }

        It 'uninstall script omits -ArgumentList when args empty' {
            $result.Uninstall | Should -Not -Match '-ArgumentList'
        }
    }
}

# ============================================================================
# Write-ContentWrappers
# ============================================================================

Describe 'Write-ContentWrappers' {
    BeforeAll {
        $outDir = Join-Path $TestDrive 'wrappers'
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null

        Write-ContentWrappers `
            -OutputPath $outDir `
            -InstallPs1Content 'echo install' `
            -UninstallPs1Content 'echo uninstall'
    }

    It 'creates install.bat' {
        Test-Path (Join-Path $outDir 'install.bat') | Should -BeTrue
    }

    It 'creates install.ps1' {
        Test-Path (Join-Path $outDir 'install.ps1') | Should -BeTrue
    }

    It 'creates uninstall.bat' {
        Test-Path (Join-Path $outDir 'uninstall.bat') | Should -BeTrue
    }

    It 'creates uninstall.ps1' {
        Test-Path (Join-Path $outDir 'uninstall.ps1') | Should -BeTrue
    }

    It 'install.bat contains @echo off' {
        $bat = Get-Content (Join-Path $outDir 'install.bat') -Raw
        $bat | Should -Match '@echo off'
    }

    It 'install.bat calls PowerShell.exe with install.ps1' {
        $bat = Get-Content (Join-Path $outDir 'install.bat') -Raw
        $bat | Should -Match 'PowerShell\.exe.*install\.ps1'
    }

    It 'install.bat propagates ERRORLEVEL by default' {
        $bat = Get-Content (Join-Path $outDir 'install.bat') -Raw
        $bat | Should -Match 'exit /b %ERRORLEVEL%'
    }

    It 'install.ps1 contains the provided content' {
        $ps1 = Get-Content (Join-Path $outDir 'install.ps1') -Raw
        $ps1 | Should -Match 'echo install'
    }

    It 'uninstall.ps1 contains the provided content' {
        $ps1 = Get-Content (Join-Path $outDir 'uninstall.ps1') -Raw
        $ps1 | Should -Match 'echo uninstall'
    }

    It 'does not overwrite existing files on second call' {
        # Overwrite install.ps1 with custom content
        Set-Content (Join-Path $outDir 'install.ps1') -Value 'custom' -Encoding ASCII

        Write-ContentWrappers `
            -OutputPath $outDir `
            -InstallPs1Content 'NEW content' `
            -UninstallPs1Content 'NEW uninstall'

        $ps1 = Get-Content (Join-Path $outDir 'install.ps1') -Raw
        $ps1 | Should -Match 'custom'
        $ps1 | Should -Not -Match 'NEW content'
    }

    Context 'custom bat exit codes' {
        BeforeAll {
            $customDir = Join-Path $TestDrive 'custom-exit'
            New-Item -ItemType Directory -Path $customDir -Force | Out-Null

            Write-ContentWrappers `
                -OutputPath $customDir `
                -InstallPs1Content 'echo install' `
                -UninstallPs1Content 'echo uninstall' `
                -InstallBatExitCode '3010' `
                -UninstallBatExitCode '0'
        }

        It 'install.bat uses custom exit code 3010' {
            $bat = Get-Content (Join-Path $customDir 'install.bat') -Raw
            $bat | Should -Match 'exit /b 3010'
        }

        It 'uninstall.bat uses custom exit code 0' {
            $bat = Get-Content (Join-Path $customDir 'uninstall.bat') -Raw
            $bat | Should -Match 'exit /b 0'
        }
    }
}

# ============================================================================
# Write-StageManifest / Read-StageManifest
# ============================================================================

Describe 'Write-StageManifest' {
    It 'writes valid JSON with SchemaVersion and StagedAt' {
        $path = Join-Path $TestDrive 'manifest.json'

        Write-StageManifest -Path $path -ManifestData @{
            AppName         = 'Test App - 1.0'
            Publisher       = 'Test Vendor'
            SoftwareVersion = '1.0'
        }

        Test-Path -LiteralPath $path | Should -BeTrue
        $json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $json.SchemaVersion | Should -Be 1
        $json.StagedAt | Should -Not -BeNullOrEmpty
        $json.AppName | Should -Be 'Test App - 1.0'
        $json.Publisher | Should -Be 'Test Vendor'
    }
}

Describe 'Read-StageManifest' {
    It 'round-trips manifest data correctly' {
        $path = Join-Path $TestDrive 'roundtrip.json'

        $data = @{
            AppName         = 'RoundTrip App - 2.5'
            Publisher       = 'Acme Corp'
            SoftwareVersion = '2.5.0'
            InstallerFile   = 'setup.msi'
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Test\Key'
                ValueName           = 'DisplayVersion'
                ExpectedValue       = '2.5.0.0'
                Operator            = 'IsEquals'
                Is64Bit             = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.AppName         | Should -Be 'RoundTrip App - 2.5'
        $manifest.Publisher       | Should -Be 'Acme Corp'
        $manifest.SoftwareVersion | Should -Be '2.5.0'
        $manifest.InstallerFile   | Should -Be 'setup.msi'
        $manifest.Detection.Type  | Should -Be 'RegistryKeyValue'
        $manifest.Detection.RegistryKeyRelative | Should -Be 'SOFTWARE\Test\Key'
        $manifest.Detection.ExpectedValue | Should -Be '2.5.0.0'
        $manifest.Detection.Operator | Should -Be 'IsEquals'
        $manifest.Detection.Is64Bit | Should -BeTrue
    }

    It 'throws when file does not exist' {
        { Read-StageManifest -Path (Join-Path $TestDrive 'nonexistent.json') } |
            Should -Throw '*not found*'
    }

    It 'throws when JSON is missing SchemaVersion' {
        $path = Join-Path $TestDrive 'bad-manifest.json'
        '{"AppName": "test"}' | Set-Content -LiteralPath $path -Encoding UTF8

        { Read-StageManifest -Path $path } |
            Should -Throw '*missing SchemaVersion*'
    }
}

# ============================================================================
# New-OdtConfigXml
# ============================================================================

Describe 'New-OdtConfigXml' {
    Context 'basic single-product download XML' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '64' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail') `
                -SourcePath 'C:\temp\ap\M365Apps-x64\16.0.19127.20532'
        }

        It 'starts with <Configuration>' {
            $xml | Should -Match '^<Configuration>'
        }

        It 'ends with closing Configuration tag' {
            $xml.TrimEnd() | Should -BeLike '*</Configuration>'
        }

        It 'includes OfficeClientEdition 64' {
            $xml | Should -Match 'OfficeClientEdition="64"'
        }

        It 'includes the version' {
            $xml | Should -Match 'Version="16\.0\.19127\.20532"'
        }

        It 'includes Channel SemiAnnual' {
            $xml | Should -Match 'Channel="SemiAnnual"'
        }

        It 'includes the SourcePath' {
            $xml | Should -Match 'SourcePath="C:\\temp\\ap\\M365Apps-x64\\16\.0\.19127\.20532"'
        }

        It 'includes the product ID' {
            $xml | Should -Match 'Product ID="O365ProPlusRetail"'
        }

        It 'excludes Groove, Lync, OneDrive, Teams, Bing' {
            $xml | Should -Match 'ExcludeApp ID="Groove"'
            $xml | Should -Match 'ExcludeApp ID="Lync"'
            $xml | Should -Match 'ExcludeApp ID="OneDrive"'
            $xml | Should -Match 'ExcludeApp ID="Teams"'
            $xml | Should -Match 'ExcludeApp ID="Bing"'
        }

        It 'includes SharedComputerLicensing' {
            $xml | Should -Match 'Name="SharedComputerLicensing" Value="1"'
        }

        It 'includes FORCEAPPSHUTDOWN' {
            $xml | Should -Match 'Name="FORCEAPPSHUTDOWN" Value="TRUE"'
        }

        It 'includes MigrateArch' {
            $xml | Should -Match 'MigrateArch="TRUE"'
        }

        It 'includes RemoveMSI' {
            $xml | Should -Match '<RemoveMSI />'
        }

        It 'includes Display Level None with AcceptEULA' {
            $xml | Should -Match 'Display Level="None" AcceptEULA="TRUE"'
        }

        It 'includes Logging element' {
            $xml | Should -Match 'Logging Level="Standard"'
        }
    }

    Context 'install XML without SourcePath' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '64' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail')
        }

        It 'does not include SourcePath attribute' {
            $xml | Should -Not -Match 'SourcePath='
        }
    }

    Context 'multi-product XML' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '64' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail', 'VisioProRetail')
        }

        It 'includes both product IDs' {
            $xml | Should -Match 'Product ID="O365ProPlusRetail"'
            $xml | Should -Match 'Product ID="VisioProRetail"'
        }

        It 'each product has its own ExcludeApp entries' {
            # Two sets of ExcludeApp blocks (one per product)
            $grooveMatches = [regex]::Matches($xml, 'ExcludeApp ID="Groove"')
            $grooveMatches.Count | Should -Be 2
        }
    }

    Context 'with CompanyName' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '32' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail') `
                -CompanyName 'Contoso Ltd'
        }

        It 'includes AppSettings block' {
            $xml | Should -Match '<AppSettings>'
        }

        It 'includes Company setup with the provided name' {
            $xml | Should -Match 'Name="Company" Value="Contoso Ltd"'
        }
    }

    Context 'without CompanyName' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '64' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail')
        }

        It 'omits AppSettings block entirely' {
            $xml | Should -Not -Match '<AppSettings>'
            $xml | Should -Not -Match 'Name="Company"'
        }
    }

    Context 'x86 edition' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '32' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail')
        }

        It 'includes OfficeClientEdition 32' {
            $xml | Should -Match 'OfficeClientEdition="32"'
        }
    }
}

# ============================================================================
# Initialize-Folder
# ============================================================================

Describe 'Initialize-Folder' {
    It 'creates a new directory' {
        $dir = Join-Path $TestDrive 'new-folder'
        Initialize-Folder -Path $dir

        Test-Path -LiteralPath $dir | Should -BeTrue
        (Get-Item $dir).PSIsContainer | Should -BeTrue
    }

    It 'does not error when directory already exists' {
        $dir = Join-Path $TestDrive 'existing-folder'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null

        { Initialize-Folder -Path $dir } | Should -Not -Throw
    }

    It 'creates nested directories' {
        $dir = Join-Path $TestDrive 'a\b\c'
        Initialize-Folder -Path $dir

        Test-Path -LiteralPath $dir | Should -BeTrue
    }
}

# ============================================================================
# Get-PackagerPreferences
# ============================================================================

Describe 'Get-PackagerPreferences' {
    It 'reads the actual packager-preferences.json file' {
        $prefsPath = Join-Path $PSScriptRoot 'packager-preferences.json'
        if (-not (Test-Path -LiteralPath $prefsPath)) {
            Set-ItResult -Skipped -Because 'packager-preferences.json not present'
            return
        }

        $prefs = Get-PackagerPreferences
        $prefs | Should -Not -BeNullOrEmpty
        $prefs.PSObject.Properties.Name | Should -Contain 'CompanyName'
    }
}

# ============================================================================
# Write-StageManifest / Read-StageManifest — per-user manifest overrides
# ============================================================================

Describe 'Stage manifest with per-user deployment overrides' {
    It 'round-trips InstallationBehaviorType and LogonRequirementType' {
        $path = Join-Path $TestDrive 'zoom-manifest.json'

        $data = @{
            AppName                  = 'Zoom Workplace - 6.6.0 (x64)'
            Publisher                = 'Zoom Video Communications'
            SoftwareVersion          = '6.6.0'
            InstallerFile            = 'ZoomInstaller.exe'
            InstallationBehaviorType = 'InstallForUser'
            LogonRequirementType     = 'OnlyWhenUserLoggedOn'
            Detection                = @{
                Type         = 'File'
                FilePath     = '%APPDATA%\Zoom\bin'
                FileName     = 'Zoom.exe'
                PropertyType = 'Existence'
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.InstallationBehaviorType | Should -Be 'InstallForUser'
        $manifest.LogonRequirementType     | Should -Be 'OnlyWhenUserLoggedOn'
        $manifest.Detection.Type           | Should -Be 'File'
        $manifest.Detection.FilePath       | Should -Be '%APPDATA%\Zoom\bin'
        $manifest.Detection.FileName       | Should -Be 'Zoom.exe'
        $manifest.Detection.PropertyType   | Should -Be 'Existence'
    }
}

# ============================================================================
# Write-StageManifest — RegistryKeyValue with fixed ARP key
# ============================================================================

Describe 'Stage manifest with fixed ARP key detection' {
    It 'round-trips RegistryKeyValue detection with named key' {
        $path = Join-Path $TestDrive 'vlc-manifest.json'

        $data = @{
            AppName         = 'VLC Media Player - 3.0.23 (x64)'
            Publisher       = 'VideoLAN'
            SoftwareVersion = '3.0.23'
            InstallerFile   = 'vlc-3.0.23-win64.msi'
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VLC media player'
                ValueName           = 'DisplayVersion'
                ExpectedValue       = '3.0.23'
                Operator            = 'IsEquals'
                Is64Bit             = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.Detection.Type | Should -Be 'RegistryKeyValue'
        $manifest.Detection.RegistryKeyRelative | Should -Match 'VLC media player'
        $manifest.Detection.ExpectedValue | Should -Be '3.0.23'
        $manifest.Detection.Operator | Should -Be 'IsEquals'
    }
}
