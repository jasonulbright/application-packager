@{
    RootModule        = 'AppPackagerCommon.psm1'
    ModuleVersion     = '0.0.8'
    GUID              = 'f5cdd2d6-eb09-47bd-8493-16dfd5666455'
    Author            = 'AppPackager'
    Description       = 'Shared helpers for AppPackager packager scripts.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'

        # Download
        'Invoke-DownloadWithRetry'

        # Environment / pre-flight
        'Test-IsAdmin'
        'Connect-CMSite'
        'Initialize-Folder'
        'Test-NetworkShareAccess'

        # Network path
        'Get-NetworkAppRoot'

        # MSI / ARP
        'Get-MsiPropertyMap'
        'Find-UninstallEntry'

        # Stage manifest
        'Write-StageManifest'
        'Read-StageManifest'

        # Content wrappers
        'Write-ContentWrappers'
        'New-MsiWrapperContent'
        'New-ExeWrapperContent'

        # MECM
        'New-MECMApplicationFromManifest'
        'Remove-CMApplicationRevisionHistoryByCIId'

        # Preferences
        'Get-PackagerPreferences'

        # ODT config XML
        'New-OdtConfigXml'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
