@{
    # Script/Manifest metadata
    RootModule        = 'utilities.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = '2c8f1d9a-3c9a-4d3e-9b7e-6b9a8f7f2c31'  # New-Guid to replace
    Author            = 'Dallas Bleak (Dallas.Bleak@va.gov)'
    CompanyName       = 'Department of Veterans Affairs'
    Copyright         = '© Department of Veterans Affairs'
    Description       = 'Utilities for elevation and structured logging in internal automation scripts.'

    # Compatible PowerShell versions and platforms
    PowerShellVersion = '5.1'          # Works on Windows PowerShell 5.1 and PowerShell 7+
    # PowerShellHostName  = ''
    # PowerShellHostVersion = ''
    # DotNetFrameworkVersion = ''
    # CLRVersion            = ''
    # ProcessorArchitecture = ''
    # RequiredAssemblies    = @()
    # RequiredModules       = @()
    # ExternalModuleDependencies = @()

    # Export controls (keep tight for perf and clarity)
    FunctionsToExport = @(
        'Invoke-Elevation'
        'New-LogSession'
        'Write-LogEntry'
        'Write-LogSection'
    )
    CmdletsToExport   = @()            # none
    VariablesToExport = @()            # none
    AliasesToExport   = @(
        'Write-Log'
    )

    # NestedModules can include additional psm1/psd1 files if you split functionality
    # NestedModules = @()

    # Private data (gallery metadata, tags, etc.)
    PrivateData       = @{
        PSData = @{
            Tags         = @('utilities', 'elevation', 'admin', 'logging', 'log', 'write-log')
            ProjectUri   = ''
            LicenseUri   = ''
            IconUri      = ''
            ReleaseNotes = 'Version 1.1.0: Adds structured logging helpers (New-LogSession, Write-LogEntry, Write-LogSection). Version 1.0.0: Initial release with Invoke-Elevation.'
        }
    }

    # Optional script blocks on import/remove
    # ScriptsToProcess  = @()
    # TypesToProcess    = @()
    # FormatsToProcess  = @()
    # FileList          = @('utilities.psm1')
    # ModuleList        = @()
    # DscResourcesToExport = @()
}
