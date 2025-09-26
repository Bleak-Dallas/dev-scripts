<#
.SYNOPSIS
    Reusable PowerShell utilities for elevation, logging, and common script helpers.

.DESCRIPTION
    The Utilities module provides functions intended to be reused across administrative
    and automation scripts. It emphasizes consistency, reliability, and secure defaults.
    Typical use cases include:
      - Elevating the current script/process with administrative privileges while
        preserving bound parameters.
      - Standardized logging patterns (file and console).
      - Input validation, normalization, and common helper operations.

.NOTES
    Module Name   : utilities
    Root Module   : utilities.psm1
    Manifest      : utilities.psd1
    Version       : 1.0.0
    Author        : Dallas Bleak (Dallas.Bleak@va.gov)
    Company       : Department of Veterans Affairs / OI&T / EUS / SLC
    Copyright     : © 2025 Department of Veterans Affairs. All rights reserved.
    License       : INTERNAL
    Source        : C:\Users\VHASLCBleakD\OneDrive - Department of Veterans Affairs\git-dev-scripts\powershell\modules\utilities
    Project URI   : REPOSITORY OR DOCS URL
    Tags          : utilities, elevation, logging, admin, helpers

.RELEASE NOTES
    1.0.0
      - Initial release with Invoke-Elevation for safe re-launch with elevation.
      - Module scaffolding and manifest with controlled exports.

.REQUIREMENTS
    - PowerShell 5.1 or later (Windows PowerShell 5.1 and PowerShell 7+ supported).
    - Windows for elevation scenarios requiring UAC (Invoke-Elevation).
    - Appropriate execution policy and permissions to import/run the module.

.SUPPORTED PLATFORMS
    - Windows 10/11, Windows Server 2016+
    - PowerShell 7+ cross-platform compatibility varies by function; elevation
      requires Windows.

.INSTALLATION
    Option A — User scope (recommended during development):
        Copy the entire 'utilities' module folder to:
            PowerShell 7+        : $HOME\Documents\PowerShell\Modules\utilities
            Windows PowerShell   : $HOME\Documents\WindowsPowerShell\Modules\utilities

    Option B — System scope (requires admin):
        Copy to:
            Windows: C:\Program Files\PowerShell\Modules\utilities

    Option C — Import by full path:
        Import-Module "C:\path\to\utilities\utilities.psm1" -Force

.IMPORTING
    # By name (if in $env:PSModulePath)
    Import-Module utilities

    # By path
    Import-Module 'C:\Users\VHASLCBleakD\OneDrive - Department of Veterans Affairs\git-dev-scripts\powershell\modules\utilities\utilities.psm1' -Force

.EXPORTS
    Public functions are controlled via the module manifest (utilities.psd1)
    in FunctionsToExport. Keep internal helpers unexported or in a Private folder.

.KEY FUNCTIONS
    Invoke-Elevation
        Re-launches the current script with administrative privileges while preserving
        bound parameters. Uses the current host process path to maintain pwsh/powershell
        consistency.
    New-LogSession
        Creates a structured log session object, ensuring directory creation, standardized
        filenames, and optional header/metadata emission.
    Write-LogEntry
        Writes timestamped log entries with severity/color handling to both file and
        console (optional), supporting shared usage across scripts.
    Write-LogSection
        Writes grouped log entries (title plus items) for collections, leveraging
        Write-LogEntry for consistent formatting.

.USAGE EXAMPLES
    # Ensure a script runs elevated and then proceed
    Invoke-Elevation -BoundParameters $PSBoundParameters
    # Continue with elevated-only operations...

    # Standardized logging with reusable helpers
    $logSession = New-LogSession -LogNamePrefix 'my_script' -Metadata @{ Host = $env:COMPUTERNAME }
    Write-LogEntry -Session $logSession -Message 'Script starting.'
    Write-LogSection -Session $logSession -Title 'Processing items' -Items $items -ItemFormatter { param($item) "Processed $item" }

.ERROR HANDLING
    Functions should throw terminating errors for unrecoverable conditions and
    write structured messages for logging. Consumers may use try/catch to handle
    errors and implement retry or fallback logic.

.LOGGING
    Standardized logging helpers:
      - New-LogSession to bootstrap a session and obtain the target log path.
      - Write-LogEntry (alias Write-Log) for individual messages with severity/color control.
      - Write-LogSection for grouped entries (e.g., profile inventories, summary tables).
    Avoid logging sensitive information (passwords, secrets, tokens) to disk.

.SECURITY
    - Invoke-Elevation uses Start-Process -Verb RunAs to prompt for UAC elevation.
    - Avoid logging sensitive information (passwords, secrets, tokens).
    - Validate and constrain user input where applicable.

.VERSIONING
    Follow semantic versioning:
      MAJOR: breaking changes
      MINOR: new functionality, backward compatible
      PATCH: fixes and internal improvements

.CONTRIBUTING
    - Use approved PowerShell verbs and singular nouns where practical.
    - Include Pester tests for new functionality where possible.
    - Run PSScriptAnalyzer and fix violations prior to commit.
    - Update the manifest (utilities.psd1) and this header when adding public functions.

#>

function Invoke-Elevation {
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.Collections.IDictionary] $BoundParameters,
        [Parameter()]
        [string] $ScriptPath
    )

    if (-not $BoundParameters) {
        $BoundParameters = @{}
    }

    if (-not $ScriptPath) {
        $ScriptPath = $MyInvocation.PSCommandPath
        if (-not $ScriptPath -and $PSCommandPath) {
            $ScriptPath = $PSCommandPath
        }
    }

    if (-not $ScriptPath) {
        throw 'Invoke-Elevation: Unable to determine script path for elevation.'
    }

    try {
        $ScriptPath = (Resolve-Path -Path $ScriptPath -ErrorAction Stop).Path
    }
    catch {
        throw "Invoke-Elevation: Unable to resolve script path '$ScriptPath'. $_"
    }

    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($currentIdentity)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host 'Elevation required. Relaunching with administrative privileges...' -ForegroundColor Yellow
        # Relaunch with the same host (pwsh or powershell)
        $processPath = (Get-Process -Id $PID).Path

        $argumentList = @(
            '-NoProfile'
            '-ExecutionPolicy', 'Bypass'
            '-File', "`"$ScriptPath`""
        )

        foreach ($entry in $BoundParameters.GetEnumerator()) {
            $key = "-$($entry.Key)"
            $value = $entry.Value

            if ($null -eq $value) { continue }

            # Switch parameters
            if ($value -is [System.Management.Automation.SwitchParameter]) {
                if ($value.IsPresent) { $argumentList += $key }
                continue
            }

            # Array parameters -> repeat key for each value
            if ($value -is [System.Array]) {
                foreach ($item in $value) {
                    if ($null -ne $item -and $item -ne '') {
                        $argumentList += $key
                        $argumentList += "`"$item`""
                    }
                }
                continue
            }

            # Scalar values
            $stringValue = [string]$value
            if ($stringValue.Length -eq 0) { continue }
            $argumentList += $key
            $argumentList += "`"$stringValue`""
        }

        Start-Process -FilePath $processPath -ArgumentList $argumentList -Verb RunAs | Out-Null
        exit
    }
}

<#
.SYNOPSIS
    Creates a standardized log session object for reuse across scripts.

.DESCRIPTION
    Ensures the log directory exists, builds a timestamped log filename, optionally
    writes a header and metadata, and returns a PSCustomObject representing the log
    session. Consumers should pass the returned object to Write-LogEntry / Write-LogSection
    (or the Write-Log alias) when emitting log entries.

.PARAMETER LogRoot
    Root directory where log files are stored. Created if it does not exist.

.PARAMETER LogNamePrefix
    Prefix used when generating the log filename (before the timestamp).

.PARAMETER LogFileName
    Explicit filename (with or without extension). When omitted, a timestamped name
    is generated using the prefix and extension.

.PARAMETER TimestampFormat
    Format string used when generating the timestamp for the filename.

.PARAMETER Extension
    File extension (defaults to .log). Dots are optional.

.PARAMETER Header
    Optional header line to write at the top of the log; defaults to a generic session
    banner if not provided.

.PARAMETER Metadata
    Hashtable of metadata entries included in the header for quick reference.

.PARAMETER NoHeader
    Skip header emission entirely but still create the file.

.EXAMPLE
    $session = New-LogSession -LogNamePrefix 'inventory' -Metadata @{ Host = $env:COMPUTERNAME }
#>
function New-LogSession {
    [CmdletBinding()]
    param(
        [string]$LogRoot = 'C:\Temp\logs',
        [string]$LogNamePrefix = 'script',
        [string]$LogFileName,
        [string]$TimestampFormat = 'yyyyMMdd_HHmmss',
        [string]$Extension = 'log',
        [string]$Header,
        [hashtable]$Metadata,
        [switch]$NoHeader
    )

    try {
        if (-not (Test-Path -LiteralPath $LogRoot)) {
            New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
        }
    }
    catch {
        throw "New-LogSession: Unable to prepare log directory '$LogRoot'. $_"
    }

    $normalizedExtension = $Extension
    if ([string]::IsNullOrWhiteSpace($normalizedExtension)) {
        $normalizedExtension = 'log'
    }
    $normalizedExtension = $normalizedExtension.TrimStart('.')

    if ([string]::IsNullOrWhiteSpace($LogFileName)) {
        $prefix = if ([string]::IsNullOrWhiteSpace($LogNamePrefix)) { 'script' } else { $LogNamePrefix }
        $safePrefix = ($prefix -replace '[^\w\-.]', '_')
        $timestamp = Get-Date -Format $TimestampFormat
        $LogFileName = '{0}_{1}.{2}' -f $safePrefix, $timestamp, $normalizedExtension
    }
    else {
        $currentExtension = [System.IO.Path]::GetExtension($LogFileName)
        if ([string]::IsNullOrWhiteSpace($currentExtension)) {
            $LogFileName = '{0}.{1}' -f $LogFileName, $normalizedExtension
        }
    }

    $logPath = Join-Path -Path $LogRoot -ChildPath $LogFileName
    $sessionStart = Get-Date

    try {
        if ($NoHeader) {
            New-Item -Path $logPath -ItemType File -Force | Out-Null
        }
        else {
            $headerLines = @()
            if ($Header) {
                $headerLines += $Header
            }
            else {
                $headerLines += "===== Log session started $sessionStart ====="
            }

            if ($Metadata) {
                foreach ($key in ($Metadata.Keys | Sort-Object)) {
                    $headerLines += "Meta[{0}] = {1}" -f $key, $Metadata[$key]
                }
            }

            Set-Content -Path $logPath -Value $headerLines -Encoding UTF8
        }
    }
    catch {
        throw "New-LogSession: Unable to initialize log file '$logPath'. $_"
    }

    $session = [pscustomobject]@{
        LogPath      = $logPath
        LogDirectory = $LogRoot
        LogFileName  = $LogFileName
        Created      = $sessionStart
        Metadata     = $Metadata
    }
    $session.PSObject.TypeNames.Insert(0, 'Utilities.LogSession')

    return $session
}

<#
.SYNOPSIS
    Writes a timestamped log entry to a session or explicit log path.

.DESCRIPTION
    Appends a formatted log line including timestamp and severity, with optional console
    output colored based on severity. Supports being called with the session object
    returned from New-LogSession or directly with a log file path.

.PARAMETER Session
    Log session object returned from New-LogSession. Supplies the target log path.

.PARAMETER LogPath
    Explicit log file path to append to. Use when a session object is not available.

.PARAMETER Message
    Message text to record.

.PARAMETER Severity
    Log severity classification (Info/Success/Warning/Error/Verbose/Debug).

.PARAMETER Color
    Console color override. When omitted, a default color based on severity is chosen.

.PARAMETER NoConsole
    Suppresses console output while still appending to the log file.

.EXAMPLE
    Write-LogEntry -Session $session -Message 'Script starting.' -Severity Info
#>
function Write-LogEntry {
    [CmdletBinding(DefaultParameterSetName = 'BySession')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'BySession')]
        [ValidateNotNull()]
        [pscustomobject]$Session,

        [Parameter(Mandatory, ParameterSetName = 'ByPath')]
        [ValidateNotNullOrEmpty()]
        [string]$LogPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Verbose', 'Debug')]
        [string]$Severity = 'Info',

        [System.ConsoleColor]$Color,

        [switch]$NoConsole
    )

    $targetPath = if ($PSCmdlet.ParameterSetName -eq 'BySession') {
        if (-not $Session.LogPath) {
            throw 'Write-LogEntry: The supplied session object does not expose a LogPath property.'
        }
        $Session.LogPath
    }
    else {
        $LogPath
    }

    try {
        if (-not (Test-Path -LiteralPath $targetPath)) {
            New-Item -Path $targetPath -ItemType File -Force | Out-Null
        }
    }
    catch {
        throw "Write-LogEntry: Unable to ensure log file '$targetPath'. $_"
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $upperSeverity = $Severity.ToUpperInvariant()
    $line = '[{0}][{1}] {2}' -f $timestamp, $upperSeverity, $Message

    try {
        Add-Content -Path $targetPath -Value $line -Encoding UTF8
    }
    catch {
        throw "Write-LogEntry: Unable to append to log file '$targetPath'. $_"
    }

    if (-not $NoConsole.IsPresent) {
        $resolvedColor = if ($PSBoundParameters.ContainsKey('Color')) {
            $Color
        }
        else {
            switch ($upperSeverity) {
                'SUCCESS' { [System.ConsoleColor]::Green }
                'WARNING' { [System.ConsoleColor]::Yellow }
                'ERROR' { [System.ConsoleColor]::Red }
                'VERBOSE' { [System.ConsoleColor]::DarkGray }
                'DEBUG' { [System.ConsoleColor]::Magenta }
                default { [System.ConsoleColor]::Gray }
            }
        }

        Write-Host ("[{0}] {1}" -f $upperSeverity, $Message) -ForegroundColor $resolvedColor
    }

    return [pscustomobject]@{
        Timestamp = $timestamp
        Severity  = $upperSeverity
        Message   = $Message
        LogPath   = $targetPath
    }
}

<#
.SYNOPSIS
    Emits a titled log section followed by individual item entries.

.DESCRIPTION
    Uses Write-LogEntry internally to write a heading line and a series of bullet
    lines representing each supplied item. Useful for summarizing collections
    (profiles, accounts, results) in a consistent format.

.PARAMETER Title
    Heading for the section.

.PARAMETER Items
    Objects or strings to be logged. When omitted or empty, a "(none)" entry is logged.

.PARAMETER ItemFormatter
    Optional script block that formats each item. Receives the item as $_ / $args[0].

.PARAMETER Severity
    Severity classification for the section heading.

.PARAMETER ItemSeverity
    Severity classification for each item entry. Defaults to the section severity.

.PARAMETER NoConsole
    Suppresses console output for both the heading and the items.

.PARAMETER Session
    Log session object produced by New-LogSession.

.PARAMETER LogPath
    Explicit log file path (alternative to Session when a log session is not used).

.EXAMPLE
    Write-LogSection -Session $session -Title 'Processed profiles' -Items $profiles `
        -ItemFormatter { param($p) "$($p.Name) - $($p.Status)" }
#>
function Write-LogSection {
    [CmdletBinding(DefaultParameterSetName = 'BySession')]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [object[]]$Items,

        [ScriptBlock]$ItemFormatter,

        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Verbose', 'Debug')]
        [string]$Severity = 'Info',

        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Verbose', 'Debug')]
        [string]$ItemSeverity,

        [switch]$NoConsole,

        [Parameter(Mandatory, ParameterSetName = 'BySession')]
        [pscustomobject]$Session,

        [Parameter(Mandatory, ParameterSetName = 'ByPath')]
        [string]$LogPath
    )

    $params = @{
        Message   = $Title
        Severity  = $Severity
        NoConsole = $NoConsole
    }

    if ($PSCmdlet.ParameterSetName -eq 'BySession') {
        $params['Session'] = $Session
    }
    else {
        $params['LogPath'] = $LogPath
    }

    Write-LogEntry @params | Out-Null

    $effectiveItemSeverity = if ($PSBoundParameters.ContainsKey('ItemSeverity')) { $ItemSeverity } else { $Severity }

    if ($Items -and $Items.Count -gt 0) {
        foreach ($item in $Items) {
            $formatted = if ($ItemFormatter) {
                & $ItemFormatter $item
            }
            else {
                [string]$item
            }

            if ([string]::IsNullOrWhiteSpace($formatted)) {
                $formatted = '(empty)'
            }

            $entryParams = @{
                Message   = " - $formatted"
                Severity  = $effectiveItemSeverity
                NoConsole = $true
            }

            if ($PSCmdlet.ParameterSetName -eq 'BySession') {
                $entryParams['Session'] = $Session
            }
            else {
                $entryParams['LogPath'] = $LogPath
            }

            Write-LogEntry @entryParams | Out-Null
        }
    }
    else {
        $entryParams = @{
            Message   = ' - (none)'
            Severity  = $effectiveItemSeverity
            NoConsole = $true
        }

        if ($PSCmdlet.ParameterSetName -eq 'BySession') {
            $entryParams['Session'] = $Session
        }
        else {
            $entryParams['LogPath'] = $LogPath
        }

        Write-LogEntry @entryParams | Out-Null
    }
}

Set-Alias -Name Write-Log -Value Write-LogEntry
