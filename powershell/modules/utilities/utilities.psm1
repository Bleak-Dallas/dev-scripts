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

.USAGE EXAMPLES
    # Ensure a script runs elevated and then proceed
    Invoke-Elevation -BoundParameters $PSBoundParameters
    # Continue with elevated-only operations...

.ERROR HANDLING
    Functions should throw terminating errors for unrecoverable conditions and
    write structured messages for logging. Consumers may use try/catch to handle
    errors and implement retry or fallback logic.

.LOGGING
    If logging helpers are included in this module, prefer:
      - Write-Log for structured messages (Info/Warn/Error/Debug)
      - Avoid Write-Host for operational logs; reserve it for user-facing UX.

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
        [System.Collections.IDictionary] $BoundParameters
    )

    if (-not $BoundParameters) {
        $BoundParameters = @{}
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
            '-File', "`"$PSCommandPath`""
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