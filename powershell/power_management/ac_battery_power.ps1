<#
.SYNOPSIS
    Queries the battery status for the local computer or a remote Windows host.

.DESCRIPTION
    Prompts the operator for a computer name (defaulting to the local host when left blank),
    ensures the script is running elevated, queries Win32_Battery information using modern
    CIM cmdlets with a fallback to legacy WMI, formats the results, prints a status legend,
    and writes a timestamped log to the user's temp directory.

.PARAMETER ComputerName
    Optional computer name to query. If omitted, the operator is prompted. Empty input
    targets the local machine. Hostnames are validated prior to execution.

.EXAMPLE
    PS C:\> .\ac_battery_power.ps1

    Prompts for a computer name and displays the battery status.

.EXAMPLE
    PS C:\> .\ac_battery_power.ps1 -ComputerName SRV-102

    Queries the battery status of SRV-102 without prompting.

.NOTES
    ┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
    │ ORIGIN STORY                                                                                │ 
    ├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
    │   NAME:       : ac_battery_power.ps1                                          │
    │   DATE        : 2025-09-23                                                                  │
    │   AUTHOR      : Dallas Bleak (Dallas.Bleak@va.gov)(based on original by John Phung)         │
    │   VERSION     : 2.0                                                                         │    
    │   RUN AS      : Elevated PowerShell (Run as Administrator) recommended.                     │
    └─────────────────────────────────────────────────────────────────────────────────────────────┘ 
#>

[CmdletBinding()]
param(
    [string]$ComputerName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

${script:LogEnabled} = $false
${script:LogPath} = $null

function Invoke-Elevation {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$BoundParameters
    )

    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($currentIdentity)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host 'Elevation required. Relaunching with administrative privileges...' -ForegroundColor Yellow

        $argumentList = @(
            '-NoProfile'
            '-ExecutionPolicy'
            'Bypass'
            '-File'
            "`"$PSCommandPath`""
        )

        foreach ($kvp in $BoundParameters.GetEnumerator()) {
            $key = "-$($kvp.Key)"
            $value = $kvp.Value

            if ($value -is [System.Management.Automation.SwitchParameter]) {
                if ($value.IsPresent) {
                    $argumentList += $key
                }
                continue
            }

            if ($null -ne $value -and $value.ToString().Length -gt 0) {
                $argumentList += $key
                if ($value.ToString().Contains(' ')) {
                    $argumentList += "`"$value`""
                }
                else {
                    $argumentList += $value
                }
            }
        }

        $processPath = (Get-Process -Id $PID).Path
        Start-Process -FilePath $processPath -ArgumentList $argumentList -Verb RunAs | Out-Null
        exit
    }
}

function Resolve-ComputerName {
    param(
        [AllowNull()]
        [string]$InputName
    )

    if ([string]::IsNullOrWhiteSpace($InputName)) {
        return $env:COMPUTERNAME
    }

    $trimmed = $InputName.Trim()
    $lower = $trimmed.ToLowerInvariant()

    if ($lower -in @('.', 'localhost')) {
        return $env:COMPUTERNAME
    }

    if ($trimmed.Length -gt 63 -or -not ($trimmed -match '^[A-Za-z0-9][A-Za-z0-9\-]{0,62}$')) {
        Write-Warning "Invalid computer name '$trimmed'. Use a standard hostname or NetBIOS name."
        return $null
    }

    return $trimmed.ToUpperInvariant()
}

function Initialize-Logging {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Target
    )

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    ${script:LogPath} = Join-Path -Path $env:TEMP -ChildPath ("battery_status_{0}_{1}.log" -f $Target, $timestamp)

    try {
        New-Item -Path ${script:LogPath} -ItemType File -Force | Out-Null
        ${script:LogEnabled} = $true
    }
    catch {
        $pathForWarning = ${script:LogPath}
        Write-Warning ("Unable to initialize log file at '{0}'. Logging disabled. {1}" -f $pathForWarning, $_)
        ${script:LogEnabled} = $false
        ${script:LogPath} = $null
    }
}

function Write-Log {
    param(
        [string[]]$Message
    )

    if (-not ${script:LogEnabled} -or [string]::IsNullOrEmpty(${script:LogPath})) {
        return
    }

    foreach ($line in $Message) {
        Add-Content -Path ${script:LogPath} -Value $line
    }
}

function Write-Message {
    param(
        [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    $lines = $Message -split [Environment]::NewLine
    foreach ($line in $lines) {
        Write-Host $line -ForegroundColor $Color
        Write-Log $line
    }
}

function Write-Banner {
    param(
        [string]$Title
    )

    $line = ('*' * 60)
    Write-Message $line ([ConsoleColor]::Cyan)
    Write-Message ("** {0} **" -f $Title) ([ConsoleColor]::Cyan)
    Write-Message $line ([ConsoleColor]::Cyan)
}

$BatteryStatusDescriptions = @{
    1  = 'Other'
    2  = 'Unknown'
    3  = 'Fully Charged'
    4  = 'Low'
    5  = 'Critical'
    6  = 'Charging'
    7  = 'Charging and High'
    8  = 'Charging and Low'
    9  = 'Charging and Critical'
    10 = 'Undefined'
    11 = 'Partially Charged'
    12 = 'Learning'
    13 = 'Sleeping'
    14 = 'Plugged and Not Charging'
    15 = 'Disconnected'
}

function Get-BatteryStatusDescription {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Code
    )

    if ($BatteryStatusDescriptions.ContainsKey($Code)) {
        return $BatteryStatusDescriptions[$Code]
    }

    return 'Status code not documented'
}

function Show-BatteryLegend {
    Write-Message ''
    Write-Message 'Status Legend:' ([ConsoleColor]::Cyan)
    Write-Message '  1  - Other' ([ConsoleColor]::DarkCyan)
    Write-Message '  2  - Unknown' ([ConsoleColor]::DarkCyan)
    Write-Message '  3  - Fully Charged' ([ConsoleColor]::DarkCyan)
    Write-Message '  4  - Low' ([ConsoleColor]::DarkCyan)
    Write-Message '  5  - Critical' ([ConsoleColor]::DarkCyan)
    Write-Message '  6  - Charging' ([ConsoleColor]::DarkCyan)
    Write-Message '  7  - Charging and High' ([ConsoleColor]::DarkCyan)
    Write-Message '  8  - Charging and Low' ([ConsoleColor]::DarkCyan)
    Write-Message '  9  - Charging and Critical' ([ConsoleColor]::DarkCyan)
    Write-Message ' 10  - Undefined' ([ConsoleColor]::DarkCyan)
    Write-Message ' 11  - Partially Charged' ([ConsoleColor]::DarkCyan)
    Write-Message ' 12  - Learning' ([ConsoleColor]::DarkCyan)
    Write-Message ' 13  - Sleeping' ([ConsoleColor]::DarkCyan)
    Write-Message ' 14  - Plugged and Not Charging' ([ConsoleColor]::DarkCyan)
    Write-Message ' 15  - Disconnected' ([ConsoleColor]::DarkCyan)
}

function Invoke-BatteryQuery {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Target
    )

    try {
        return Get-CimInstance -ClassName Win32_Battery -ComputerName $Target -ErrorAction Stop
    }
    catch {
        Write-Message "Get-CimInstance failed on ${Target}: $($_.Exception.Message)" ([ConsoleColor]::Yellow)
        Write-Message 'Retrying with legacy WMI provider...' ([ConsoleColor]::Yellow)

        try {
            return Get-WmiObject -Class Win32_Battery -ComputerName $Target -ErrorAction Stop
        }
        catch {
            throw
        }
    }
}

function Write-BatteryResults {
    param(
        [Parameter(Mandatory = $true)]
        $BatteryInstances,

        [Parameter(Mandatory = $true)]
        [string]$Target
    )

    if (-not $BatteryInstances) {
        Write-Message "No battery instances were returned for $Target. The device may not contain a battery, or access was denied." ([ConsoleColor]::Yellow)
        return 3
    }

    $output = $BatteryInstances | ForEach-Object {
        [PSCustomObject]@{
            Name                    = $_.Name
            StatusCode              = $_.BatteryStatus
            Status                  = Get-BatteryStatusDescription -Code ($_.BatteryStatus)
            ChargePercent           = if ($_.EstimatedChargeRemaining -ne $null) { "$($_.EstimatedChargeRemaining)%" } else { 'N/A' }
            EstimatedRunTimeMinutes = if ($_.EstimatedRunTime -ne $null -and $_.EstimatedRunTime -gt 0) { $_.EstimatedRunTime } else { 'N/A' }
            Chemistry               = $_.Chemistry
        }
    }

    $table = $output | Format-Table -AutoSize | Out-String
    Write-Message $table ([ConsoleColor]::White)
    return 0
}

Invoke-Elevation -BoundParameters $PSBoundParameters

$resolvedComputer = $null

if ($PSBoundParameters.ContainsKey('ComputerName')) {
    $resolvedComputer = Resolve-ComputerName -InputName $ComputerName
    if (-not $resolvedComputer) {
        exit 1
    }
}
else {
    do {
        $input = Read-Host 'Enter computer name (leave blank for local machine)'
        $resolvedComputer = Resolve-ComputerName -InputName $input
        if (-not $resolvedComputer) {
            Write-Warning 'Please provide a valid computer name.'
        }
    } until ($resolvedComputer)
}

Initialize-Logging -Target $resolvedComputer

if (${script:LogEnabled}) {
    Write-Message ("Logging output to {0}" -f ${script:LogPath}) ([ConsoleColor]::DarkGray)
}
else {
    Write-Message 'Logging disabled; audit file could not be created.' ([ConsoleColor]::DarkYellow)
}

Write-Banner -Title ("Battery Status for {0}" -f $resolvedComputer)
Write-Message ("Target computer: {0}" -f $resolvedComputer) ([ConsoleColor]::Gray)

try {
    $batteryInstances = Invoke-BatteryQuery -Target $resolvedComputer
}
catch {
    Write-Message ("Failed to retrieve battery information from {0}: {1}" -f $resolvedComputer, $_.Exception.Message) ([ConsoleColor]::Red)
    Write-Log ("Query failure: {0}" -f $_.Exception.ToString())
    exit 2
}

$statusCode = Write-BatteryResults -BatteryInstances $batteryInstances -Target $resolvedComputer
Show-BatteryLegend

if ($statusCode -ne 0) {
    Write-Message ("Completed with warnings. Exit code: {0}" -f $statusCode) ([ConsoleColor]::Yellow)
    exit $statusCode
}

Write-Message 'Battery query completed successfully.' ([ConsoleColor]::Green)

if (${script:LogEnabled}) {
    Write-Message ("Battery information logged to {0}" -f ${script:LogPath}) ([ConsoleColor]::DarkGray)
}

exit 0
