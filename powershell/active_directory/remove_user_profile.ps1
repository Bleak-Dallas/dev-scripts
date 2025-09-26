[CmdletBinding()]
param(
    [string]$ComputerName,
    [string[]]$AccountsToKeep = @()
)

<#
.SYNOPSIS
    Remotely prune user profiles from a Windows workstation while honoring a caller-supplied
    preservation list and producing an auditable activity log.

.DESCRIPTION
    1. Loads the shared utilities module (from local or network locations) and relaunches the script
       with elevation if required.
    2. Validates connectivity to the target computer, inventories existing user profiles, and prompts
       for keep-list entries when none were provided.
    3. Normalizes keep entries (names/SIDs), surfaces any unmatched accounts, and requests explicit
       operator confirmation before continuing.
    4. Invokes a remote removal routine that deletes profiles not included in the keep lists, capturing
       structured results (removed vs. skipped) and writing before/after snapshots to the log.
    5. Writes a fully timestamped log to C:\Temp\logs using the shared logging helpers so administrators
       can review the run end-to-end.

.PARAMETER ComputerName
    Optional. The remote computer whose local profiles should be evaluated. Prompts if omitted or blank.

.PARAMETER AccountsToKeep
    Optional. Array of profile names to preserve. Names are normalized (case-insensitive) and mapped to SIDs
    when possible. A prompt is shown if no keep list is supplied.

.INPUTS
    None. All input is provided via parameters or interactive prompts.

.OUTPUTS
    None. Operational output is written to the console for user feedback and to a log file for auditability.

.NOTES
    ┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
    │ ORIGIN STORY                                                                                │ 
    ├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
    │   DATE        : 2025-09-26                                                                  │
    │   AUTHOR      : Dallas Bleak  (Dallas.Bleak@va.gov)                                         │
    │   VERSION     : 1.0                                                                         │
    │   Run As      : Elevated PowerShell (Run as Administrator) recommended.                     │
    └─────────────────────────────────────────────────────────────────────────────────────────────┘

.EXAMPLE
    PS> .\remove_user_profile.ps1 -ComputerName "SLC-WS12345" -AccountsToKeep @('admin', 'tech')

    Removes all non-system profiles from SLC-WS12345 except "admin" and "tech", recording actions in C:\Temp\logs.
#>

# region Module bootstrap and elevation
$scriptDirectory = Split-Path -Path $PSCommandPath -Parent
$localModulePath = Join-Path -Path $scriptDirectory -ChildPath '..\modules\utilities\utilities.psm1'

$moduleCandidates = @(
    $localModulePath
    '\\va.gov\cn\Salt Lake City\VHASLC\TechDrive\Scripts\Dallas\helper_modules\utilities\utilities.psm1'
)

$importErrors = @()
$moduleLoaded = $false
$UtilitiesModulePath = $null

foreach ($candidate in $moduleCandidates) {
    $resolvedCandidate = $candidate
    try {
        $resolvedCandidate = (Resolve-Path -Path $candidate -ErrorAction Stop).Path
    }
    catch {
        $importErrors += "Path resolution failed for '$candidate'. $_"
        continue
    }

    try {
        Write-Host ("Attempting to load utilities module from '{0}'" -f $resolvedCandidate) -ForegroundColor DarkCyan
        Import-Module -Name $resolvedCandidate -Force -ErrorAction Stop
        $moduleLoaded = $true
        $UtilitiesModulePath = $resolvedCandidate
        break
    }
    catch {
        $importErrors += "Import failed for '$resolvedCandidate'. $_"
    }
}

if (-not $moduleLoaded) {
    Write-Error ("Unable to load helper module from any candidate path.`n{0}" -f ($importErrors -join "`n"))
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ("Loaded utilities module from '{0}'." -f $UtilitiesModulePath) -ForegroundColor Green

Invoke-Elevation -BoundParameters $PSBoundParameters -ScriptPath $PSCommandPath

<#
Logging helpers from utilities module:
  - Use New-LogSession (via Initialize-LogSession below) to centralize log file creation and metadata.
  - Use Write-Log (alias of Write-LogEntry) with the returned session for single-line messages and severity coloring.
  - Use Write-LogSection for grouped entries; helper wrappers below demonstrate common patterns for profile data.
#>

# region User input and discovery helpers
function Read-ComputerName {
    param(
        [string]$InitialName
    )

    $name = $InitialName
    while ([string]::IsNullOrWhiteSpace($name)) {
        $name = Read-Host "Enter the target computer name"
        if ([string]::IsNullOrWhiteSpace($name)) {
            Write-Host "Computer name cannot be empty. Please try again." -ForegroundColor Yellow
        }
    }

    return $name.Trim()
}

function Test-ComputerConnectivity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    Write-Host ""
    Write-Host "Testing connectivity to $ComputerName..." -ForegroundColor Cyan
    try {
        if (-not (Test-Connection -ComputerName $ComputerName -Count 3 -Quiet -ErrorAction Stop)) {
            throw "Computer '$ComputerName' is not reachable."
        }
    }
    catch {
        throw "Unable to reach '$ComputerName'. $_"
    }
}

function Get-RemoteProfiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    $scriptBlock = {
        $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
        Where-Object { -not $_.Special -and $_.LocalPath }

        $results = @()

        foreach ($userProfile in $profiles) {
            $profileName = Split-Path $userProfile.LocalPath -Leaf
            if (-not $profileName) {
                continue
            }

            $lastUse = $null
            $rawLastUse = $userProfile.LastUseTime
            if ($rawLastUse) {
                try {
                    $lastUse = [Management.ManagementDateTimeConverter]::ToDateTime($rawLastUse)
                }
                catch {
                    $lastUse = $null
                }
            }

            $results += [pscustomobject]@{
                ProfileName = $profileName
                SID         = $userProfile.SID
                LocalPath   = $userProfile.LocalPath
                LastUseTime = $lastUse
                Loaded      = $userProfile.Loaded
            }
        }

        return $results
    }

    try {
        return Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction Stop
    }
    catch {
        throw "Failed to retrieve profiles from '$ComputerName' using the current credentials. $_"
    }
}

# Logging helper wrappers that tailor the shared utilities module to this script's needs
function Initialize-LogSession {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    $metadata = @{
        ComputerName   = $ComputerName
        InitiatingHost = $env:COMPUTERNAME
        InitiatingUser = $env:USERNAME
    }

    $header = "===== remove_user_profile run {0} for {1} =====" -f (Get-Date), $ComputerName

    return New-LogSession -LogNamePrefix 'remove_user_profile' -Header $header -Metadata $metadata
}

function Write-ProfileSection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [object[]]$Profiles,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Session,
        [switch]$NoConsole
    )

    Write-LogSection -Session $Session -Title $Title -Items $Profiles -ItemFormatter {
        param($profileRecord)

        if (-not $profileRecord) { return '(null)' }

        $lastUseStamp = if ($profileRecord.LastUseTime) { $profileRecord.LastUseTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'n/a' }
        "{0} (SID: {1}; Path: {2}; Loaded: {3}; LastUse: {4})" -f $profileRecord.ProfileName, $profileRecord.SID, $profileRecord.LocalPath, $profileRecord.Loaded, $lastUseStamp
    } -NoConsole:$NoConsole.IsPresent
}

function Write-AccountsToKeepLog {
    param(
        [string[]]$Accounts,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Session
    )

    Write-LogSection -Session $Session -Title 'Accounts preserved:' -Items $Accounts -ItemFormatter { param($account) $account } -NoConsole
}

function Write-RemovalResultsLog {
    param(
        [object[]]$RemovedProfiles = @(),
        [object[]]$SkippedProfiles = @(),
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Session
    )

    $removedCount = ($RemovedProfiles | Measure-Object).Count
    Write-Log -Session $Session -Message ("Removed {0} profile(s)." -f $removedCount) -Severity 'Info' -NoConsole
    Write-LogSection -Session $Session -Title 'Removed profiles' -Items $RemovedProfiles -ItemFormatter {
        param($record)
        "{0} (SID: {1}; Path: {2}; Loaded: {3}; Reason: {4})" -f $record.ProfileName, $record.SID, $record.LocalPath, $record.Loaded, $record.Reason
    } -NoConsole

    $skippedCount = ($SkippedProfiles | Measure-Object).Count
    Write-Log -Session $Session -Message ("Skipped {0} profile(s)." -f $skippedCount) -Severity 'Info' -NoConsole
    Write-LogSection -Session $Session -Title 'Skipped profiles' -Items $SkippedProfiles -ItemFormatter {
        param($record)
        "{0} (Reason: {1}; SID: {2}; Path: {3}; Loaded: {4})" -f $record.ProfileName, $record.Reason, $record.SID, $record.LocalPath, $record.Loaded
    } -NoConsole
}

function Get-AccountsToKeep {
    param(
        [string[]]$ExistingAccounts = @()
    )

    $accounts = @()
    if ($ExistingAccounts) {
        $accounts += $ExistingAccounts
    }

    do {
        $inputValue = Read-Host "Type each username you want to keep. Type 'END' when finished"
        if ($null -eq $inputValue) {
            continue
        }

        $trimmed = $inputValue.Trim()
        if ($trimmed.Length -eq 0) {
            continue
        }

        if ($trimmed.Equals('END', [System.StringComparison]::InvariantCultureIgnoreCase)) {
            break
        }

        $accounts += $trimmed
    } while ($true)

    return $accounts
}

function ConvertTo-NormalizedAccounts {
    param(
        [string[]]$Accounts
    )

    $normalized = $Accounts |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    ForEach-Object { $_.Trim() }

    if (-not $normalized) {
        return @()
    }

    return , ($normalized | Sort-Object { $_.ToLowerInvariant() } -Unique)
}

function Show-KeepList {
    param(
        [string[]]$AccountsToKeep
    )

    Write-Host ""
    Write-Host "You are going to delete all user profiles except:" -ForegroundColor Cyan

    if (-not $AccountsToKeep -or $AccountsToKeep.Count -eq 0) {
        Write-Host "  (none specified)" -ForegroundColor DarkYellow
    }
    else {
        foreach ($account in $AccountsToKeep) {
            Write-Host ("  {0}" -f $account)
        }
    }

    Write-Host ""
}

function Confirm-Operation {
    param(
        [string]$ComputerName,
        [string[]]$AccountsToKeep
    )

    Show-KeepList -AccountsToKeep $AccountsToKeep

    while ($true) {
        Write-Host "Do you want to continue removing profiles from '$ComputerName'?" -ForegroundColor Cyan
        Write-Host "1: Yes, continue"
        Write-Host "2: No, cancel"
        $selection = Read-Host "Please make a selection"

        switch ($selection) {
            '1' { return $true }
            '2' { return $false }
            default {
                Clear-Host
                Write-Host "Invalid selection. Please choose 1 or 2." -ForegroundColor Yellow
                Show-KeepList -AccountsToKeep $AccountsToKeep
            }
        }
    }
}
# endregion

# region Remote profile evaluation and removal helpers
function Remove-RemoteProfiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [string[]]$AccountsToKeep = @(),
        [string[]]$AccountSidsToKeep = @()
    )

    $normalizedKeepLower = @()
    if ($AccountsToKeep) {
        $normalizedKeepLower = @($AccountsToKeep | ForEach-Object { $_.ToLowerInvariant() })
    }

    $normalizedKeepSids = @()
    if ($AccountSidsToKeep) {
        $normalizedKeepSids = @($AccountSidsToKeep | ForEach-Object { $_.ToUpperInvariant() })
    }

    $scriptBlock = {
        param(
            [string[]]$keepListLower,
            [string[]]$keepSidListUpper
        )

        if (-not $keepListLower) { $keepListLower = @() }
        if (-not $keepSidListUpper) { $keepSidListUpper = @() }

        $removed = @()
        $skipped = @()

        $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
        Where-Object { -not $_.Special -and $_.LocalPath }

        foreach ($remoteProfile in $profiles) {
            $profileName = Split-Path $remoteProfile.LocalPath -Leaf
            if (-not $profileName) {
                continue
            }

            $profileNameLower = $profileName.ToLowerInvariant()
            $rawSid = $remoteProfile.SID
            $sidUpper = if ($rawSid) { $rawSid.ToUpperInvariant() } else { $null }

            if ($sidUpper -and $keepSidListUpper -contains $sidUpper) {
                $skipped += [pscustomobject]@{
                    Action      = 'Skipped'
                    Reason      = 'KeepListSid'
                    ProfileName = $profileName
                    SID         = $remoteProfile.SID
                    LocalPath   = $remoteProfile.LocalPath
                    Loaded      = $remoteProfile.Loaded
                }
                continue
            }

            if ($keepListLower -contains $profileNameLower) {
                $skipped += [pscustomobject]@{
                    Action      = 'Skipped'
                    Reason      = 'KeepListName'
                    ProfileName = $profileName
                    SID         = $remoteProfile.SID
                    LocalPath   = $remoteProfile.LocalPath
                    Loaded      = $remoteProfile.Loaded
                }
                continue
            }

            if ($remoteProfile.SID -match '^S-1-5-(18|19|20)$') {
                $skipped += [pscustomobject]@{
                    Action      = 'Skipped'
                    Reason      = 'SystemSid'
                    ProfileName = $profileName
                    SID         = $remoteProfile.SID
                    LocalPath   = $remoteProfile.LocalPath
                    Loaded      = $remoteProfile.Loaded
                }
                continue
            }

            if ($remoteProfile.Loaded) {
                $skipped += [pscustomobject]@{
                    Action      = 'Skipped'
                    Reason      = 'ProfileLoaded'
                    ProfileName = $profileName
                    SID         = $remoteProfile.SID
                    LocalPath   = $remoteProfile.LocalPath
                    Loaded      = $remoteProfile.Loaded
                }
                continue
            }

            Remove-CimInstance -InputObject $remoteProfile -ErrorAction Stop

            $removed += [pscustomobject]@{
                Action      = 'Removed'
                Reason      = 'Removed'
                ProfileName = $profileName
                SID         = $remoteProfile.SID
                LocalPath   = $remoteProfile.LocalPath
                Loaded      = $remoteProfile.Loaded
            }
        }

        return [pscustomobject]@{
            Removed = $removed
            Skipped = $skipped
        }
    }

    try {
        return Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $normalizedKeepLower, $normalizedKeepSids -ErrorAction Stop
    }
    catch {
        throw "Failed to remove profiles from '$ComputerName' using the current credentials. $_"
    }
}

try {
    # Gather target computer information and ensure prerequisites are satisfied
    $ComputerName = Read-ComputerName -InitialName $ComputerName
    Test-ComputerConnectivity -ComputerName $ComputerName

    # Start a structured logging session and surface the log location
    $logSession = Initialize-LogSession -ComputerName $ComputerName
    $logPath = $logSession.LogPath
    Write-Log -Session $logSession -Message ("Log file: {0}" -f $logPath) -Severity 'Info' -Color ([System.ConsoleColor]::Cyan)

    Write-Host ""
    # Capture the current profile inventory prior to any modifications for before/after auditing
    Write-Host "Retrieving user profile list from '$ComputerName' using the current elevated credentials..." -ForegroundColor Cyan
    $existingProfiles = Get-RemoteProfiles -ComputerName $ComputerName
    Write-ProfileSection -Title "Profiles before removal" -Profiles $existingProfiles -Session $logSession -NoConsole

    if ($existingProfiles -and $existingProfiles.Count -gt 0) {
        Write-Host "Current profiles found on '$ComputerName':" -ForegroundColor Cyan
        $existingProfiles | Sort-Object ProfileName | Format-Table ProfileName, SID, LocalPath, LastUseTime, Loaded -AutoSize
    }
    else {
        Write-Host "No removable user profiles were found on '$ComputerName'." -ForegroundColor Yellow
    }

    Write-Host ""

    # Ensure we have an actionable keep-list (prompting interactively when one was not provided)
    if (-not $PSBoundParameters.ContainsKey('AccountsToKeep') -or $AccountsToKeep.Count -eq 0) {
        $AccountsToKeep = Get-AccountsToKeep
    }

    $AccountsToKeep = ConvertTo-NormalizedAccounts -Accounts $AccountsToKeep
    Write-AccountsToKeepLog -Accounts $AccountsToKeep -Session $logSession

    $AccountsToKeepSids = @()
    $unmatchedAccounts = @()

    if ($AccountsToKeep -and $existingProfiles) {
        # Resolve requested keep accounts to one or more SIDs so duplicates and roaming profiles remain intact
        $profileLookup = @{}

        foreach ($profileRecord in $existingProfiles) {
            if (-not $profileRecord.ProfileName) { continue }
            $nameLower = $profileRecord.ProfileName.ToLowerInvariant()
            if (-not $profileLookup.ContainsKey($nameLower)) {
                $profileLookup[$nameLower] = @()
            }

            if ($profileRecord.SID) {
                $profileLookup[$nameLower] += $profileRecord.SID
            }
        }

        foreach ($account in $AccountsToKeep) {
            $accountLower = $account.ToLowerInvariant()
            if ($profileLookup.ContainsKey($accountLower)) {
                $AccountsToKeepSids += $profileLookup[$accountLower]
            }
            else {
                $unmatchedAccounts += $account
            }
        }
    }
    elseif ($AccountsToKeep) {
        $unmatchedAccounts = @($AccountsToKeep)
    }

    if ($AccountsToKeepSids -and $AccountsToKeepSids.Count -gt 0) {
        # Log the resolved SID list for auditors to cross-reference with the final state
        $AccountsToKeepSids = $AccountsToKeepSids | Sort-Object -Unique
        Write-Log -Session $logSession -Message "Matching profile SIDs preserved:" -NoConsole
        foreach ($sid in $AccountsToKeepSids) {
            Write-Log -Session $logSession -Message (" - {0}" -f $sid) -NoConsole
        }
    }

    if ($unmatchedAccounts -and $unmatchedAccounts.Count -gt 0) {
        Write-Host "Warning: No matching profiles found for the following keep entries:" -ForegroundColor Yellow
        foreach ($missingAccount in $unmatchedAccounts) {
            Write-Host ("  {0}" -f $missingAccount) -ForegroundColor Yellow
            Write-Log -Session $logSession -Message ("Warning: No matching profile found for keep entry '{0}'." -f $missingAccount) -Severity 'Warning'
        }
    }

    if (-not (Confirm-Operation -ComputerName $ComputerName -AccountsToKeep $AccountsToKeep)) {
        # Gracefully exit if the operator decides to abort after reviewing the inputs
        Write-Host ""
        Write-Host "Operation cancelled. No profiles were removed from '$ComputerName'." -ForegroundColor Yellow
        Write-Log -Session $logSession -Message ("Operation cancelled. No profiles were removed from '{0}'." -f $ComputerName) -Severity 'Warning'
        return
    }

    Write-Host ""
    Write-Host "Removing user profiles from '$ComputerName'..." -ForegroundColor Cyan
    # Invoke the remote cleanup using the curated keep lists

    $removalResult = Remove-RemoteProfiles -ComputerName $ComputerName -AccountsToKeep $AccountsToKeep -AccountSidsToKeep $AccountsToKeepSids

    $removedProfiles = @()
    $skippedProfiles = @()

    if ($removalResult) {
        # Normalize the removal payload for easier reporting/logging downstream
        if ($removalResult.Removed) {
            $removedProfiles = @($removalResult.Removed)
        }
        if ($removalResult.Skipped) {
            $skippedProfiles = @($removalResult.Skipped)
        }
    }

    $removedProfileCount = ($removedProfiles | Measure-Object).Count
    $skippedProfileCount = ($skippedProfiles | Measure-Object).Count

    Write-Log -Session $logSession -Message ("Removed {0} profile(s) from '{1}'." -f $removedProfileCount, $ComputerName)
    Write-Log -Session $logSession -Message ("Skipped {0} profile(s) on '{1}'." -f $skippedProfileCount, $ComputerName)
    Write-RemovalResultsLog -RemovedProfiles $removedProfiles -SkippedProfiles $skippedProfiles -Session $logSession

    $remainingProfiles = Get-RemoteProfiles -ComputerName $ComputerName
    # Publish a second snapshot so operators can compare the delta post-removal
    Write-ProfileSection -Title "Profiles after removal" -Profiles $remainingProfiles -Session $logSession -NoConsole

    Write-Host ""
    if ($removedProfileCount -gt 0) {
        Write-Host ("Removed {0} profile(s) from '{1}':" -f $removedProfileCount, $ComputerName) -ForegroundColor Green
        $removedProfiles | Select-Object ProfileName, SID, LocalPath | Format-Table -AutoSize
    }
    else {
        Write-Host ("No profiles were removed from '{0}'. Either none matched the criteria or all specified profiles were preserved." -f $ComputerName) -ForegroundColor Green
    }

    if ($skippedProfileCount -gt 0) {
        Write-Host ""
        Write-Host ("Skipped {0} profile(s) on '{1}' (see log for details)." -f $skippedProfileCount, $ComputerName) -ForegroundColor Cyan
        $skippedProfiles | Select-Object ProfileName, Reason, SID, LocalPath, Loaded | Format-Table -AutoSize
    }

    Show-KeepList -AccountsToKeep $AccountsToKeep

    if ($remainingProfiles -and $remainingProfiles.Count -gt 0) {
        Write-Host "Profiles remaining on '$ComputerName':" -ForegroundColor Cyan
        $remainingProfiles | Sort-Object ProfileName | Format-Table ProfileName, SID, LocalPath, LastUseTime, Loaded -AutoSize
    }
    else {
        Write-Host "No user profiles remain on '$ComputerName'." -ForegroundColor Green
    }

    # Record a terminating event with timestamp to close the log session
    Write-Log -Session $logSession -Message ("Profile removal script completed for '{0}' at {1}." -f $ComputerName, (Get-Date))
    Write-Host ("Profile removal script completed for '{0}' at {1}." -f $ComputerName, (Get-Date)) -ForegroundColor Cyan
}
catch {
    Write-Error $_
    exit 1
}
