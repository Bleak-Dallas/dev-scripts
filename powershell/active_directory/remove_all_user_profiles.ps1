[CmdletBinding()]
param(
    [string]$ComputerName
)

<#
.SYNOPSIS
    Remotely remove every non-system user profile from a Windows workstation.

.DESCRIPTION
    1. Loads the shared utilities module (from local or network locations) and relaunches the script
       with elevation if required.
    2. Validates connectivity to the target computer, inventories existing user profiles, and warns the
       operator that all profiles are about to be removed.
    3. Requests explicit operator confirmation before continuing with the destructive action.
    4. Invokes a remote removal routine that deletes all eligible profiles, capturing structured
       results (removed vs. skipped) and writing before/after snapshots to the log.
    5. Writes a fully timestamped log to C:\Temp\logs using the shared logging helpers so administrators
       can review the run end-to-end.

.PARAMETER ComputerName
    Optional. The remote computer whose local profiles should be evaluated. Prompts if omitted or blank.

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
    PS> .\remove_all_user_profiles.ps1 -ComputerName "SLC-WS12345"

    Removes every non-system, non-loaded profile from SLC-WS12345, recording actions in C:\Temp\logs.
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

$elevationResult = Invoke-Elevation -BoundParameters $PSBoundParameters -ScriptPath $PSCommandPath
if (-not $elevationResult) {
    Write-Host 'Elevated session launched in a new window. Complete the workflow there and close the elevated window when finished.' -ForegroundColor Cyan
    return
}

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

    $header = "===== remove_all_user_profiles run {0} for {1} =====" -f (Get-Date), $ComputerName

    return New-LogSession -LogNamePrefix 'remove_all_user_profiles' -Header $header -Metadata $metadata
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

function Confirm-RemoveAll {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [object[]]$ProfilesToRemove = @()
    )

    $profileCount = ($ProfilesToRemove | Measure-Object).Count
    Write-Host ""
    Write-Host ("WARNING: This will permanently delete {0} profile(s) from '{1}'." -f $profileCount, $ComputerName) -ForegroundColor Yellow
    Write-Host "Profiles marked for removal include non-system accounts only. Loaded profiles and mandatory system accounts will be skipped automatically." -ForegroundColor DarkYellow

    while ($true) {
        Write-Host "Do you want to continue?" -ForegroundColor Cyan
        Write-Host "1: Yes, remove all profiles"
        Write-Host "2: No, cancel"
        $selection = Read-Host "Please make a selection"

        switch ($selection) {
            '1' { return $true }
            '2' { return $false }
            default {
                Write-Host "Invalid selection. Please choose 1 or 2." -ForegroundColor Yellow
            }
        }
    }
}
# endregion

# region Remote profile removal helper
function Remove-RemoteProfiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    $scriptBlock = {
        $removed = @()
        $skipped = @()

        $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
        Where-Object { -not $_.Special -and $_.LocalPath }

        foreach ($remoteProfile in $profiles) {
            $profileName = Split-Path $remoteProfile.LocalPath -Leaf
            if (-not $profileName) {
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

            try {
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
            catch {
                $skipped += [pscustomobject]@{
                    Action      = 'Skipped'
                    Reason      = ('RemovalFailed: {0}' -f $_.Exception.Message)
                    ProfileName = $profileName
                    SID         = $remoteProfile.SID
                    LocalPath   = $remoteProfile.LocalPath
                    Loaded      = $remoteProfile.Loaded
                }
            }
        }

        return [pscustomobject]@{
            Removed = $removed
            Skipped = $skipped
        }
    }

    try {
        return Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction Stop
    }
    catch {
        throw "Failed to remove profiles from '$ComputerName' using the current credentials. $_"
    }
}
# endregion

try {
    $ComputerName = Read-ComputerName -InitialName $ComputerName
    Test-ComputerConnectivity -ComputerName $ComputerName

    $logSession = Initialize-LogSession -ComputerName $ComputerName
    $logPath = $logSession.LogPath
    Write-Log -Session $logSession -Message ("Log file: {0}" -f $logPath) -Severity 'Info' -Color ([System.ConsoleColor]::Cyan)

    Write-Host ""
    Write-Host "Retrieving user profile list from '$ComputerName' using the current elevated credentials..." -ForegroundColor Cyan
    $existingProfiles = Get-RemoteProfiles -ComputerName $ComputerName
    $existingProfilesArray = @($existingProfiles)

    Write-ProfileSection -Title "Profiles before removal" -Profiles $existingProfilesArray -Session $logSession -NoConsole

    if ($existingProfilesArray.Count -eq 0) {
        Write-Host "No removable user profiles were found on '$ComputerName'." -ForegroundColor Green
        Write-Log -Session $logSession -Message ("No removable user profiles were found on '{0}'." -f $ComputerName) -Severity 'Info'
        Write-Log -Session $logSession -Message ("remove_all_user_profiles script completed for '{0}' at {1}." -f $ComputerName, (Get-Date))
        Write-Host ("remove_all_user_profiles script completed for '{0}' at {1}." -f $ComputerName, (Get-Date)) -ForegroundColor Cyan
        return
    }

    Write-Host "The following profiles were discovered on '$ComputerName':" -ForegroundColor Cyan
    $existingProfilesArray | Sort-Object ProfileName | Format-Table ProfileName, SID, LocalPath, LastUseTime, Loaded -AutoSize
    Write-Host ""
    Write-Host "All listed profiles (excluding system and currently loaded profiles) will be removed if you continue." -ForegroundColor Yellow

    if (-not (Confirm-RemoveAll -ComputerName $ComputerName -ProfilesToRemove $existingProfilesArray)) {
        Write-Host ""
        Write-Host "Operation cancelled. No profiles were removed from '$ComputerName'." -ForegroundColor Yellow
        Write-Log -Session $logSession -Message ("Operation cancelled. No profiles were removed from '{0}'." -f $ComputerName) -Severity 'Warning'
        return
    }

    Write-Host ""
    Write-Host "Removing user profiles from '$ComputerName'..." -ForegroundColor Cyan
    $removalResult = Remove-RemoteProfiles -ComputerName $ComputerName

    $removedProfiles = @()
    $skippedProfiles = @()

    if ($removalResult) {
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
    $remainingProfilesArray = @($remainingProfiles)
    Write-ProfileSection -Title "Profiles after removal" -Profiles $remainingProfilesArray -Session $logSession -NoConsole

    Write-Host ""
    if ($removedProfileCount -gt 0) {
        Write-Host ("Removed {0} profile(s) from '{1}':" -f $removedProfileCount, $ComputerName) -ForegroundColor Green
        $removedProfiles | Select-Object ProfileName, SID, LocalPath | Format-Table -AutoSize
    }
    else {
        Write-Host ("No profiles were removed from '{0}'. All eligible profiles may have been skipped." -f $ComputerName) -ForegroundColor Yellow
    }

    if ($skippedProfileCount -gt 0) {
        Write-Host ""
        Write-Host ("Skipped {0} profile(s) on '{1}' (see log for details)." -f $skippedProfileCount, $ComputerName) -ForegroundColor Cyan
        $skippedProfiles | Select-Object ProfileName, Reason, SID, LocalPath, Loaded | Format-Table -AutoSize
    }

    if ($remainingProfilesArray.Count -gt 0) {
        Write-Host ""
        Write-Host "Profiles remaining on '$ComputerName':" -ForegroundColor Cyan
        $remainingProfilesArray | Sort-Object ProfileName | Format-Table ProfileName, SID, LocalPath, LastUseTime, Loaded -AutoSize
    }
    else {
        Write-Host ""
        Write-Host "No user profiles remain on '$ComputerName'." -ForegroundColor Green
    }

    Write-Log -Session $logSession -Message ("remove_all_user_profiles script completed for '{0}' at {1}." -f $ComputerName, (Get-Date))
    Write-Host ("remove_all_user_profiles script completed for '{0}' at {1}." -f $ComputerName, (Get-Date)) -ForegroundColor Cyan
}
catch {
    Write-Error $_
    exit 1
}
