[CmdletBinding()]
param(
    [string]$ComputerName,
    [string[]]$AccountsToKeep = @()
)

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

function Initialize-Logger {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    $logRoot = 'C:\Temp\logs'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    try {
        if (-not (Test-Path -LiteralPath $logRoot)) {
            New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
        }
    }
    catch {
        throw "Failed to ensure log directory '$logRoot'. $_"
    }

    $logFilePath = Join-Path $logRoot ("remove_user_profile_{0}_{1}.log" -f $ComputerName, $timestamp)
    $header = "===== remove_user_profile run {0} for {1} =====" -f (Get-Date), $ComputerName
    Set-Content -Path $logFilePath -Value $header

    return $logFilePath
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        [System.ConsoleColor]$Color = [System.ConsoleColor]::Gray,
        [switch]$NoConsole
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] {1}" -f $timestamp, $Message
    Add-Content -Path $LogPath -Value $line

    if (-not $NoConsole) {
        Write-Host $Message -ForegroundColor $Color
    }
}

function Write-ProfileLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [object[]]$Profiles,
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    Write-Log -Message $Title -LogPath $LogPath -NoConsole

    if ($Profiles -and $Profiles.Count -gt 0) {
        foreach ($profileRecord in $Profiles) {
            $lastUseStamp = if ($profileRecord.LastUseTime) { $profileRecord.LastUseTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'n/a' }
            $entry = " - {0} (SID: {1}; Path: {2}; Loaded: {3}; LastUse: {4})" -f $profileRecord.ProfileName, $profileRecord.SID, $profileRecord.LocalPath, $profileRecord.Loaded, $lastUseStamp
            Write-Log -Message $entry -LogPath $LogPath -NoConsole
        }
    }
    else {
        Write-Log -Message " - (none)" -LogPath $LogPath -NoConsole
    }
}

function Write-AccountsToKeep {
    param(
        [string[]]$Accounts,
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    Write-Log -Message "Accounts preserved:" -LogPath $LogPath -NoConsole

    if ($Accounts -and $Accounts.Count -gt 0) {
        foreach ($account in $Accounts) {
            Write-Log -Message (" - {0}" -f $account) -LogPath $LogPath -NoConsole
        }
    }
    else {
        Write-Log -Message " - (none)" -LogPath $LogPath -NoConsole
    }
}

function Write-RemovalResults {
    param(
        [object[]]$RemovedProfiles = @(),
        [object[]]$SkippedProfiles = @(),
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    $removedCount = ($RemovedProfiles | Measure-Object).Count
    Write-Log -Message ("Removed {0} profile(s)." -f $removedCount) -LogPath $LogPath -NoConsole

    if ($removedCount -gt 0) {
        foreach ($record in $RemovedProfiles) {
            $entry = " - Removed {0} (SID: {1}; Path: {2}; Loaded: {3}; Reason: {4})" -f $record.ProfileName, $record.SID, $record.LocalPath, $record.Loaded, $record.Reason
            Write-Log -Message $entry -LogPath $LogPath -NoConsole
        }
    }
    else {
        Write-Log -Message " - (none removed)" -LogPath $LogPath -NoConsole
    }

    $skippedCount = ($SkippedProfiles | Measure-Object).Count
    Write-Log -Message ("Skipped {0} profile(s)." -f $skippedCount) -LogPath $LogPath -NoConsole

    if ($skippedCount -gt 0) {
        foreach ($record in $SkippedProfiles) {
            $entry = " - Skipped {0} (Reason: {1}; SID: {2}; Path: {3}; Loaded: {4})" -f $record.ProfileName, $record.Reason, $record.SID, $record.LocalPath, $record.Loaded
            Write-Log -Message $entry -LogPath $LogPath -NoConsole
        }
    }
    else {
        Write-Log -Message " - (none skipped)" -LogPath $LogPath -NoConsole
    }
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
    $ComputerName = Read-ComputerName -InitialName $ComputerName
    Test-ComputerConnectivity -ComputerName $ComputerName

    $logFilePath = Initialize-Logger -ComputerName $ComputerName
    Write-Log -Message ("Log file: {0}" -f $logFilePath) -LogPath $logFilePath -Color ([System.ConsoleColor]::Cyan)

    Write-Host ""
    Write-Host "Retrieving user profile list from '$ComputerName' using the current elevated credentials..." -ForegroundColor Cyan
    $existingProfiles = Get-RemoteProfiles -ComputerName $ComputerName
    Write-ProfileLog -Title "Profiles before removal" -Profiles $existingProfiles -LogPath $logFilePath

    if ($existingProfiles -and $existingProfiles.Count -gt 0) {
        Write-Host "Current profiles found on '$ComputerName':" -ForegroundColor Cyan
        $existingProfiles | Sort-Object ProfileName | Format-Table ProfileName, SID, LocalPath, LastUseTime, Loaded -AutoSize
    }
    else {
        Write-Host "No removable user profiles were found on '$ComputerName'." -ForegroundColor Yellow
    }

    Write-Host ""

    if (-not $PSBoundParameters.ContainsKey('AccountsToKeep') -or $AccountsToKeep.Count -eq 0) {
        $AccountsToKeep = Get-AccountsToKeep
    }

    $AccountsToKeep = ConvertTo-NormalizedAccounts -Accounts $AccountsToKeep
    Write-AccountsToKeep -Accounts $AccountsToKeep -LogPath $logFilePath

    $AccountsToKeepSids = @()
    $unmatchedAccounts = @()

    if ($AccountsToKeep -and $existingProfiles) {
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
        $AccountsToKeepSids = $AccountsToKeepSids | Sort-Object -Unique
        Write-Log -Message "Matching profile SIDs preserved:" -LogPath $logFilePath -NoConsole
        foreach ($sid in $AccountsToKeepSids) {
            Write-Log -Message (" - {0}" -f $sid) -LogPath $logFilePath -NoConsole
        }
    }

    if ($unmatchedAccounts -and $unmatchedAccounts.Count -gt 0) {
        Write-Host "Warning: No matching profiles found for the following keep entries:" -ForegroundColor Yellow
        foreach ($missingAccount in $unmatchedAccounts) {
            Write-Host ("  {0}" -f $missingAccount) -ForegroundColor Yellow
            Write-Log -Message ("Warning: No matching profile found for keep entry '{0}'." -f $missingAccount) -LogPath $logFilePath
        }
    }

    if (-not (Confirm-Operation -ComputerName $ComputerName -AccountsToKeep $AccountsToKeep)) {
        Write-Host ""
        Write-Host "Operation cancelled. No profiles were removed from '$ComputerName'." -ForegroundColor Yellow
        Write-Log -Message ("Operation cancelled. No profiles were removed from '{0}'." -f $ComputerName) -LogPath $logFilePath
        return
    }

    Write-Host ""
    Write-Host "Removing user profiles from '$ComputerName'..." -ForegroundColor Cyan

    $removalResult = Remove-RemoteProfiles -ComputerName $ComputerName -AccountsToKeep $AccountsToKeep -AccountSidsToKeep $AccountsToKeepSids

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

    Write-Log -Message ("Removed {0} profile(s) from '{1}'." -f $removedProfileCount, $ComputerName) -LogPath $logFilePath
    Write-Log -Message ("Skipped {0} profile(s) on '{1}'." -f $skippedProfileCount, $ComputerName) -LogPath $logFilePath
    Write-RemovalResults -RemovedProfiles $removedProfiles -SkippedProfiles $skippedProfiles -LogPath $logFilePath

    $remainingProfiles = Get-RemoteProfiles -ComputerName $ComputerName
    Write-ProfileLog -Title "Profiles after removal" -Profiles $remainingProfiles -LogPath $logFilePath

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

    Write-Log -Message ("Profile removal script completed for '{0}' at {1}." -f $ComputerName, (Get-Date)) -LogPath $logFilePath
    Write-Host ("Profile removal script completed for '{0}' at {1}." -f $ComputerName, (Get-Date)) -ForegroundColor Cyan
}
catch {
    Write-Error $_
    exit 1
}
