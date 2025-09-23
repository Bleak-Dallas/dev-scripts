#requires -Version 5.1
<#
.SYNOPSIS
    Remotely copies and installs software on Windows computers (single or multiple).

.DESCRIPTION
    This script helps a technician choose one or more software packages from a shared network
    folder and install them on a single computer or on a list of computers. It uses PowerShell
    Remoting (WinRM) to:
      - Open a remote session to the target computer
      - Create C:\Temp on the target if it does not exist
      - Copy the selected software folder(s) into C:\Temp\<App>
      - Run the app's Install.cmd with a quiet/silent switch when applicable

    It logs computers that are unreachable to a CSV file (real CSV, not plain text),
    and it avoids turning off services (WinRM) that were already running before the script started.

.PARAMETER Computer
    (Used by Install-Applications) The target computer name (DNS or NetBIOS).

.PARAMETER Applications
    (Used by Install-Applications) One or more application folder names that exist in $SoftwareDir.

.NOTES
    ┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
    │ ORIGIN STORY                                                                                │ 
    ├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
    │   DATE        : 2025-08-19                                                                  │
    │   AUTHOR      : Dallas Bleak                                                                │
    │   VERSION     : 1.4.1                                                                       │
    │   Run As      : Elevated PowerShell (Run as Administrator) recommended.                     │
    └─────────────────────────────────────────────────────────────────────────────────────────────┘ 

.LINK
    Get-Help about_Remote

.LIMITATIONS
    - The remote installer must be "Install.cmd" inside each app folder.
    - Your current user must have admin rights on the remote machines.
    - PowerShell Remoting (WinRM) must be allowed between your machine and the targets.
#>

# ======================
# GLOBAL CONSTANTS
# ======================

# Root share where software and lists live
$ShareRoot = '\\va.gov\cn\Salt Lake City\VHASLC\TechDrive'
# Folder containing application subfolders (each with an Install.cmd)
$SoftwareDir = Join-Path $ShareRoot 'Software'
# Folder where per-tech text files of computer names are kept
$ListDir = Join-Path $ShareRoot 'Scripts\fileNames'
# CSV path for logging offline/unreachable computers
$LogPath = 'C:\Temp\Logs\laptopOffline.csv'

# Ensure the log directory exists (no error if already present)
$null = New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force -ErrorAction SilentlyContinue

# Remember the Cisco device name once per run to avoid repeatedly prompting
$script:DeviceName = $null

<#
.SYNOPSIS
    Append an OFFLINE entry to the CSV log for a computer.

.DESCRIPTION
    Creates (or appends to) a real CSV file at $LogPath with columns:
    When, Computer, Status. Writes a red host message as well.

.PARAMETER Computer
    The computer name that was unreachable by WSMan/WinRM.

.EXAMPLE
    Write-OfflineLog -Computer 'PC-123'

.OUTPUTS
    None (writes to CSV and host).
#>
function Write-OfflineLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Computer
    )

    # Construct a row object with timestamp and status
    $row = [pscustomobject]@{
        When     = (Get-Date)
        Computer = $Computer
        Status   = 'OFFLINE'
    }

    # Append to CSV (creates if missing)
    $row | Export-Csv -Path $LogPath -NoTypeInformation -Append
    Write-Host "$Computer OFFLINE (logged to $LogPath)" -ForegroundColor Red
}

<#
.SYNOPSIS
    Copy and install selected applications on a single remote computer.

.DESCRIPTION
    Checks WSMan connectivity, establishes a PSSession, ensures C:\Temp exists,
    copies each selected application folder to C:\Temp\<App>, and runs
    Install.cmd with arguments when needed.

.PARAMETER Computer
    Target computer name (DNS or NetBIOS).

.PARAMETER Applications
    One or more application folder names that exist under $SoftwareDir.

.PARAMETER WhatIf
    Shows what would happen if the command runs. No actions are performed.

.PARAMETER Confirm
    Prompts for confirmation before executing any action that changes the system.

.EXAMPLE
    Install-Applications -Computer 'PC-123' -Applications '7-Zip v24.07','CiscoIPCommunicator' -Verbose

.NOTES
    Requires WinRM connectivity and admin rights on the remote machine.
#>
function Install-Applications {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)][string]   $Computer,
        [Parameter(Mandatory = $true)][string[]] $Applications
    )

    # Early exit if no apps were provided
    if (-not $Applications -or $Applications.Count -eq 0) {
        Write-Verbose "No applications were selected; nothing to do."
        return
    }

    Write-Host "===== $Computer =====" -ForegroundColor Cyan

    # STEP 1: Confirm WSMan/WinRM endpoint is answering (authoritative reachability)
    $wsmanOk = $false
    for ($i = 1; $i -le 2 -and -not $wsmanOk; $i++) {
        try {
            # Test-WSMan throws if unreachable; treat as authoritative
            Test-WSMan -ComputerName $Computer -ErrorAction Stop | Out-Null
            $wsmanOk = $true
        }
        catch {
            Write-Verbose "WSMan not responding on $Computer (attempt $i). $_"
            Start-Sleep -Seconds 3
        }
    }

    # If still no WSMan, log offline and move on
    if (-not $wsmanOk) {
        Write-OfflineLog -Computer $Computer
        return
    }

    # STEP 2: Open a PSSession and do all remote work through it
    $session = $null
    try {
        Write-Verbose "Opening a remote PowerShell session to $Computer…"
        $session = New-PSSession -ComputerName $Computer -ErrorAction Stop

        # Ensure C:\Temp exists on target
        Invoke-Command -Session $session -ScriptBlock {
            if (-not (Test-Path -LiteralPath 'C:\Temp')) {
                New-Item -Path 'C:\Temp' -ItemType Directory | Out-Null
            }
        }

        # Process each selected app
        foreach ($app in $Applications) {
            if (-not $app) { continue }

            # Construct target path C:\Temp\<App>
            $remoteAppPath = "C:\Temp\$app"

            # Create destination folder (safe with -Force)
            Invoke-Command -Session $session -ScriptBlock {
                param($path)
                New-Item -Path $path -ItemType Directory -Force | Out-Null
            } -ArgumentList $remoteAppPath

            # Source path \\share\Software\<App>
            $srcPath = Join-Path $SoftwareDir $app

            # Visual progress
            Write-Host  "Copying [$app] to $Computer…" -ForegroundColor Yellow
            Write-Verbose "Copy-Item -ToSession from '$srcPath\*' to '$remoteAppPath'"

            # Copy app files over the remoting session (fast and firewall-friendly)
            Copy-Item -Path (Join-Path $srcPath '*') `
                -Destination $remoteAppPath `
                -ToSession $session `
                -Recurse -Force -ErrorAction Stop

            # Build argument string for Install.cmd (defaults to quiet)
            $argString = '--quiet'

            # Special case: Cisco IP Communicator needs device name; prompt once per run
            if ($app -ieq 'CiscoIPCommunicator') {
                if (-not $script:DeviceName) {
                    $script:DeviceName = Read-Host -Prompt 'Device Name of the Softphone'
                }
                # Combine device name with quiet switch safely
                $parts = @($script:DeviceName, '--quiet') | Where-Object { $_ -and $_.ToString().Trim() -ne '' }
                $argString = ($parts -join ' ')
            }

            # Per-app whitelist requiring no args at all (exact folder names)
            $noArgApps = @(
                '7-Zip v24.07'
            )
            if ($app -in $noArgApps) {
                $argString = ''
            }

            # Normalize to a string (never let it be $null)
            if ($null -eq $argString) { $argString = '' }
            $argString = $argString.ToString()

            # Verify Install.cmd exists in the remote app folder
            $installCmd = Join-Path $remoteAppPath 'Install.cmd'
            $exists = Invoke-Command -Session $session -ScriptBlock {
                param($path) Test-Path -LiteralPath $path
            } -ArgumentList $installCmd

            if (-not $exists) {
                Write-Warning "[$Computer] $app -> 'Install.cmd' not found at $installCmd. Skipping."
                continue
            }

            Write-Host "Installing [$app] on $Computer…" -ForegroundColor Yellow
            if ([string]::IsNullOrWhiteSpace($argString)) {
                Write-Verbose "Start-Process '$installCmd' (no args) -Wait"
            }
            else {
                Write-Verbose "Start-Process '$installCmd' -ArgumentList '$argString' -Wait"
            }

            # Invoke remotely. Important: omit -ArgumentList entirely when empty to avoid binder issues
            Invoke-Command -Session $session -ArgumentList $installCmd, $argString -ScriptBlock {
                param($cmd, $argStringInner)

                # Use Start-Process to run Install.cmd and wait for completion
                if ([string]::IsNullOrWhiteSpace($argStringInner)) {
                    Start-Process -FilePath $cmd -Wait
                }
                else {
                    Start-Process -FilePath $cmd -ArgumentList $argStringInner -Wait
                }

                # Lightweight throttle to allow any short-lived post-install activities
                Write-Host "Install completed: $cmd" -ForegroundColor Green
                Start-Sleep -Seconds 2
            }
        }
    }
    catch {
        # Catch any failures during copy/install/session creation
        Write-Error "Failed installing on $Computer. $_"
    }
    finally {
        # Always clean up the session if we opened one
        if ($session) {
            Write-Verbose "Closing remote session for $Computer…"
            Remove-PSSession $session
        }
    }
}

<#
.SYNOPSIS
    Interactive flow to install on a single computer.

.DESCRIPTION
    Prompts for a single target computer name, then opens a GUI picker (Out-GridView)
    to select one or more applications discovered in $SoftwareDir. Invokes Install-Applications
    with the selected apps.

.EXAMPLE
    Install-Single

.NOTES
    Requires Out-GridView (PowerShell ISE or Windows PowerShell with GUI components).
#>
function Install-Single {
    [CmdletBinding()]
    param()

    # Prompt for one computer name
    $computer = Read-Host "Enter device name (hostname) for install"
    if (-not $computer) { return }

    # Discover app folders and present a picker UI
    $applications = Get-ChildItem -Path $SoftwareDir -Directory `
    | Select-Object -ExpandProperty Name `
    | Sort-Object `
    | Out-GridView -PassThru -Title 'Select Software for remote install'

    # If user cancels, return to main choice menu
    if (-not $applications) {
        Write-Verbose "User canceled app selection."
        Get-Choice
        return
    }

    # Perform the install with current -Verbose preference honored
    Install-Applications -Computer $computer -Applications $applications -Verbose:$VerbosePreference

    # Go back to main menu after completion
    Get-Choice
}

<#
.SYNOPSIS
    Interactive flow to install on multiple computers from a per-tech list.

.DESCRIPTION
    Determines the technician's username, ensures a per-user text file exists in $ListDir,
    opens it in Notepad for editing (one computer name per line), then lets the user select
    applications via Out-GridView and installs to each listed computer in sequence.

.EXAMPLE
    Install-Multiple

.NOTES
    If launching Notepad against a UNC path fails, the function falls back to a local temp
    copy and syncs it back to the share.
#>
function Install-Multiple {
    [CmdletBinding()]
    param()

    # Resolve the real technician username, preferring the active console user if available
    $techUser = $env:UserName
    try {
        $activeUser = (Get-CimInstance Win32_ComputerSystem).UserName
        if ($activeUser -and $activeUser.Contains('\')) {
            $candidate = $activeUser.Split('\')[-1]
            if ($candidate) { $techUser = $candidate }
        }
    }
    catch {
        # Ignore errors and retain $env:UserName
    }

    # Ensure list directory exists
    $null = New-Item -ItemType Directory -Path $ListDir -Force -ErrorAction SilentlyContinue

    # Build per-user list file, e.g. \\share\Scripts\fileNames\jdoe.txt
    $file = Join-Path $ListDir "$techUser.txt"

    # Create the file if missing
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        try {
            $null = New-Item -ItemType File -Path $file -Force -ErrorAction Stop
        }
        catch {
            throw "Failed to create list file at '$file'. $_"
        }
    }

    # Try to open the file in Notepad directly. If UNC launch fails, copy to local temp.
    $opened = $false
    try {
        Start-Process -FilePath "notepad.exe" -ArgumentList "`"$file`""
        $opened = $true
    }
    catch {
        Write-Warning "Could not open '$file' with Notepad directly. Will try a local temp copy."
    }

    if (-not $opened) {
        try {
            # Local temp fallback file path (per user)
            $localTemp = Join-Path $env:TEMP ("computers_{0}.txt" -f $techUser)
            # Copy content to local temp and open for edit
            Copy-Item -Path $file -Destination $localTemp -Force
            Start-Process -FilePath "notepad.exe" -ArgumentList "`"$localTemp`""
            # Pause to let the user finish edits; then sync back to the share
            Read-Host -Prompt "Edit the local temp list ($localTemp), save, then press ENTER to sync back"
            Copy-Item -Path $localTemp -Destination $file -Force
        }
        catch {
            # Final fallback: ask user to edit manually and press enter
            Write-Warning "Could not open or sync a local temp copy for '$file'. Please edit it manually, then press ENTER to continue."
            Read-Host | Out-Null
        }
    }
    else {
        # Normal path: wait for user to complete edits in Notepad
        Read-Host -Prompt "Add one computer name per line, save, then press ENTER to continue"
    }

    # Load names AFTER editing; ignore blank lines/whitespace-only
    $computerNames = Get-Content -Path $file | Where-Object { $_ -and $_.Trim() }
    if (-not $computerNames) {
        Write-Warning "No computer names found in $file."
        Get-Choice
        return
    }

    # App selection (same picker UI as single)
    $applications = Get-ChildItem -Path $SoftwareDir -Directory `
    | Select-Object -ExpandProperty Name `
    | Sort-Object `
    | Out-GridView -PassThru -Title 'Select Software for remote install'

    if (-not $applications) {
        Write-Verbose "User canceled app selection."
        Get-Choice
        return
    }

    # Execute for each computer in the list
    foreach ($computer in $computerNames) {
        Install-Applications -Computer $computer -Applications $applications -Verbose:$VerbosePreference
    }

    # Return to menu
    Get-Choice
}

<#
.SYNOPSIS
    Shows a menu to choose single or multiple-computer installs.

.DESCRIPTION
    Presents a simple console choice dialog with two options:
    Single (prompt for one host) or Multiple (use per-user list).
    After the chosen flow completes, the menu is shown again.

.EXAMPLE
    Get-Choice
#>
function Get-Choice {
    [CmdletBinding()]
    param()

    # Build two menu options: Single and Multiple
    $options = @(
        (New-Object System.Management.Automation.Host.ChoiceDescription '&Single', 'Install on a single computer'),
        (New-Object System.Management.Automation.Host.ChoiceDescription '&Multiple', 'Install on multiple computers')
    )

    # Present the prompt and get selection (default 0 => Single)
    $selection = $Host.UI.PromptForChoice(
        'Install',
        'Choose target scope:',
        $options,
        0
    )

    # Route to the appropriate function
    if ($selection -eq 0) {
        Install-Single
    }
    else {
        Install-Multiple
    }
}

# ======================
# ENTRY POINT
# ======================

# Start the menu-driven flow
Get-Choice