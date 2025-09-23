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

$RefactorPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'installSoftware.refactored.ps1'
if (-not (Test-Path $RefactorPath)) {
    Write-Host "Refactored script not found at $RefactorPath" -ForegroundColor Red
    return
}

function Select-Applications {
    param(
        [switch]$AllowNone
    )
    $apps = Get-ChildItem -Path (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) '..\\Software') -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object
    if (-not $apps) { Write-Host 'No applications found in Software folder' -ForegroundColor Yellow; return @() }
    $selected = $apps | Out-GridView -Title 'Select Software' -PassThru
    if (-not $selected -and -not $AllowNone) { Write-Host 'No selection made; aborting' -ForegroundColor Yellow; return @() }
    return $selected
}

function Read-ComputersFromFile {
    $file = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'fileNames\\' + $env:UserName + '.txt'
    if (-not (Test-Path -Path $file)) {
        New-Item -ItemType File -Path $file -Force | Out-Null
        Write-Host "Created $file - please add computer names (one per line) and press Enter to continue" -ForegroundColor Yellow
        Invoke-Item $file
        Read-Host -Prompt 'Press Enter when ready to continue'
    }
    else {
        Invoke-Item $file
        Read-Host -Prompt 'Press Enter when ready to continue'
    }
    return Get-Content -Path $file
}

function Invoke-Refactor {
    param(
        [string[]]$ComputerNames,
        [string[]]$Applications,
        [switch]$DryRun
    )

    $splat = @{ }
    if ($ComputerNames) { $splat['ComputerNames'] = $ComputerNames }
    if ($Applications) { $splat['Applications'] = $Applications }
    if ($DryRun) { $splat['NoPrompt'] = $true }

    # Ask for credential if needed
    $cred = $null
    $useCred = Read-Host -Prompt 'Provide credentials? (Y/N)'
    if ($useCred -match '^[Yy]') { $cred = Get-Credential; $splat['Credential'] = $cred }

    # Call refactored script via splatting
    & pwsh -NoProfile -ExecutionPolicy Bypass -File $RefactorPath @splat
}

function Show-Menu {
    Write-Host '1) Install on a single computer'
    Write-Host '2) Install on multiple computers (from file)'
    Write-Host '3) Dry-run (show what will run)'
    Write-Host 'Q) Quit'
}

while ($true) {
    Show-Menu
    $choice = Read-Host -Prompt 'Choice'
    switch ($choice.ToUpper()) {
        '1' {
            $computer = Read-Host -Prompt 'Enter device for install'
            $apps = Select-Applications
            if (-not $apps) { break }
            Invoke-Refactor -ComputerNames $computer -Applications $apps
        }
        '2' {
            $computers = Read-ComputersFromFile
            $apps = Select-Applications
            if (-not $apps) { break }
            Invoke-Refactor -ComputerNames $computers -Applications $apps
        }
        '3' {
            $computers = Read-ComputersFromFile
            $apps = Select-Applications
            if (-not $apps) { break }
            Invoke-Refactor -ComputerNames $computers -Applications $apps -DryRun
        }
        'Q' { break }
        default { Write-Host 'Invalid choice' -ForegroundColor Red }
    }
}
