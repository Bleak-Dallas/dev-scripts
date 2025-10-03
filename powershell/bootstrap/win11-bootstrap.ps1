<# ============================================================================
 Windows 11 Bootstrap (Interactive) - Dallas
 Goal: Fast “day-0” setup (apps + a few settings). Store is BLOCKED, winget ALLOWED.
 Scope: Installs apps via winget, enables WSL platform + Ubuntu (if possible),
        applies Explorer tweaks, sets Git identity, adds optional local installers.
 Logging: C:\Temp\Logs  |  Safe to re-run (idempotent where practical).
 ============================================================================ #>

#Requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'

#region --- Session & Logging --------------------------------------------------
$LogRoot = 'C:\Temp\Logs'
if (-not (Test-Path $LogRoot)) { New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null }
$Transcript = Join-Path $LogRoot ("win11-bootstrap-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $Transcript -Force | Out-Null

function Info($m) { Write-Host "✅ $m" }
function Step($m) { Write-Host "— $m" -ForegroundColor Cyan }
function Skip($m) { Write-Host "➡️  $m" -ForegroundColor DarkGray }
function Warn2($m) { Write-Warning $m }
$rebootNeeded = $false
#endregion ---------------------------------------------------------------------

#region --- Helpers ------------------------------------------------------------
function Invoke-ExecutionPolicyProcess {
    Step "Set relaxed execution policy for this session"
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
    Info "Execution policy set for this session (Bypass)"
}

function Test-InstalledWingetId {
    param([Parameter(Mandatory)][string]$Id)
    $null = winget list --id $Id --exact --accept-source-agreements 2>$null
    return ($LASTEXITCODE -eq 0)
}

function Install-WingetId {
    param(
        [Parameter(Mandatory)][string]$Id
    )
    if (Test-InstalledWingetId -Id $Id) { Skip "Already installed: $Id"; return }
    Step "Installing: $Id"
    winget install --id $Id --exact --accept-package-agreements --accept-source-agreements --scope machine --silent
    if ($LASTEXITCODE -eq 0) { Info "Installed: $Id" } else { Warn2 "Install failed for $Id (exit $LASTEXITCODE)" }
}

function Enable-FeatureIfMissing {
    param([string]$Name)
    $state = (Get-WindowsOptionalFeature -Online -FeatureName $Name).State
    if ($state -ne 'Enabled') {
        Step "Enabling Windows feature: $Name"
        Enable-WindowsOptionalFeature -Online -FeatureName $Name -All -NoRestart | Out-Null
        Info "Enabled: $Name"
        $script:rebootNeeded = $true
    }
    else {
        Skip "Feature already enabled: $Name"
    }
}

function Install-FromInstaller {
    <#
      .SYNOPSIS
        Runs a local installer (EXE/MSI) silently when possible.

      .PARAMETER Path
        Full path to the installer.

      .PARAMETER SilentArgs
        Silent arguments; auto-detected defaults applied if omitted.

      .EXAMPLES
        Install-FromInstaller -Path 'D:\Installers\SecureCRT-9.6.3-x64.exe' -SilentArgs '/S'
        Install-FromInstaller -Path '\\srv\share\AD\rsat-cab\Setup.msi' -SilentArgs '/qn /norestart'
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$SilentArgs
    )
    if (-not (Test-Path $Path)) { Warn2 "Installer not found: $Path"; return }

    $ext = ([System.IO.Path]::GetExtension($Path)).ToLowerInvariant()
    if (-not $SilentArgs) {
        $SilentArgs = switch ($ext) {
            '.msi' { '/qn /norestart' }
            default { '/S' } # common for NSIS/InstallShield/Inno; adjust per vendor
        }
    }
    Step "Running local installer: $Path $SilentArgs"
    if ($ext -eq '.msi') {
        Start-Process msiexec.exe -ArgumentList "/i `"$Path`" $SilentArgs" -Wait -NoNewWindow
    }
    else {
        Start-Process $Path -ArgumentList $SilentArgs -Wait -NoNewWindow
    }
    Info "Installer completed: $Path"
}

#endregion ---------------------------------------------------------------------

Ensure-ExecutionPolicyProcess

#region --- Winget availability ------------------------------------------------
Step "Check winget availability (Store blocked is OK if winget is present)"
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Warn2 "winget is not available. Install 'App Installer' via offline package or enable via policy, then re-run."
    Stop-Transcript | Out-Null
    exit 2
}
try {
    winget source update --name winget  | Out-Null
    winget source update --name msstore | Out-Null   # harmless if Store is blocked
    Info "winget sources refreshed"
}
catch {
    Skip "winget source refresh had warnings (Store likely blocked)."
}
#endregion ---------------------------------------------------------------------

#region --- Install catalog from JSON -----------------------------------------
Step "Install core apps from JSON catalog"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CatalogPath = Join-Path $ScriptDir 'win11-apps.json'
if (-not (Test-Path $CatalogPath)) {
    Warn2 "Catalog missing at $CatalogPath. Place win11-apps.json next to the script."
}
else {
    try {
        $apps = (Get-Content -Raw -Path $CatalogPath | ConvertFrom-Json).packages
        foreach ($pkg in $apps) { Install-WingetId -Id $pkg.id }
    }
    catch {
        Warn2 "Failed to parse/install from catalog: $($_.Exception.Message)"
    }
}
#endregion ---------------------------------------------------------------------

#region --- Git identity -------------------------------------------------------
Step "Configure Git identity (global)"
if (Get-Command git -ErrorAction SilentlyContinue) {
    git config --global user.name  "Dallas Bleak"
    git config --global user.email "dbleak42@gmail.com"
    Info "Git identity set"
}
else {
    Skip "Git not installed yet (skipping identity); re-run after Git is installed"
}
#endregion ---------------------------------------------------------------------

#region --- WSL: Platform + Ubuntu --------------------------------------------
# Enable WSL platform features; attempt Ubuntu install if possible.
Step "Enable WSL platform (no Store dependency)"
Enable-FeatureIfMissing -Name 'Microsoft-Windows-Subsystem-Linux'
Enable-FeatureIfMissing -Name 'VirtualMachinePlatform'

# Default WSL version = 2
try {
    wsl --set-default-version 2 2>$null
    Info "WSL default version set to 2"
}
catch {
    Skip "WSL default may already be 2 or requires reboot first"
}

# Try to install Ubuntu (note: may require Store; if blocked, this can fail).
$installed = & wsl -l -q 2>$null
if ($LASTEXITCODE -eq 0 -and ($installed -notcontains 'Ubuntu')) {
    Step "Attempting to install Ubuntu (may fail if Store content is blocked)"
    try {
        wsl --install -d Ubuntu
        Info "Ubuntu install requested"
        $rebootNeeded = $true
    }
    catch {
        Warn2 "Ubuntu install could not be triggered (Store/content may be blocked). Use offline .appx if needed."
    }
}
else {
    Skip "Ubuntu already installed or WSL listing unavailable"
}
#endregion ---------------------------------------------------------------------

#region --- Explorer tweaks ----------------------------------------------------
Step "Apply File Explorer visibility tweaks"
$adv = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
New-Item -Path $adv -Force | Out-Null
New-ItemProperty -Path $adv -Name HideFileExt -Value 0 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $adv -Name Hidden      -Value 1 -PropertyType DWord -Force | Out-Null
Info "Extensions visible; hidden files visible"
#endregion ---------------------------------------------------------------------

#region --- File association helpers (safe mode) -------------------------------
<#
  Windows 11 protects per-user "UserChoice" associations with a hash.
  Instead of forcing defaults (fragile), we add context-menu verbs that are reliable.
  After running, right-click a .ps1 → "Run with PowerShell 7"
                     a .md  → "Open with VS Code"
#>

Step "Register context menu verbs for PowerShell 7 and VS Code (non-destructive)"
# PowerShell 7 context verb for .ps1
$ps7 = "${env:ProgramFiles}\PowerShell\7\pwsh.exe"
if (Test-Path $ps7) {
    $ps1Key = 'HKCU:\Software\Classes\Microsoft.PowerShellScript.1\shell\RunWithPS7\command'
    New-Item -Path $ps1Key -Force | Out-Null
    Set-ItemProperty -Path (Split-Path $ps1Key) -Name '(Default)' -Value 'Run with PowerShell 7' -Force
    Set-ItemProperty -Path $ps1Key -Name '(Default)' -Value "`"$ps7`" -NoExit -File `"%1`"" -Force
    Info "Added 'Run with PowerShell 7' context verb for .ps1"
}
else {
    Skip "PowerShell 7 not found at $ps7 — install if desired and re-run this section"
}

# VS Code context verb for .md
$code = "${env:ProgramFiles}\Microsoft VS Code\Code.exe"
if (Test-Path $code) {
    $mdKey = 'HKCU:\Software\Classes\Markdown.File\shell\OpenWithVSCode\command'
    New-Item -Path $mdKey -Force | Out-Null
    Set-ItemProperty -Path (Split-Path $mdKey) -Name '(Default)' -Value 'Open with VS Code' -Force
    Set-ItemProperty -Path $mdKey -Name '(Default)' -Value "`"$code`" `"%1`"" -Force
    Info "Added 'Open with VS Code' context verb for .md"
}
else {
    Skip "VS Code not found at $code — install if desired and re-run this section"
}
#endregion ---------------------------------------------------------------------

#region --- Local installers (SecureCRT, AD RSAT offline) ----------------------
# Set your local paths here (network share or local disk); script will skip if blank/missing.
$SecureCRTInstaller = ''   # e.g., '\\fileserver\share\SecureCRT-9.6.3-x64.exe'
$SecureCRTSilent = '/S' # adjust per vendor

$RsatFoDPath = ''          # Folder that contains FoD .cab/.mum for RSAT AD tools (optional / offline)
# If you prefer online: use Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0

if ($SecureCRTInstaller) {
    Install-FromInstaller -Path $SecureCRTInstaller -SilentArgs $SecureCRTSilent
}
else {
    Skip "SecureCRT installer path not set; update \$SecureCRTInstaller"
}

if ($RsatFoDPath) {
    Step "Attempting RSAT AD Tools offline install from $RsatFoDPath"
    try {
        # Generic FoD add (expects proper .cab layout). Adjust package name if needed.
        dism /online /add-capability /capabilityname:Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 /source:$RsatFoDPath /limitaccess
        Info "RSAT AD Tools requested via DISM (offline source)"
    }
    catch {
        Warn2 "Offline RSAT install failed; verify FoD source folder and package name."
    }
}
else {
    Skip "RSAT offline source not set; skipping AD tools. (Online alt: Add-WindowsCapability -Online ...)"
}
#endregion ---------------------------------------------------------------------

#region --- Reboot prompt & Done ----------------------------------------------
if ($rebootNeeded) {
    Warn2 "A reboot is recommended to finalize platform changes (WSL/VM Platform)."
    Write-Host "Run: shutdown /r /t 0  (or reboot later)"
}

Info "Bootstrap complete. Log: $Transcript"
Stop-Transcript | Out-Null
#endregion ---------------------------------------------------------------------
