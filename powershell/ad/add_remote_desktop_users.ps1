<#
.SYNOPSIS
    Adds a user to the Remote Desktop Users group on a remote computer.

.DESCRIPTION
    This script combines admin privilege checking with remote desktop user management.
    It will automatically request elevation if not running as administrator, then
    connect to a remote computer to add a user to the Remote Desktop Users group.

.PARAMETER ComputerName
    The name of the remote computer to connect to.

.NOTES
    ┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
    │ ORIGIN STORY                                                                                │ 
    ├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
    │   DATE        : 2025-08-19                                                                  │
    │   AUTHOR      : Dallas Bleak (Dallas.Bleak@va.gov)                                          │
    │   VERSION     : 1.0                                                                        │
    │   Run As      : Elevated PowerShell (Run as Administrator) recommended.                     │
    └─────────────────────────────────────────────────────────────────────────────────────────────┘ 
#>

param(
    [string]$ComputerName
)

# Function to check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to test WinRM connectivity
function Test-WinRMConnectivity {
    param([string]$ComputerName)
    
    Write-Host "Testing WinRM connectivity to $ComputerName..." -ForegroundColor Yellow
    
    try {
        $result = Test-WSMan -ComputerName $ComputerName -ErrorAction Stop
        Write-Host "✓ WinRM is enabled and accessible on $ComputerName" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ WinRM connectivity failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to enable WinRM on remote computer
function Enable-RemoteWinRM {
    param([string]$ComputerName)
    
    Write-Host "`nAttempting to enable WinRM on $ComputerName..." -ForegroundColor Yellow
    
    # Method 1: Try using WMI to execute Enable-PSRemoting
    try {
        Write-Host "Trying WMI method to enable WinRM..." -ForegroundColor Cyan
        
        $process = Invoke-WmiMethod -ComputerName $ComputerName -Class Win32_Process -Name Create -ArgumentList "powershell.exe -Command `"Enable-PSRemoting -Force -SkipNetworkProfileCheck`""
        
        if ($process.ReturnValue -eq 0) {
            Write-Host "WMI command executed successfully. Waiting for WinRM to initialize..." -ForegroundColor Green
            Start-Sleep -Seconds 10
            
            # Test if WinRM is now working
            if (Test-WinRMConnectivity -ComputerName $ComputerName) {
                return $true
            }
        }
    }
    catch {
        Write-Host "WMI method failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Method 2: Try using PsExec-style approach (if available)
    Write-Host "WMI method unsuccessful. Manual intervention may be required." -ForegroundColor Yellow
    
    return $false
}

# Function to provide manual WinRM setup instructions
function Show-ManualWinRMInstructions {
    param([string]$ComputerName)
    
    Write-Host "`n" + "="*80 -ForegroundColor Yellow
    Write-Host "MANUAL WINRM SETUP REQUIRED" -ForegroundColor Yellow
    Write-Host "="*80 -ForegroundColor Yellow
    Write-Host "`nTo enable WinRM on $ComputerName, please do ONE of the following:" -ForegroundColor White
    Write-Host "`nOPTION 1 - Run on the target computer ($ComputerName):" -ForegroundColor Cyan
    Write-Host "  1. Open PowerShell as Administrator" -ForegroundColor White
    Write-Host "  2. Run: Enable-PSRemoting -Force" -ForegroundColor Green
    Write-Host "  3. Run: Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*' -Force" -ForegroundColor Green
    Write-Host "`nOPTION 2 - Use Group Policy (Domain Environment):" -ForegroundColor Cyan
    Write-Host "  1. Configure 'Allow automatic configuration of listeners' policy" -ForegroundColor White
    Write-Host "  2. Configure 'Allow Basic authentication' if needed" -ForegroundColor White
    Write-Host "`nOPTION 3 - Use Remote Registry/WMI (if accessible):" -ForegroundColor Cyan
    Write-Host "  1. Ensure Remote Registry service is running" -ForegroundColor White
    Write-Host "  2. Configure WinRM through registry modifications" -ForegroundColor White
    Write-Host "`nAfter enabling WinRM, re-run this script." -ForegroundColor Yellow
    Write-Host "="*80 -ForegroundColor Yellow
}

# Check for admin privileges and elevate if necessary
if (-not (Test-Administrator)) {
    Write-Host "Requesting administrative privileges..." -ForegroundColor Yellow
    
    # Get the current script path and arguments
    $scriptPath = $MyInvocation.MyCommand.Path
    $arguments = ""
    
    # Preserve the ComputerName parameter if provided
    if ($ComputerName) {
        $arguments = "-ComputerName `"$ComputerName`""
    }
    
    # Start new elevated process
    try {
        Start-Process -FilePath "powershell.exe" -ArgumentList "-File `"$scriptPath`" $arguments" -Verb RunAs -Wait
        exit
    }
    catch {
        Write-Error "Failed to elevate privileges: $_"
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Main script logic (runs with admin privileges)
Write-Host "Add Users to the Remote Desktop Users Group" -ForegroundColor Green
Write-Host "=" * 50

# Get computer name if not provided
if (-not $ComputerName) {
    $ComputerName = Read-Host -Prompt 'Please Enter the Computer Name (i.e. SLC-WS1GA04)'
}

# Test basic connectivity first
Write-Host "`nTesting basic connectivity to $ComputerName..." -ForegroundColor Yellow
if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
    Write-Host "✗ Cannot reach $ComputerName. Please verify:" -ForegroundColor Red
    Write-Host "  - Computer name is correct" -ForegroundColor White
    Write-Host "  - Computer is powered on and connected to network" -ForegroundColor White
    Write-Host "  - No network firewall blocking ICMP" -ForegroundColor White
    Read-Host "`nPress Enter to exit"
    exit 1
}
Write-Host "✓ Basic connectivity successful" -ForegroundColor Green

# Test WinRM connectivity
$winrmEnabled = Test-WinRMConnectivity -ComputerName $ComputerName

if (-not $winrmEnabled) {
    Write-Host "`nWinRM is not enabled on $ComputerName. Attempting automatic enablement..." -ForegroundColor Yellow
    
    $winrmEnabled = Enable-RemoteWinRM -ComputerName $ComputerName
    
    if (-not $winrmEnabled) {
        Write-Host "`nAutomatic WinRM enablement failed." -ForegroundColor Red
        Show-ManualWinRMInstructions -ComputerName $ComputerName
        Read-Host "`nPress Enter to exit"
        exit 1
    }
}

try {
    # Create remote session
    Write-Host "`nEstablishing PowerShell session to $ComputerName..." -ForegroundColor Yellow
    $session = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
    Write-Host "✓ PowerShell session established successfully" -ForegroundColor Green
    
    # Get username to add
    $userName = Read-Host -Prompt "Please Enter the User Name (i.e. VHASLCMousM)"
    
    # Add user to Remote Desktop Users group
    Write-Host "`nAdding user '$userName' to Remote Desktop Users group..." -ForegroundColor Yellow
    Invoke-Command -Session $session -ScriptBlock {
        param($user)
        Add-LocalGroupMember -Group "Remote Desktop Users" -Member $user
    } -ArgumentList $userName
    
    Write-Host "User added successfully!" -ForegroundColor Green
    
    # Display current group members for verification
    Write-Host "`nPlease check the list below to verify the User was added to the Remote Desktop Users Group:" -ForegroundColor Cyan
    Write-Host "-" * 80
    
    $groupMembers = Invoke-Command -Session $session -ScriptBlock {
        Get-LocalGroupMember -Group "Remote Desktop Users" | Select-Object -Property Name
    }
    
    $groupMembers | Format-Table -AutoSize
    
    # Clean up session
    Remove-PSSession -Session $session
    
    Write-Host "`nOperation completed successfully!" -ForegroundColor Green
}
catch {
    Write-Host "`nAn error occurred during the operation:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    # Provide specific guidance based on error type
    if ($_.Exception.Message -match "WinRM|remote server|access is denied") {
        Write-Host "`nThis appears to be a WinRM/Remote Access issue. Possible solutions:" -ForegroundColor Yellow
        Write-Host "1. Verify WinRM is properly configured on $ComputerName" -ForegroundColor White
        Write-Host "2. Check if you have administrative rights on the target computer" -ForegroundColor White
        Write-Host "3. Ensure Windows Firewall allows WinRM traffic" -ForegroundColor White
        Write-Host "4. Try running: winrm quickconfig on the target computer" -ForegroundColor White
    }
    elseif ($_.Exception.Message -match "user|member|group") {
        Write-Host "`nThis appears to be a user/group related issue. Check:" -ForegroundColor Yellow
        Write-Host "1. User name format (try DOMAIN\\Username or Username@domain.com)" -ForegroundColor White
        Write-Host "2. User exists in Active Directory" -ForegroundColor White
        Write-Host "3. User is not already a member of the group" -ForegroundColor White
    }
    else {
        Write-Host "`nGeneral troubleshooting steps:" -ForegroundColor Yellow
        Write-Host "1. Verify network connectivity to $ComputerName" -ForegroundColor White
        Write-Host "2. Check if target computer is domain-joined" -ForegroundColor White
        Write-Host "3. Ensure you have appropriate permissions" -ForegroundColor White
    }
    
    # Clean up session if it exists
    if ($session) {
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    }
}

# Pause before exit
Write-Host "`nPress Enter to exit..." -ForegroundColor Gray
Read-Host
