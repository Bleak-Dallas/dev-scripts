<#
.SYNOPSIS
    Launches BigFix Remote Control (BFRC) on a remote target computer with enhanced error handling and logging.

.DESCRIPTION
    This script connects to a remote computer and launches BigFix Remote Control with a specified connection code.
    It includes robust error handling, logging, service management, and validation of the connection attempt.
    All operations are logged to C:\Temp\logs for audit and troubleshooting purposes.

.PARAMETER TargetComputer
    The FQDN of the target computer (e.g., SLC-LT126450.v19.med.va.gov)

.PARAMETER ConnectionCode
    The BigFix Remote Control connection code

.PARAMETER Credential
    Optional credentials for remote operations (if not provided, uses current user context)

.PARAMETER MaxRetries
    Maximum number of connection attempts (default: 3)

.PARAMETER BigFixPath
    Custom path to BigFix Remote Control executable (default: standard installation path)

.PARAMETER Verbose
    Enable verbose output for detailed troubleshooting

.EXAMPLE
    .\launch_big_fix_remote_enhanced.ps1
    Prompts for target computer and connection code, then launches BFRC

.EXAMPLE
    .\launch_big_fix_remote_enhanced.ps1 -TargetComputer "SLC-LT126450.v19.med.va.gov" -ConnectionCode "123456"
    Launches BFRC on specified target with provided connection code

.NOTES
    ┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
    │ ORIGIN STORY                                                                                │ 
    ├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
    │   NAME:       : launch_big_fix_remote_enhanced.ps1                                          │
    │   DATE        : 2025-09-23                                                                  │
    │   AUTHOR      : Dallas Bleak (Dallas.Bleak@va.gov)(based on original by John Phung)         │
    │   VERSION     : 2.0                                                                         │    
    │   RUN AS      : Elevated PowerShell (Run as Administrator) recommended.                     │
    └─────────────────────────────────────────────────────────────────────────────────────────────┘ 
    
    Requirements:
    - PowerShell 5.1 or later
    - Administrative privileges on target computer
    - BigFix Remote Control installed on target
    - Network connectivity to target computer
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,}$')]
    [string]$TargetComputer,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^\d+$')]
    [string]$ConnectionCode,
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,
    
    [Parameter(Mandatory = $false)]
    [string]$BigFixPath = "C:\Program Files (x86)\BigFix\Remote Control\Target\trc_gui.exe",
    
    [Parameter(Mandatory = $false)]
    [switch]$VerboseLogging
)

# Global variables
$script:LogFile = ""
$script:StartTime = Get-Date

#region Logging Functions

function Initialize-LogFile {
    <#
    .SYNOPSIS
        Initializes the log file and creates necessary directories
    #>
    
    try {
        # Create log directory structure
        $logDir = "C:\Temp\logs"
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
            Write-Host "✓ Created log directory: $logDir" -ForegroundColor Green
        }
        
        # Create timestamped log file
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $script:LogFile = Join-Path $logDir "BigFixRemote_$timestamp.log"
        
        # Initialize log file with header
        $header = @"
================================================================================
BigFix Remote Control Launch Script - Enhanced Version
================================================================================
Script Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Target Computer: $TargetComputer
Connection Code: $ConnectionCode
Max Retries: $MaxRetries
BigFix Path: $BigFixPath
================================================================================

"@
        
        $header | Out-File -FilePath $script:LogFile -Encoding UTF8
        Write-Host "✓ Log file initialized: $script:LogFile" -ForegroundColor Green
        
        return $true
    }
    catch {
        Write-Warning "Failed to initialize log file: $($_.Exception.Message)"
        return $false
    }
}

function Write-LogEntry {
    <#
    .SYNOPSIS
        Writes an entry to the log file with timestamp and severity level
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    
    try {
        $logEntry | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
        
        if ($VerboseLogging -or $Level -in @('WARNING', 'ERROR')) {
            switch ($Level) {
                'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
                'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
                'ERROR' { Write-Host $logEntry -ForegroundColor Red }
                'DEBUG' { Write-Host $logEntry -ForegroundColor Cyan }
                default { Write-Host $logEntry -ForegroundColor White }
            }
        }
    }
    catch {
        Write-Warning "Failed to write to log file: $($_.Exception.Message)"
    }
}

#endregion

#region Input Validation Functions

function Get-ValidatedTargetComputer {
    <#
    .SYNOPSIS
        Prompts for and validates target computer name
    #>
    
    if ([string]::IsNullOrWhiteSpace($TargetComputer)) {
        do {
            $userInput = Read-Host "`nEnter Target PC Name using FQDN (e.g., SLC-LT126450.v19.med.va.gov)"
            $userInput = $userInput.Trim().ToUpper()
            
            if ($userInput -match '^[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,}$') {
                $script:TargetComputer = $userInput
                Write-LogEntry "Target computer set to: $script:TargetComputer"
                break
            }
            else {
                Write-Host "✗ Invalid FQDN format. Please use format: hostname.domain.com" -ForegroundColor Red
                Write-LogEntry "Invalid FQDN format entered: $userInput" -Level 'WARNING'
            }
        } while ($true)
    }
    else {
        $script:TargetComputer = $TargetComputer.Trim().ToUpper()
        Write-LogEntry "Target computer parameter: $script:TargetComputer"
    }
}

function Get-ValidatedConnectionCode {
    <#
    .SYNOPSIS
        Prompts for and validates connection code
    #>
    
    if ([string]::IsNullOrWhiteSpace($ConnectionCode)) {
        do {
            $userInput = Read-Host "`nEnter BFRC Connection Code (numbers only)"
            $userInput = $userInput.Trim()
            
            if ($userInput -match '^\d+$') {
                $script:ConnectionCode = $userInput
                Write-LogEntry "Connection code set (length: $($userInput.Length))"
                break
            }
            else {
                Write-Host "✗ Invalid connection code. Please enter numbers only." -ForegroundColor Red
                Write-LogEntry "Invalid connection code format entered" -Level 'WARNING'
            }
        } while ($true)
    }
    else {
        $script:ConnectionCode = $ConnectionCode.Trim()
        Write-LogEntry "Connection code parameter provided (length: $($ConnectionCode.Length))"
    }
}

#endregion

#region Connectivity and Validation Functions

function Test-TargetConnectivity {
    <#
    .SYNOPSIS
        Tests basic network connectivity to target computer
    #>
    param([string]$ComputerName)
    
    Write-Host "`nTesting connectivity to $ComputerName..." -ForegroundColor Yellow
    Write-LogEntry "Testing connectivity to $ComputerName"
    
    try {
        $result = Test-Connection -ComputerName $ComputerName -Count 2 -Quiet -ErrorAction Stop
        
        if ($result) {
            Write-Host "✓ $ComputerName is ONLINE" -ForegroundColor Green
            Write-LogEntry "$ComputerName is online - connectivity test passed" -Level 'SUCCESS'
            return $true
        }
        else {
            Write-Host "✗ $ComputerName is OFFLINE" -ForegroundColor Red
            Write-LogEntry "$ComputerName is offline - connectivity test failed" -Level 'ERROR'
            return $false
        }
    }
    catch {
        Write-Host "✗ Connectivity test failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry "Connectivity test error: $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

function Test-BigFixInstallation {
    <#
    .SYNOPSIS
        Validates BigFix Remote Control installation on target computer
    #>
    param([string]$ComputerName, [string]$ExecutablePath)
    
    Write-Host "Validating BigFix installation..." -ForegroundColor Yellow
    Write-LogEntry "Validating BigFix installation at: $ExecutablePath"
    
    try {
        $scriptBlock = {
            param($Path)
            Test-Path $Path
        }
        
        $sessionParams = @{
            ComputerName = $ComputerName
            ScriptBlock  = $scriptBlock
            ArgumentList = $ExecutablePath
            ErrorAction  = 'Stop'
        }
        
        if ($Credential) {
            $sessionParams.Credential = $Credential
        }
        
        $exists = Invoke-Command @sessionParams
        
        if ($exists) {
            Write-Host "✓ BigFix Remote Control found on target" -ForegroundColor Green
            Write-LogEntry "BigFix installation validated successfully" -Level 'SUCCESS'
            return $true
        }
        else {
            Write-Host "✗ BigFix Remote Control not found at: $ExecutablePath" -ForegroundColor Red
            Write-LogEntry "BigFix executable not found at: $ExecutablePath" -Level 'ERROR'
            return $false
        }
    }
    catch {
        Write-Host "✗ Failed to validate BigFix installation: $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry "BigFix validation error: $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

function Test-TRCTargetService {
    <#
    .SYNOPSIS
        Checks the status of the TRCTARGET service
    #>
    param([string]$ComputerName)
    
    Write-Host "Checking TRCTARGET service status..." -ForegroundColor Yellow
    Write-LogEntry "Checking TRCTARGET service status on $ComputerName"
    
    try {
        $scriptBlock = {
            Get-Service -Name "TRCTARGET" -ErrorAction Stop
        }
        
        $sessionParams = @{
            ComputerName = $ComputerName
            ScriptBlock  = $scriptBlock
            ErrorAction  = 'Stop'
        }
        
        if ($Credential) {
            $sessionParams.Credential = $Credential
        }
        
        $service = Invoke-Command @sessionParams
        
        Write-Host "✓ TRCTARGET service status: $($service.Status)" -ForegroundColor Green
        Write-LogEntry "TRCTARGET service status: $($service.Status)" -Level 'INFO'
        
        return $service.Status -eq 'Running'
    }
    catch {
        Write-Host "✗ Failed to check TRCTARGET service: $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry "TRCTARGET service check error: $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

#endregion

#region BigFix Operations

function Start-TRCTargetService {
    <#
    .SYNOPSIS
        Starts or restarts the TRCTARGET service
    #>
    param([string]$ComputerName)
    
    Write-Host "Restarting TRCTARGET service..." -ForegroundColor Yellow
    Write-LogEntry "Attempting to restart TRCTARGET service on $ComputerName"
    
    try {
        $scriptBlock = {
            Restart-Service -Name "TRCTARGET" -Force -ErrorAction Stop
            Start-Sleep -Seconds 3
            Get-Service -Name "TRCTARGET"
        }
        
        $sessionParams = @{
            ComputerName = $ComputerName
            ScriptBlock  = $scriptBlock
            ErrorAction  = 'Stop'
        }
        
        if ($Credential) {
            $sessionParams.Credential = $Credential
        }
        
        $service = Invoke-Command @sessionParams
        
        if ($service.Status -eq 'Running') {
            Write-Host "✓ TRCTARGET service restarted successfully" -ForegroundColor Green
            Write-LogEntry "TRCTARGET service restarted successfully" -Level 'SUCCESS'
            return $true
        }
        else {
            Write-Host "✗ TRCTARGET service failed to start: $($service.Status)" -ForegroundColor Red
            Write-LogEntry "TRCTARGET service restart failed: $($service.Status)" -Level 'ERROR'
            return $false
        }
    }
    catch {
        Write-Host "✗ Failed to restart TRCTARGET service: $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry "TRCTARGET service restart error: $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

function Invoke-BigFixRemoteControl {
    <#
    .SYNOPSIS
        Launches BigFix Remote Control with the specified connection code
    #>
    param(
        [string]$ComputerName,
        [string]$ConnectionCode,
        [string]$ExecutablePath,
        [int]$AttemptNumber
    )
    
    Write-Host "`nAttempt #$AttemptNumber - Launching BFRC on $ComputerName..." -ForegroundColor Cyan
    Write-LogEntry "Attempt #$AttemptNumber - Launching BFRC with connection code: $ConnectionCode"
    
    try {
        $scriptBlock = {
            param($ExePath, $ConnCode)
            & $ExePath "--connect-with-cc=$ConnCode"
        }
        
        $sessionParams = @{
            ComputerName = $ComputerName
            ScriptBlock  = $scriptBlock
            ArgumentList = @($ExecutablePath, $ConnectionCode)
            ErrorAction  = 'Stop'
        }
        
        if ($Credential) {
            $sessionParams.Credential = $Credential
        }
        
        Invoke-Command @sessionParams
        
        Write-Host "✓ BFRC launch command executed" -ForegroundColor Green
        Write-LogEntry "BFRC launch command executed successfully" -Level 'SUCCESS'
        
        # Wait for the application to initialize
        Start-Sleep -Seconds 5
        
        return $true
    }
    catch {
        Write-Host "✗ Failed to launch BFRC: $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry "BFRC launch error: $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

function Test-BigFixConnection {
    <#
    .SYNOPSIS
        Checks BigFix logs to verify successful connection
    #>
    param([string]$ComputerName, [string]$ConnectionCode)
    
    Write-Host "Verifying connection in BigFix logs..." -ForegroundColor Yellow
    Write-LogEntry "Checking BigFix logs for connection code: $ConnectionCode"
    
    try {
        # Get current day abbreviation for log file
        $dayOfWeek = (Get-Date).DayOfWeek.ToString().Substring(0, 3)
        $logPath = "C:\ProgramData\BigFix\Remote Control\trc_gui_$dayOfWeek.log"
        
        $scriptBlock = {
            param($LogPath, $ConnCode)
            
            if (Test-Path $LogPath) {
                $content = Get-Content $LogPath -ErrorAction Stop
                $pattern = "Connection code '$ConnCode'"
                $matches = $content | Select-String -Pattern $pattern
                return $matches.Count -gt 0
            }
            else {
                return $false
            }
        }
        
        $sessionParams = @{
            ComputerName = $ComputerName
            ScriptBlock  = $scriptBlock
            ArgumentList = @($logPath, $ConnectionCode)
            ErrorAction  = 'Stop'
        }
        
        if ($Credential) {
            $sessionParams.Credential = $Credential
        }
        
        $connectionFound = Invoke-Command @sessionParams
        
        if ($connectionFound) {
            Write-Host "✓ Connection verified in BigFix logs" -ForegroundColor Green
            Write-LogEntry "Connection code found in BigFix logs - connection successful" -Level 'SUCCESS'
            return $true
        }
        else {
            Write-Host "✗ Connection code not found in logs" -ForegroundColor Yellow
            Write-LogEntry "Connection code not found in BigFix logs" -Level 'WARNING'
            return $false
        }
    }
    catch {
        Write-Host "✗ Failed to check BigFix logs: $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry "BigFix log check error: $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

#endregion

#region Main Execution Functions

function Start-BigFixRemoteSession {
    <#
    .SYNOPSIS
        Main function to orchestrate the BigFix Remote Control connection
    #>
    
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
    Write-Host "BigFix Remote Control Launch - Enhanced Version" -ForegroundColor Cyan
    Write-Host "="*80 -ForegroundColor Cyan
    
    # Initialize logging
    if (-not (Initialize-LogFile)) {
        Write-Warning "Continuing without logging..."
    }
    
    Write-LogEntry "Script execution started"
    
    try {
        # Get and validate inputs
        Get-ValidatedTargetComputer
        Get-ValidatedConnectionCode
        
        # Test connectivity
        if (-not (Test-TargetConnectivity -ComputerName $script:TargetComputer)) {
            throw "Target computer is not reachable"
        }
        
        # Validate BigFix installation
        if (-not (Test-BigFixInstallation -ComputerName $script:TargetComputer -ExecutablePath $BigFixPath)) {
            throw "BigFix Remote Control is not installed or not accessible"
        }
        
        # Check service status
        $serviceRunning = Test-TRCTargetService -ComputerName $script:TargetComputer
        
        # Attempt connection with retries
        $attemptCount = 0
        $connectionSuccessful = $false
        
        while ($attemptCount -lt $MaxRetries -and -not $connectionSuccessful) {
            $attemptCount++
            
            # Launch BigFix Remote Control
            if (Invoke-BigFixRemoteControl -ComputerName $script:TargetComputer -ConnectionCode $script:ConnectionCode -ExecutablePath $BigFixPath -AttemptNumber $attemptCount) {
                
                # Wait a bit longer for logs to be written
                Start-Sleep -Seconds 3
                
                # Check if connection was successful
                if (Test-BigFixConnection -ComputerName $script:TargetComputer -ConnectionCode $script:ConnectionCode) {
                    $connectionSuccessful = $true
                    break
                }
            }
            
            # If this was attempt 3 and service restart hasn't been tried, restart service
            if ($attemptCount -eq 3 -and -not $connectionSuccessful) {
                Write-Host "`nAttempting service restart before final attempt..." -ForegroundColor Yellow
                if (Start-TRCTargetService -ComputerName $script:TargetComputer) {
                    # Give one more attempt after service restart
                    $MaxRetries++
                }
            }
            
            if (-not $connectionSuccessful -and $attemptCount -lt $MaxRetries) {
                Write-Host "Attempt #$attemptCount failed. Waiting before retry..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
        }
        
        # Final result
        if ($connectionSuccessful) {
            Write-Host "`n" + "="*80 -ForegroundColor Green
            Write-Host "SUCCESS: BigFix Remote Control connection established!" -ForegroundColor Green
            Write-Host "Target: $script:TargetComputer" -ForegroundColor Green
            Write-Host "Connection Code: $script:ConnectionCode" -ForegroundColor Green
            Write-Host "Attempts: $attemptCount" -ForegroundColor Green
            Write-Host "="*80 -ForegroundColor Green
            Write-LogEntry "BigFix Remote Control connection successful after $attemptCount attempts" -Level 'SUCCESS'
        }
        else {
            Write-Host "`n" + "="*80 -ForegroundColor Red
            Write-Host "FAILED: Unable to establish BigFix Remote Control connection" -ForegroundColor Red
            Write-Host "Target: $script:TargetComputer" -ForegroundColor Red
            Write-Host "Connection Code: $script:ConnectionCode" -ForegroundColor Red
            Write-Host "Attempts: $attemptCount" -ForegroundColor Red
            Write-Host "="*80 -ForegroundColor Red
            Write-LogEntry "BigFix Remote Control connection failed after $attemptCount attempts" -Level 'ERROR'
            
            # Provide troubleshooting guidance
            Write-Host "`nTroubleshooting suggestions:" -ForegroundColor Yellow
            Write-Host "1. Verify the connection code is correct and active" -ForegroundColor White
            Write-Host "2. Check if BigFix Remote Control is running on the target" -ForegroundColor White
            Write-Host "3. Ensure TRCTARGET service is running" -ForegroundColor White
            Write-Host "4. Verify network connectivity and firewall settings" -ForegroundColor White
            Write-Host "5. Check BigFix logs on the target computer" -ForegroundColor White
        }
        
    }
    catch {
        Write-Host "`n✗ Script execution failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry "Script execution failed: $($_.Exception.Message)" -Level 'ERROR'
    }
    finally {
        # Log completion
        $endTime = Get-Date
        $duration = $endTime - $script:StartTime
        Write-LogEntry "Script execution completed. Duration: $($duration.ToString('mm\:ss'))"
        
        if ($script:LogFile) {
            Write-Host "`nLog file: $script:LogFile" -ForegroundColor Cyan
        }
        
        Write-Host "`nPress any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

#endregion

# Script execution entry point
if ($MyInvocation.InvocationName -ne '.') {
    Start-BigFixRemoteSession
}
