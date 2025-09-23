<#
.SYNOPSIS
    Adds a printer to the default user profile for all existing and new users.

.DESCRIPTION
    This script adds a specified local or network printer to the default account 
    for all existing/new users on a target computer. The script requires administrator 
    privileges and can operate on local or remote computers.

.NOTES
    ┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
    │ ORIGIN STORY                                                                                │ 
    ├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
    │   NAME:       : add_printer_for_all_users.ps1                                           │
    │   DATE        : 2020-05-11                                                                 │
    │   AUTHOR      : Dallas Bleak (Dallas.Bleak@va.govs)                                         │
    │   VERSION     : 2.0                                                                         │
    │   REVISION    : 2025-09-23 Converted from batch to PowerShell                               │
    └─────────────────────────────────────────────────────────────────────────────────────────────┘
    
    IMPORTANT: 
    - Printer names with spaces will NOT be accepted
    - Script requires administrator privileges
    - If you get a pop-up ERROR box, close script and try again

.EXAMPLE
    .\add_printer_for_all_users.ps1
    Runs the script interactively with prompts for computer name and printer name.
#>

#Requires -RunAsAdministrator

# Set console colors and clear screen
function Set-ConsoleAppearance {
    param(
        [string]$BackgroundColor = "Black",
        [string]$ForegroundColor = "White"
    )
    
    $Host.UI.RawUI.BackgroundColor = $BackgroundColor
    $Host.UI.RawUI.ForegroundColor = $ForegroundColor
    Clear-Host
}

# Function to test network connectivity
function Test-ComputerConnectivity {
    param(
        [string]$ComputerName
    )
    
    Write-Host ""
    Write-Host "Attempting to ping $ComputerName" -ForegroundColor Yellow
    
    try {
        $pingResult = Test-Connection -ComputerName $ComputerName -Count 3 -Quiet
        return $pingResult
    }
    catch {
        Write-Host "Error testing connectivity: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to add printer for all users
function Add-PrinterForAllUsers {
    param(
        [string]$ComputerName,
        [string]$PrinterName
    )
    
    Clear-Host
    Write-Host ""
    Write-Host "Attempting to add $PrinterName for all users on $ComputerName" -ForegroundColor Green
    
    try {
        # Use rundll32 with printui.dll to add printer for all users
        $arguments = "PrintUIEntry /ga /c \\$ComputerName /n \\$PrinterName"
        Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,$arguments" -Wait -NoNewWindow
        
        Write-Host "Printer installation command executed successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error adding printer: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to restart print spooler service
function Restart-PrintSpooler {
    param(
        [string]$ComputerName
    )
    
    try {
        Write-Host "Stopping print spooler service on $ComputerName..." -ForegroundColor Yellow
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            # Local computer
            Stop-Service -Name "Spooler" -Force
            Start-Service -Name "Spooler"
        }
        else {
            # Remote computer - use sc command for compatibility
            Start-Process -FilePath "sc.exe" -ArgumentList "\\$ComputerName stop spooler" -Wait -NoNewWindow
            Start-Process -FilePath "sc.exe" -ArgumentList "\\$ComputerName start spooler" -Wait -NoNewWindow
        }
        
        Write-Host ""
        Write-Host "Print Spooler Service restarted successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error restarting print spooler: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to get user input with validation
function Get-ValidatedInput {
    param(
        [string]$Prompt,
        [switch]$AllowSpaces = $false
    )
    
    do {
        $input = Read-Host $Prompt
        
        if ([string]::IsNullOrWhiteSpace($input)) {
            Write-Host "Input cannot be empty. Please try again." -ForegroundColor Red
            continue
        }
        
        if (-not $AllowSpaces -and $input.Contains(" ")) {
            Write-Host "Input cannot contain spaces. Please try again." -ForegroundColor Red
            continue
        }
        
        return $input.Trim()
    } while ($true)
}

# Function to get Yes/No input
function Get-YesNoInput {
    param(
        [string]$Prompt
    )
    
    do {
        $input = Read-Host "$Prompt (Y/N)"
        $input = $input.Trim().ToLower()
        
        if ($input -eq "y" -or $input -eq "yes") {
            return $true
        }
        elseif ($input -eq "n" -or $input -eq "no") {
            return $false
        }
        else {
            Write-Host "Please enter Y or N." -ForegroundColor Red
        }
    } while ($true)
}

# Main script execution
function Start-PrinterInstallation {
    do {
        # Initialize and display header
        Set-ConsoleAppearance -ForegroundColor "White"
        Write-Host ""
        Write-Host "This script adds the specified local or network printer" -ForegroundColor Cyan
        Write-Host "to the default account for all existing/new users." -ForegroundColor Cyan
        Write-Host "*IMPORTANT* Printer names with spaces will NOT be accepted." -ForegroundColor Yellow
        Write-Host "*IMPORTANT* If you get a pop up ERROR box close script and try again." -ForegroundColor Yellow
        Write-Host "************************************************************" -ForegroundColor Cyan
        Write-Host ""
        
        # Get user input
        $computerName = Get-ValidatedInput -Prompt "Enter target computer name"
        $printerName = Get-ValidatedInput -Prompt "Enter Printserver/Printername (do not include \\)"
        
        # Get local hostname for comparison
        $localHostName = $env:COMPUTERNAME
        
        # Check if target is local or remote
        if ($localHostName -eq $computerName) {
            Write-Host "Target is local computer." -ForegroundColor Green
            $pingSuccess = $true
        }
        else {
            # Test connectivity to remote computer
            $pingSuccess = Test-ComputerConnectivity -ComputerName $computerName
            
            if (-not $pingSuccess) {
                Set-ConsoleAppearance -BackgroundColor "Black" -ForegroundColor "Red"
                Write-Host ""
                Write-Host "Computer $computerName does not respond to ping." -ForegroundColor Red
                Write-Host ""
                
                if (Get-YesNoInput -Prompt "Would you like to try again?") {
                    continue
                }
                else {
                    break
                }
            }
        }
        
        # Add printer if connectivity test passed
        if ($pingSuccess) {
            $printerAdded = Add-PrinterForAllUsers -ComputerName $computerName -PrinterName $printerName
            
            if ($printerAdded) {
                # Ask about restarting spooler
                Set-ConsoleAppearance -BackgroundColor "Blue" -ForegroundColor "White"
                Write-Host ""
                Write-Host "New printers will NOT appear until spooler is restarted." -ForegroundColor Yellow
                Write-Host ""
                
                if (Get-YesNoInput -Prompt "Reset print spooler?") {
                    $spoolerRestarted = Restart-PrintSpooler -ComputerName $computerName
                    
                    if ($spoolerRestarted) {
                        Write-Host ""
                        Read-Host "Press Enter to continue"
                    }
                }
            }
            
            # Ask if user wants to add another printer
            Set-ConsoleAppearance -ForegroundColor "White"
            Write-Host ""
            Write-Host "Would you like to add another printer to $computerName or from a different computer?" -ForegroundColor Cyan
            Write-Host ""
            
            if (-not (Get-YesNoInput -Prompt "Add another printer?")) {
                break
            }
        }
    } while ($true)
    
    Write-Host ""
    Write-Host "Script completed." -ForegroundColor Green
}

# Script entry point
try {
    Write-Host "Starting Printer Installation Script..." -ForegroundColor Green
    Start-PrinterInstallation
}
catch {
    Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    # Reset console colors
    Set-ConsoleAppearance -ForegroundColor "White"
}
