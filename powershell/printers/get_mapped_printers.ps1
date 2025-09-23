<#
.SYNOPSIS
    Get Mapped Printers - Retrieves mapped printers for a specified computer
    
.DESCRIPTION
    This script retrieves the mapped printers for a specified computer by querying
    the registry on the remote machine. It requires administrative privileges and
    will prompt for elevation if not already running as administrator.
    
.PARAMETER ComputerName
    The name of the computer to query for mapped printers
    
.EXAMPLE
    .\get_mapped_printers.ps1
    Prompts for computer name and retrieves mapped printers
    
.EXAMPLE
    .\get_mapped_printers.ps1 -ComputerName "COMPUTER01"
    Retrieves mapped printers for COMPUTER01
    
.NOTES
    Name:     get_mapped_printers.ps1
    Purpose:  Retrieves the mapped printers for the specified computer
    Author:   Dallas Bleak
    Revision: September 2025 - Combined batch and PowerShell functionality
┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ ORIGIN STORY                                                                                │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 2020-05-09                                                                  │       
│   AUTHOR      : Dallas Bleak (Dallas.Bleak@va.gov)                                          │
│   DESCRIPTION : Initial Draft                                                               │
│   VERSION     : 1.1                                                                         │
│   REVISION    : 2025-09-23 - Combined batch and PowerShell functionality                    │
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ComputerName
)

# Function to check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to restart script with elevated privileges
function Start-ElevatedProcess {
    Write-Host "Requesting administrative privileges..." -ForegroundColor Yellow
    
    $arguments = ""
    if ($ComputerName) {
        $arguments = "-ComputerName `"$ComputerName`""
    }
    
    try {
        Start-Process PowerShell -Verb RunAs -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`" $arguments"
        exit
    }
    catch {
        Write-Error "Failed to elevate privileges: $($_.Exception.Message)"
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Function to get mapped printers from remote computer
function Get-MappedPrinters {
    param([string]$TargetComputer)
    
    Write-Host "Connecting to $TargetComputer..." -ForegroundColor Green
    
    try {
        $printers = Invoke-Command -ComputerName $TargetComputer -ScriptBlock {
            Get-ChildItem Registry::\HKEY_Users | 
            Where-Object { $_.PSChildName -NotMatch ".DEFAULT|S-1-5-18|S-1-5-19|S-1-5-20|_Classes" } | 
            Select-Object -ExpandProperty PSChildName | 
            ForEach-Object { 
                try {
                    Get-ChildItem Registry::\HKEY_Users\$_\Printers\Connections -Recurse -ErrorAction SilentlyContinue | 
                    Select-Object -ExpandProperty Name
                }
                catch {
                    # Silently continue if user doesn't have printer connections
                }
            }
        } -ErrorAction Stop
        
        if ($printers) {
            Write-Host "`nMapped Printers found on $TargetComputer`:" -ForegroundColor Cyan
            Write-Host ("=" * 50) -ForegroundColor Cyan
            $printers | ForEach-Object {
                $printerName = $_ -replace "HKEY_USERS\\[^\\]+\\Printers\\Connections\\", ""
                Write-Host "  $printerName" -ForegroundColor White
            }
            Write-Host ("=" * 50) -ForegroundColor Cyan
            Write-Host "Total printers found: $($printers.Count)" -ForegroundColor Green
        }
        else {
            Write-Host "No mapped printers found on $TargetComputer" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Failed to connect to $TargetComputer`: $($_.Exception.Message)"
        Write-Host "Please verify:" -ForegroundColor Yellow
        Write-Host "  - Computer name is correct" -ForegroundColor Yellow
        Write-Host "  - Computer is online and accessible" -ForegroundColor Yellow
        Write-Host "  - You have administrative rights on the target computer" -ForegroundColor Yellow
        Write-Host "  - Windows Remote Management (WinRM) is enabled on the target" -ForegroundColor Yellow
    }
}

# Main script execution
Clear-Host
Write-Host "Get Mapped Printers Utility" -ForegroundColor Cyan
Write-Host ("=" * 30) -ForegroundColor Cyan

# Check for administrative privileges
if (-not (Test-Administrator)) {
    Start-ElevatedProcess
}

Write-Host "Running with administrative privileges" -ForegroundColor Green

# Get computer name if not provided as parameter
if (-not $ComputerName) {
    do {
        $ComputerName = Read-Host -Prompt "`nEnter Computer Name"
        if ([string]::IsNullOrWhiteSpace($ComputerName)) {
            Write-Host "Computer name cannot be empty. Please try again." -ForegroundColor Red
        }
    } while ([string]::IsNullOrWhiteSpace($ComputerName))
}

# Validate computer name format (basic validation)
if ($ComputerName -match '[<>:"/\\|?*]') {
    Write-Error "Invalid computer name format. Computer names cannot contain special characters."
    Read-Host "Press Enter to exit"
    exit 1
}

# Execute the main function
Get-MappedPrinters -TargetComputer $ComputerName

Write-Host "`nOperation completed." -ForegroundColor Green
Read-Host "Press Enter to exit"
