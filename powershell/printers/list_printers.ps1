<#
.SYNOPSIS
    Lists available printers on local or remote computers.

.DESCRIPTION
    This script discovers and displays all available printers on a specified computer,
    including local printers, network printers, and user-specific printer connections.
    It uses multiple discovery methods for comprehensive printer enumeration.

.NOTES
    Name:      list_printers.ps1
    Purpose:   List and discover available printers
    Author:    Dallas Bleak
    Created:   December 2024
    Based on:  delete_printer_for_all_users.ps1
    ┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
    │ ORIGIN STORY                                                                                │ 
    ├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
    │   NAME:       : list_printers.ps1                                                           │
    │   DATE        : 2024-12-23                                                                  │
    │   AUTHOR      : Dallas Bleak (Dallas.Bleak@va.gov)                                          │
    │   VERSION     : 1.0                                                                         │
    │   BASED ON    : delete_printer_for_all_users.ps1 (printer discovery functions)              │
    └─────────────────────────────────────────────────────────────────────────────────────────────┘

.EXAMPLE
    .\list_printers.ps1
    Run the script and follow onscreen directions to list printers
#>

function Write-ColoredText {
    param(
        [string]$Text,
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::White,
        [ConsoleColor]$BackgroundColor = [ConsoleColor]::Black
    )
    $originalForeground = $Host.UI.RawUI.ForegroundColor
    $originalBackground = $Host.UI.RawUI.BackgroundColor
    
    $Host.UI.RawUI.ForegroundColor = $ForegroundColor
    $Host.UI.RawUI.BackgroundColor = $BackgroundColor
    
    Write-Host $Text
    
    $Host.UI.RawUI.ForegroundColor = $originalForeground
    $Host.UI.RawUI.BackgroundColor = $originalBackground
}

function Test-ComputerConnectivity {
    param([string]$ComputerName)
    
    # Trim any whitespace from computer name
    $ComputerName = $ComputerName.Trim()
    
    Write-Host ""
    Write-Host "Attempting to ping $ComputerName"
    
    try {
        $pingResult = Test-Connection -ComputerName $ComputerName -Count 3 -ErrorAction Stop
        if ($pingResult) {
            Write-Host "Ping successful - $($pingResult.Count) replies received" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "Ping failed - no replies received" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Ping failed with error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Get-UserPrinterConnections {
    param(
        [string]$ComputerName,
        [string]$UserName = $null  # If null, query all users
    )
    
    # Trim any whitespace from computer name
    $ComputerName = $ComputerName.Trim()
    
    $scriptBlock = {
        param($UserName)
        
        function Get-ConnectionsFromHivePath($hivePath, $userSid, $userName) {
            if (-not (Test-Path $hivePath)) { return @() }
            $connections = @()
            Get-ChildItem $hivePath -ErrorAction SilentlyContinue | ForEach-Object {
                $parts = ($_.PSChildName -split ',' | Where-Object { $_ })  # remove leading empty entries from ",,"
                if ($parts.Count -ge 2) {
                    $connections += [pscustomobject]@{
                        Computer     = $env:COMPUTERNAME
                        User         = $userName
                        SID          = $userSid
                        PrintServer  = $parts[0]
                        PrinterShare = $parts[1]
                        Connection   = "\\$($parts[0])\$($parts[1])"
                        RegistryKey  = $_.Name
                    }
                }
            }
            return $connections
        }
        
        $allConnections = @()
        
        if ($UserName) {
            # Query specific user
            try {
                $sid = ([System.Security.Principal.NTAccount]$UserName).Translate([System.Security.Principal.SecurityIdentifier]).Value
                
                # Try loaded hive first
                $loadedPath = "Registry::HKEY_USERS\$sid\Printers\Connections"
                $results = Get-ConnectionsFromHivePath $loadedPath $sid $UserName
                
                # If not loaded, try loading the profile hive temporarily
                if (-not $results) {
                    $prof = Get-CimInstance Win32_UserProfile -Filter "SID='$sid'" -ErrorAction SilentlyContinue
                    if ($prof -and (Test-Path ($prof.LocalPath + "\NTUSER.DAT"))) {
                        $tempHive = "HKU\Temp_$($sid -replace '[^A-Za-z0-9_]','_')"
                        try {
                            Start-Process -FilePath "reg.exe" -ArgumentList "load", $tempHive, ($prof.LocalPath + "\NTUSER.DAT") -WindowStyle Hidden -Wait -RedirectStandardOutput $null -RedirectStandardError $null -ErrorAction SilentlyContinue
                            $results = Get-ConnectionsFromHivePath ("Registry::" + $tempHive + "\Printers\Connections") $sid $UserName
                        }
                        finally {
                            Start-Process -FilePath "reg.exe" -ArgumentList "unload", $tempHive -WindowStyle Hidden -Wait -RedirectStandardOutput $null -RedirectStandardError $null -ErrorAction SilentlyContinue
                        }
                    }
                }
                $allConnections += $results
            }
            catch {
                Write-Warning "Failed to query user $UserName`: $($_.Exception.Message)"
            }
        }
        else {
            # Query all users
            $userProfiles = Get-CimInstance Win32_UserProfile | Where-Object { $_.SID -match "S-1-5-21" }
            foreach ($profile in $userProfiles) {
                try {
                    $userName = ([System.Security.Principal.SecurityIdentifier]$profile.SID).Translate([System.Security.Principal.NTAccount]).Value
                    
                    # Try loaded hive first
                    $loadedPath = "Registry::HKEY_USERS\$($profile.SID)\Printers\Connections"
                    $results = Get-ConnectionsFromHivePath $loadedPath $profile.SID $userName
                    
                    # If not loaded, try loading the profile hive temporarily
                    if (-not $results -and (Test-Path ($profile.LocalPath + "\NTUSER.DAT"))) {
                        $tempHive = "HKU\Temp_$($profile.SID -replace '[^A-Za-z0-9_]','_')"
                        try {
                            Start-Process -FilePath "reg.exe" -ArgumentList "load", $tempHive, ($profile.LocalPath + "\NTUSER.DAT") -WindowStyle Hidden -Wait -RedirectStandardOutput $null -RedirectStandardError $null -ErrorAction SilentlyContinue
                            $results = Get-ConnectionsFromHivePath ("Registry::" + $tempHive + "\Printers\Connections") $profile.SID $userName
                        }
                        finally {
                            Start-Process -FilePath "reg.exe" -ArgumentList "unload", $tempHive -WindowStyle Hidden -Wait -RedirectStandardOutput $null -RedirectStandardError $null -ErrorAction SilentlyContinue
                        }
                    }
                    $allConnections += $results
                }
                catch {
                    # Skip users we can't process
                    continue
                }
            }
        }
        
        return $allConnections
    }
    
    if ($ComputerName -eq $env:COMPUTERNAME) {
        return & $scriptBlock -UserName $UserName
    }
    else {
        return Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $UserName -ErrorAction SilentlyContinue
    }
}

function Get-AvailablePrinters {
    param(
        [string]$ComputerName,
        [string]$UserName = $null
    )
    
    try {
        Write-Host "Retrieving available printers on $ComputerName..." -ForegroundColor Cyan
        
        # Method 1: Get-Printer cmdlet (local printers)
        $getPrinters = @()
        try {
            if ($ComputerName -eq $env:COMPUTERNAME) {
                $getPrinters = Get-Printer -ErrorAction SilentlyContinue
            }
            else {
                $getPrinters = Get-Printer -ComputerName $ComputerName -ErrorAction SilentlyContinue
            }
        }
        catch {
            # Silently continue if Get-Printer fails
        }
        
        # Method 2: WMI/CIM query (additional network printers)
        $wmiPrinters = @()
        try {
            if ($ComputerName -eq $env:COMPUTERNAME) {
                $wmiPrinters = Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue
            }
            else {
                $wmiPrinters = Get-CimInstance -ClassName Win32_Printer -ComputerName $ComputerName -ErrorAction SilentlyContinue
            }
        }
        catch {
            # Silently continue if WMI fails
        }
        
        # Method 3: User-specific registry discovery
        $userConnections = @()
        try {
            $userConnections = Get-UserPrinterConnections -ComputerName $ComputerName -UserName $UserName
        }
        catch {
            Write-Warning "Failed to retrieve user printer connections: $($_.Exception.Message)"
        }
        
        # Combine and deduplicate results
        $allPrinters = @()
        $printerNames = @()
        
        # Add Get-Printer results
        foreach ($printer in $getPrinters) {
            if ($printer.Name -notin $printerNames) {
                $allPrinters += [PSCustomObject]@{
                    Name       = $printer.Name
                    Status     = if ($printer.PrinterStatus -eq "Normal") { "Online" } else { $printer.PrinterStatus }
                    Type       = $printer.Type
                    Location   = $printer.Location
                    DriverName = $printer.DriverName
                    Source     = "Local"
                    User       = ""
                }
                $printerNames += $printer.Name
            }
        }
        
        # Add WMI results
        foreach ($printer in $wmiPrinters) {
            if ($printer.Name -notin $printerNames) {
                $printerType = "Local"
                if ($printer.Network -eq $true) {
                    $printerType = "Network"
                }
                elseif ($printer.Shared -eq $true) {
                    $printerType = "Shared"
                }
                
                $status = "Unknown"
                if ($printer.PrinterStatus -ne $null) {
                    switch ($printer.PrinterStatus) {
                        1 { $status = "Other" }
                        2 { $status = "Unknown" }
                        3 { $status = "Idle" }
                        4 { $status = "Printing" }
                        5 { $status = "Warmup" }
                        6 { $status = "Stopped Printing" }
                        7 { $status = "Offline" }
                        default { $status = "Online" }
                    }
                }
                
                $allPrinters += [PSCustomObject]@{
                    Name       = $printer.Name
                    Status     = $status
                    Type       = $printerType
                    Location   = $printer.Location
                    DriverName = $printer.DriverName
                    Source     = "WMI"
                    User       = ""
                }
                $printerNames += $printer.Name
            }
        }
        
        # Add user-specific network printer connections
        foreach ($connection in $userConnections) {
            if ($connection.Connection -notin $printerNames) {
                $allPrinters += [PSCustomObject]@{
                    Name       = $connection.Connection
                    Status     = "User Connection"
                    Type       = "Network (User)"
                    Location   = "Print Server: $($connection.PrintServer)"
                    DriverName = ""
                    Source     = "User Registry"
                    User       = $connection.User
                }
                $printerNames += $connection.Connection
            }
        }
        
        if ($allPrinters.Count -gt 0) {
            Write-Host ""
            Write-Host "Available printers on ${ComputerName}:" -ForegroundColor Green
            Write-Host "------------------------------------------------------------"
            
            # Sort printers: Network printers first, then local
            $sortedPrinters = $allPrinters | Sort-Object @{Expression = { if ($_.Type -like "Network*") { 0 } else { 1 } } }, Name
            
            foreach ($printer in $sortedPrinters) {
                Write-Host "  Name: $($printer.Name)" -ForegroundColor White
                
                $statusColor = switch ($printer.Status) {
                    "Online" { "Green" }
                    "Idle" { "Green" }
                    "User Connection" { "Cyan" }
                    "Offline" { "Red" }
                    "Stopped Printing" { "Red" }
                    default { "Yellow" }
                }
                Write-Host "  Status: $($printer.Status)" -ForegroundColor $statusColor
                Write-Host "  Type: $($printer.Type)" -ForegroundColor Gray
                
                if ($printer.Location) {
                    Write-Host "  Location: $($printer.Location)" -ForegroundColor Gray
                }
                if ($printer.User) {
                    Write-Host "  User: $($printer.User)" -ForegroundColor Gray
                }
                if ($printer.DriverName) {
                    Write-Host "  Driver: $($printer.DriverName)" -ForegroundColor Gray
                }
                Write-Host ""
            }
            
            return $allPrinters
        }
        else {
            Write-Host "No printers found on $ComputerName" -ForegroundColor Yellow
            return $null
        }
    }
    catch {
        Write-Host "Error retrieving printers: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Show-InitialPrompt {
    Write-Host ""
    Write-Host "============================================================"
    Write-ColoredText -Text "This script lists all available printers on the specified" -ForegroundColor White
    Write-ColoredText -Text "computer, including local, network, and user-specific printers." -ForegroundColor White
    Write-ColoredText -Text "Multiple discovery methods are used for comprehensive results." -ForegroundColor Cyan
    Write-Host "============================================================"
    Write-Host ""
}

function Get-UserInput {
    param([string]$ComputerName)
    
    if (-not $ComputerName) {
        $computerName = (Read-Host "Enter target computer name (or press Enter for local computer)").Trim()
        if ([string]::IsNullOrWhiteSpace($computerName)) {
            $computerName = $env:COMPUTERNAME
        }
    }
    else {
        $computerName = $ComputerName.Trim()
    }
    
    # Ask about user scope for printer discovery
    Write-Host ""
    do {
        $userScope = Read-Host "Query printers for: (A)ll users or (S)pecific user? [A/S]"
        $userScope = $userScope.ToLower()
    } while ($userScope -ne "a" -and $userScope -ne "s")
    
    $userName = $null
    if ($userScope -eq "s") {
        $userName = Read-Host "Enter username (DOMAIN\username or COMPUTER\username format)"
    }
    
    return @{
        ComputerName = $computerName
        UserName     = $userName
    }
}

function Show-ContinuePrompt {
    param([string]$ComputerName)
    
    # Trim any whitespace from computer name for clean display
    $ComputerName = $ComputerName.Trim()
    
    Write-Host ""
    Write-Host "------------------------------------------------------------"
    Write-Host "Would you like to list printers from $ComputerName again or from a different computer?"
    
    do {
        $restart = Read-Host "List printers again Y/N?"
        $restart = $restart.ToLower()
    } while ($restart -ne "y" -and $restart -ne "n")
    
    return ($restart -eq "y")
}

function Show-PingFailedPrompt {
    param([string]$ComputerName)
    
    Write-Host ""
    Write-Host "------------------------------------------------------------"
    Write-ColoredText -Text "Computer $ComputerName does not ping." -ForegroundColor Red
    Write-Host "Would you like to try again?"
    
    do {
        $restart = Read-Host "Run script again Y/N?"
        $restart = $restart.ToLower()
    } while ($restart -ne "y" -and $restart -ne "n")
    
    return ($restart -eq "y")
}

# Main script execution
function Main {
    $Host.UI.RawUI.WindowTitle = "List Available Printers"
    
    $currentComputerName = $null
    
    do {
        Show-InitialPrompt
        $userInput = Get-UserInput -ComputerName $currentComputerName
        $computerName = $userInput.ComputerName
        $userName = $userInput.UserName
        $localHostName = $env:COMPUTERNAME
        
        # Store computer name for potential reuse (trim whitespace)
        if ([string]::IsNullOrWhiteSpace($computerName) -or $computerName -eq $null) {
            $currentComputerName = $null
        }
        else {
            try {
                $currentComputerName = $computerName.Trim()
            }
            catch {
                $currentComputerName = $null
            }
        }
        
        $continueScript = $true
        
        # Check if target is local computer
        if ($localHostName -eq $computerName) {
            $pingSuccess = $true
        }
        else {
            $pingSuccess = Test-ComputerConnectivity -ComputerName $computerName
        }
        
        if ($pingSuccess) {
            # List available printers
            $printers = Get-AvailablePrinters -ComputerName $computerName -UserName $userName
            
            if ($printers) {
                Write-Host ""
                Write-ColoredText -Text "Printer listing completed successfully!" -ForegroundColor Green
                Write-Host "Found $($printers.Count) printer(s) on $computerName"
            }
            else {
                Write-ColoredText -Text "No printers were found on $computerName" -ForegroundColor Yellow
            }
            
            Read-Host "Press Enter to continue"
            
            # Ask if user wants to continue
            $continueScript = Show-ContinuePrompt -ComputerName $computerName
        }
        else {
            # Ping failed
            $continueScript = Show-PingFailedPrompt -ComputerName $computerName
        }
        
    } while ($continueScript)
    
    Write-Host "Script completed."
}

# Execute main function
Main
