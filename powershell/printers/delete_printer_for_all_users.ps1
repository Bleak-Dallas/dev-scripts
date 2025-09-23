#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Deletes a single printer from the default user profile for all users.

.DESCRIPTION
    This script deletes a specified local or network printer from the default account 
    for all existing/new users. Printer names with spaces will NOT be accepted.

.NOTES
    Name:      delete_printer_for_allU_uers.ps1
    Purpose:   Delete printer from default user profile
    Author:    Dallas Bleak
    Revision:  Converted from batch to PowerShell - December 2024
    ┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
    │ ORIGIN STORY                                                                                │ 
    ├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
    │   NAME:       : delete_printer_for_allU_uers.ps1                                            │
    │   DATE        : 2020-05-15                                                                  │
    │   AUTHOR      : Dallas Bleak (Dallas.Bleak@va.govs)                                         │
    │   VERSION     : 2.0                                                                         │
    │   REVISION    : 2025-09-23 Converted from batch to PowerShell                               │
    └─────────────────────────────────────────────────────────────────────────────────────────────┘

.EXAMPLE
    .\delete_printer_for_allU_uers.ps1
    Run the script and follow onscreen directions
#>

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

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

function Find-SimilarPrinters {
    param(
        [string]$SearchName,
        [array]$AvailablePrinters
    )
    
    if (-not $AvailablePrinters) { return $null }
    
    $suggestions = @()
    foreach ($printer in $AvailablePrinters) {
        # Check for partial matches or similar names
        if ($printer.Name -like "*$SearchName*" -or $SearchName -like "*$($printer.Name)*") {
            $suggestions += $printer.Name
        }
    }
    
    return $suggestions
}

function Remove-NetworkPrinterFromRegistry {
    param(
        [string]$ComputerName,
        [string]$PrinterName
    )
    
    try {
        Write-Host "Attempting to delete network printer from registry..." -ForegroundColor Yellow
        
        # Convert printer name to registry format (replace \ with ,,)
        $registryPrinterName = $PrinterName -replace "\\", ",,"
        
        $scriptBlock = {
            param($RegistryPrinterName)
            
            $deletedCount = 0
            
            # Get all user profiles
            $userProfiles = Get-CimInstance Win32_UserProfile | Where-Object { $_.SID -match "S-1-5-21" }
            
            foreach ($profile in $userProfiles) {
                try {
                    $userName = ([System.Security.Principal.SecurityIdentifier]$profile.SID).Translate([System.Security.Principal.NTAccount]).Value
                    
                    # Try loaded hive first
                    $loadedPath = "Registry::HKEY_USERS\$($profile.SID)\Printers\Connections\$RegistryPrinterName"
                    if (Test-Path $loadedPath) {
                        Remove-Item $loadedPath -Force -ErrorAction SilentlyContinue
                        Write-Host "  Deleted registry entry for user: $userName (loaded hive)" -ForegroundColor Green
                        $deletedCount++
                    }
                    
                    # Try unloaded hive if needed
                    if (Test-Path ($profile.LocalPath + "\NTUSER.DAT")) {
                        $tempHive = "HKU\Temp_$($profile.SID -replace '[^A-Za-z0-9_]','_')"
                        try {
                            Start-Process -FilePath "reg.exe" -ArgumentList "load", $tempHive, ($profile.LocalPath + "\NTUSER.DAT") -WindowStyle Hidden -Wait -RedirectStandardOutput $null -RedirectStandardError $null -ErrorAction SilentlyContinue
                            
                            $tempPath = "Registry::" + $tempHive + "\Printers\Connections\$RegistryPrinterName"
                            if (Test-Path $tempPath) {
                                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                                Write-Host "  Deleted registry entry for user: $userName (temp hive)" -ForegroundColor Green
                                $deletedCount++
                            }
                        }
                        finally {
                            Start-Process -FilePath "reg.exe" -ArgumentList "unload", $tempHive -WindowStyle Hidden -Wait -RedirectStandardOutput $null -RedirectStandardError $null -ErrorAction SilentlyContinue
                        }
                    }
                }
                catch {
                    # Skip users we can't process
                    continue
                }
            }
            
            return $deletedCount
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $deletedCount = & $scriptBlock -RegistryPrinterName $registryPrinterName
        }
        else {
            $deletedCount = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $registryPrinterName -ErrorAction SilentlyContinue
        }
        
        if ($deletedCount -gt 0) {
            Write-Host "Successfully deleted $deletedCount registry entries for printer: $PrinterName" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "No registry entries found to delete for printer: $PrinterName" -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        Write-Host "Error deleting from registry: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Remove-NetworkPrinterConnection {
    param(
        [string]$ComputerName,
        [string]$PrinterName
    )
    
    # Try multiple PrintUI methods with different parameters
    # Note: /ga parameter ensures deletion for all users
    $printUIMethods = @(
        @{ Args = "/gd /ga /q /c\\$ComputerName /n$PrinterName"; Description = "Global delete for all users with quiet mode" },
        @{ Args = "/gd /q /c\\$ComputerName /n$PrinterName"; Description = "Global delete with quiet mode" },
        @{ Args = "/dd /q /c\\$ComputerName /n$PrinterName"; Description = "Driver delete with quiet mode" },
        @{ Args = "/dn /q /c\\$ComputerName /n$PrinterName"; Description = "Name delete with quiet mode" },
        @{ Args = "/gd /c\\$ComputerName /n$PrinterName"; Description = "Global delete (original method)" }
    )
    
    foreach ($method in $printUIMethods) {
        try {
            Write-Host "Trying PrintUI method: $($method.Description)..." -ForegroundColor Yellow
            
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = "rundll32.exe"
            $processInfo.Arguments = "printui.dll,PrintUIEntry $($method.Args)"
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $true
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $process.Start() | Out-Null
            
            # Set a timeout to prevent hanging
            if ($process.WaitForExit(10000)) {
                # 10 second timeout
                if ($process.ExitCode -eq 0) {
                    Write-Host "Successfully deleted network printer using PrintUI: $PrinterName" -ForegroundColor Green
                    return $true
                }
                else {
                    Write-Host "PrintUI method failed with exit code: $($process.ExitCode)" -ForegroundColor Red
                }
            }
            else {
                Write-Host "PrintUI method timed out, killing process..." -ForegroundColor Red
                $process.Kill()
            }
        }
        catch {
            Write-Host "Error with PrintUI method: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "All PrintUI methods failed, trying registry deletion..." -ForegroundColor Yellow
    return Remove-NetworkPrinterFromRegistry -ComputerName $ComputerName -PrinterName $PrinterName
}

function Remove-PrinterForAllUsers {
    param(
        [string]$ComputerName,
        [string]$PrinterName
    )
    
    # Trim any whitespace from parameters
    $ComputerName = $ComputerName.Trim()
    $PrinterName = $PrinterName.Trim()
    
    try {
        Write-Host ""
        Write-Host "Attempting to delete $PrinterName for all users on $ComputerName"
        
        # Format printer name - handle different input formats
        if ($PrinterName -like "\\*") {
            # Already has UNC prefix, use as-is
            $fullPrinterName = $PrinterName
        }
        elseif ($PrinterName -like "*\*") {
            # Contains backslash but no UNC prefix, add it
            $fullPrinterName = "\\$PrinterName"
        }
        else {
            # Local printer name, use as-is
            $fullPrinterName = $PrinterName
        }
        
        # Determine if this is a network printer
        $isNetworkPrinter = $fullPrinterName -like "\\*"
        
        if ($isNetworkPrinter) {
            # For network printers, try PrintUI approach first (like original batch file)
            Write-Host "Detected network printer, using PrintUI deletion method..." -ForegroundColor Cyan
            $success = Remove-NetworkPrinterConnection -ComputerName $ComputerName -PrinterName $fullPrinterName
            if ($success) {
                return $true
            }
            else {
                Write-Host "PrintUI method failed, trying PowerShell Remove-Printer..." -ForegroundColor Yellow
            }
        }
        
        # Check if printer exists using Get-Printer
        $printer = $null
        if ($ComputerName -eq $env:COMPUTERNAME) {
            # Local computer
            $printer = Get-Printer -Name $fullPrinterName -ErrorAction SilentlyContinue
        }
        else {
            # Remote computer
            $printer = Get-Printer -ComputerName $ComputerName -Name $fullPrinterName -ErrorAction SilentlyContinue
        }
        
        if ($printer) {
            # Printer found, attempt deletion using Remove-Printer
            Write-Host "Printer found via Get-Printer, attempting PowerShell deletion..." -ForegroundColor Cyan
            if ($ComputerName -eq $env:COMPUTERNAME) {
                Remove-Printer -Name $fullPrinterName -Confirm:$false
            }
            else {
                Remove-Printer -ComputerName $ComputerName -Name $fullPrinterName -Confirm:$false
            }
            Write-Host "Successfully deleted printer: $fullPrinterName" -ForegroundColor Green
            return $true
        }
        else {
            # Printer not found via Get-Printer
            if ($isNetworkPrinter) {
                Write-Host "Network printer not found via Get-Printer (this is normal for user-installed network printers)" -ForegroundColor Yellow
                Write-Host "The PrintUI deletion method above should have handled the removal." -ForegroundColor Yellow
            }
            else {
                Write-Host "Printer '$fullPrinterName' not found on $ComputerName" -ForegroundColor Yellow
            }
            
            # Get available printers for suggestions
            $availablePrinters = Get-AvailablePrinters -ComputerName $ComputerName
            
            if ($availablePrinters) {
                # Look for similar printer names
                $suggestions = Find-SimilarPrinters -SearchName $PrinterName -AvailablePrinters $availablePrinters
                
                if ($suggestions) {
                    Write-Host ""
                    Write-Host "Did you mean one of these similar printers?" -ForegroundColor Cyan
                    foreach ($suggestion in $suggestions) {
                        Write-Host "  - $suggestion" -ForegroundColor White
                    }
                }
            }
            
            return $false
        }
    }
    catch [System.UnauthorizedAccessException] {
        Write-Host "Access denied: You don't have sufficient permissions to manage printers on $ComputerName" -ForegroundColor Red
        Write-Host "Try running as Administrator or check your domain permissions." -ForegroundColor Yellow
        return $false
    }
    catch [System.Management.Automation.Remoting.PSRemotingTransportException] {
        Write-Host "PowerShell remoting error: Cannot connect to $ComputerName" -ForegroundColor Red
        Write-Host "Check if PowerShell remoting is enabled and firewall allows connections." -ForegroundColor Yellow
        return $false
    }
    catch {
        Write-Host "Error deleting printer: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Error type: $($_.Exception.GetType().Name)" -ForegroundColor Gray
        
        # Provide specific guidance based on error type
        if ($_.Exception.Message -like "*Access*denied*") {
            Write-Host "This appears to be a permissions issue." -ForegroundColor Yellow
        }
        elseif ($_.Exception.Message -like "*network*" -or $_.Exception.Message -like "*RPC*") {
            Write-Host "This appears to be a network connectivity issue." -ForegroundColor Yellow
        }
        
        return $false
    }
}

function Restart-PrintSpooler {
    param([string]$ComputerName)
    
    try {
        # Trim any whitespace from computer name
        $ComputerName = $ComputerName.Trim()
        
        Write-Host "Stopping print spooler on $ComputerName..."
        if ($ComputerName -eq $env:COMPUTERNAME) {
            Stop-Service -Name "Spooler" -Force
            Start-Service -Name "Spooler"
        }
        else {
            # Use sc.exe command for remote computers with proper argument formatting
            $stopArgs = @("\\$ComputerName", "stop", "spooler")
            $startArgs = @("\\$ComputerName", "start", "spooler")
            
            $stopProcess = Start-Process -FilePath "sc.exe" -ArgumentList $stopArgs -Wait -NoNewWindow -PassThru
            if ($stopProcess.ExitCode -eq 0) {
                Write-Host "Spooler stopped successfully" -ForegroundColor Green
            }
            else {
                Write-Host "Warning: Spooler stop returned exit code $($stopProcess.ExitCode)" -ForegroundColor Yellow
            }
            
            Start-Sleep -Seconds 2  # Brief pause between stop and start
            
            $startProcess = Start-Process -FilePath "sc.exe" -ArgumentList $startArgs -Wait -NoNewWindow -PassThru
            if ($startProcess.ExitCode -eq 0) {
                Write-Host "Spooler started successfully" -ForegroundColor Green
            }
            else {
                Write-Host "Warning: Spooler start returned exit code $($startProcess.ExitCode)" -ForegroundColor Yellow
            }
        }
        
        Write-Host ""
        Write-Host "Print Spooler Service restarted" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error restarting print spooler: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Show-InitialPrompt {
    Write-Host ""
    Write-Host "============================================================"
    Write-ColoredText -Text "This script deletes the specified local or network printer" -ForegroundColor White
    Write-ColoredText -Text "from the default account for all existing/new users." -ForegroundColor White
    Write-ColoredText -Text "*IMPORTANT* Printer names with spaces will NOT be accepted." -ForegroundColor Yellow
    Write-ColoredText -Text "*IMPORTANT* If you get a pop up ERROR box close the script and try again." -ForegroundColor Yellow
    Write-ColoredText -Text "************************************************************" -ForegroundColor White
    Write-Host ""
}

function Get-UserInput {
    param([string]$ComputerName)
    
    if (-not $ComputerName) {
        $computerName = (Read-Host "Enter target computer name").Trim()
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
    
    # Ask if user wants to see available printers first
    Write-Host ""
    do {
        $showPrinters = Read-Host "Would you like to see available printers first? Y/N"
        $showPrinters = $showPrinters.ToLower()
    } while ($showPrinters -ne "y" -and $showPrinters -ne "n")
    
    if ($showPrinters -eq "y") {
        Get-AvailablePrinters -ComputerName $computerName -UserName $userName
    }
    
    $printerName = Read-Host "Enter Printserver/Printername (do not include \\)"
    
    return @{
        ComputerName = $computerName
        PrinterName  = $printerName
        UserName     = $userName
    }
}

function Show-SpoolerPrompt {
    Write-Host ""
    Write-Host "------------------------------------------------------------"
    Write-ColoredText -Text "Deleted printers will NOT disappear until spooler is restarted." -ForegroundColor Cyan -BackgroundColor Black
    Write-Host ""
    
    do {
        $reset = Read-Host "Reset print spooler Y/N?"
        $reset = $reset.ToLower()
    } while ($reset -ne "y" -and $reset -ne "n")
    
    return ($reset -eq "y")
}

function Show-ContinuePrompt {
    param([string]$ComputerName)
    
    # Trim any whitespace from computer name for clean display
    $ComputerName = $ComputerName.Trim()
    
    Write-Host ""
    Write-Host "------------------------------------------------------------"
    Write-Host "Would you like to delete another printer from $ComputerName or from a different computer?"
    
    do {
        $restart = Read-Host "Delete another printer Y/N?"
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
    # Check if running as administrator
    if (-not (Test-Administrator)) {
        Write-ColoredText -Text "REQUESTING ADMIN PRIVILEGES..." -ForegroundColor Yellow
        Write-Host "Please run this script as Administrator."
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    $Host.UI.RawUI.WindowTitle = "Delete Printer for All Users"
    
    $currentComputerName = $null
    
    do {
        Show-InitialPrompt
        $userInput = Get-UserInput -ComputerName $currentComputerName
        $computerName = $userInput.ComputerName
        $printerName = $userInput.PrinterName
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
            # Attempt to delete printer
            $deleteSuccess = Remove-PrinterForAllUsers -ComputerName $computerName -PrinterName $printerName
            
            if ($deleteSuccess) {
                # Ask about spooler restart
                $restartSpooler = Show-SpoolerPrompt
                
                if ($restartSpooler) {
                    Restart-PrintSpooler -ComputerName $computerName
                    Read-Host "Press Enter to continue"
                }
                
                # Ask if user wants to continue
                $continueScript = Show-ContinuePrompt -ComputerName $computerName
            }
            else {
                $continueScript = Show-ContinuePrompt -ComputerName $computerName
            }
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
