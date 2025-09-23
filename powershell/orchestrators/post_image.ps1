<# 
.SYNOPSIS 
Completes post imaging tasks

.DESCRIPTION 
Post imaging script that copy's common dll's to proper folder
Adds group members to the Remote Desktop Users Group
Prevents the laptop from sleeping when closing the laptop lid
Adds a description to the AD description field

.NOTES 
┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ ORIGIN STORY                                                                                │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 2021.08.26
│   AUTHOR      : Dallas Bleak
│   DESCRIPTION : Initial Draft 
    VERSION     : 1.1 - Changed folder destinations
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 

.PARAMETER Param01 
$computerName - Takes the name of the computer supplied by the cmd script
$global:AD_Description - Takes the description that will go into the AD desription field
#> 

### GLOBAL VARIABLES ###
$global:isDesktop = 1
$global:isADDescription = 0
$global:AD_Description = $null
function Install-Single {
    #====== BEGIN - Install Apllications on Single Computer =====#
    $computerName = READ-HOST "Enter Device for Install"
    get-service -Name winrm -ComputerName $computerName | start-service
    Clear-Host
    Test-Connectivity ($computerName)
    Get-ADDescriptionChoice ($computerName)
    Clear-Host
    Copy-DLL($computerName)
    Clear-Host
    Add-GroupMembers ($computerName)
    Clear-Host
    if (1 -eq $global:isADDescription) {
        Write-Host "ADDING AD DESCRIPTION "-NoNewline; Write-Host $global:AD_Description -ForegroundColor Green -NoNewline; Write-Host " TO " -NoNewline; Write-Host $computerName -ForegroundColor Blue -NoNewline
        Set-ADComputer $computerName -Description $global:AD_Description
    }
    Clear-Host
    if (0 -eq $global:isDesktop) {
        Optimize-Power ($computerName)
    }
    get-service -Name winrm -ComputerName $computerName | stop-service
    Clear-Host
    Get-ChoiceLaptopDesktop
    #====== END - Install Apllications on Single Computer =====#
}

function Install-MultipleDifferentAD {
        #====== BEGIN - Create or open file for multiple computer names =====#
        # Full path of the file
        $file = "\\va.gov\cn\Salt Lake City\VHASLC\TechDrive\Scripts\fileNames\$env:UserName.txt"

        #If the file does not exist, create it and open the file to be edited.
        if (-not(Test-Path -Path $file -PathType Leaf)) {
            try {
                $null = New-Item -ItemType File -Path $file -Force -ErrorAction Stop
                Invoke-Item "\\va.gov\cn\Salt Lake City\VHASLC\TechDrive\Scripts\fileNames\$env:UserName.txt"
                Read-Host -Prompt "Press any key to continue"
            }
            catch {
                throw $_.Exception.Message
            }
        }
        # If the file already exists open the file to be edited.
        else {
            Invoke-Item "\\va.gov\cn\Salt Lake City\VHASLC\TechDrive\Scripts\fileNames\$env:UserName.txt"
            Read-Host -Prompt "Press any key to continue"
            $computerNameNames = Get-Content -Path $file
            #Write-Output $computerNameNames
        }
        #====== END - Create or open file for multiple computer names =====#

    #====== BEGIN - Install Apllications on Multiple Computers =====#
    foreach($computerName in $computerNameNames) {
        Test-Connectivity ($computerName)
        get-service -Name winrm -ComputerName $computerName | start-service
        Get-ADDescriptionChoice ($computerName)
        Clear-Host
        Copy-DLL($computerName)
        Clear-Host
        Add-GroupMembers ($computerName)
        Clear-Host
        if (1 -eq $global:isADDescription) {
            Write-Host "ADDING AD DESCRIPTION "-NoNewline; Write-Host $global:AD_Description -ForegroundColor Green -NoNewline; Write-Host " TO " -NoNewline; Write-Host $computerName -ForegroundColor Blue -NoNewline
            Set-ADComputer $computerName -Description $global:AD_Description
        }
        Clear-Host
        if (0 -eq $global:isDesktop) {
            Optimize-Power ($computerName)
        }
        Clear-Host
        get-service -Name winrm -ComputerName $computerName | start-service
    }
    Get-ChoiceLaptopDesktop
    #====== END - Install Apllications on Multiple Computers =====#
}

function Install-MultipleSameAD {
    #====== BEGIN - Create or open file for multiple computer names =====#
    # Full path of the file
    $file = "\\va.gov\cn\Salt Lake City\VHASLC\TechDrive\Scripts\fileNames\$env:UserName.txt"

    #If the file does not exist, create it and open the file to be edited.
    if (-not(Test-Path -Path $file -PathType Leaf)) {
        try {
            $null = New-Item -ItemType File -Path $file -Force -ErrorAction Stop
            Invoke-Item "\\va.gov\cn\Salt Lake City\VHASLC\TechDrive\Scripts\fileNames\$env:UserName.txt"
            Read-Host -Prompt "Press any key to continue"
        }
        catch {
            throw $_.Exception.Message
        }
    }
    # If the file already exists open the file to be edited.
    else {
        Invoke-Item "\\va.gov\cn\Salt Lake City\VHASLC\TechDrive\Scripts\fileNames\$env:UserName.txt"
        Read-Host -Prompt "Press any key to continue"
        $computerNameNames = Get-Content -Path $file
        #Write-Output $computerNameNames
    }
    #====== END - Create or open file for multiple computer names =====#

#====== BEGIN - Install Apllications on Multiple Computers SAME=====#
$global:AD_Description = Read-Host "PLEASE ENTER A DESCRIPTION FOR ACTIVE DIRECTORY"
foreach($computerName in $computerNameNames) {
    Test-Connectivity ($computerName)
    get-service -Name winrm -ComputerName $computerName | start-service
    Copy-DLL($computerName)
    Clear-Host
    Add-GroupMembers ($computerName)
    Clear-Host
    Write-Host "ADDING AD DESCRIPTION "-NoNewline; Write-Host $global:AD_Description -ForegroundColor Green -NoNewline; Write-Host " TO " -NoNewline; Write-Host $computerName -ForegroundColor Blue -NoNewline
    Set-ADComputer $computerName -Description $global:AD_Description
    Clear-Host
    if (0 -eq $global:isDesktop) {
        Optimize-Power ($computerName)
    }
    Clear-Host
    get-service -Name winrm -ComputerName $computerName | start-service
}
Get-ChoiceLaptopDesktop
#====== END - Install Apllications on Multiple Computers SAME=====#
}

function Test-Connectivity ($computerName) {
    if (test-connection -ComputerName $computerName -Count 1 -quiet -buffer 8){
        write-host "$computerName ONLINE" -ForegroundColor Green
    } else {
        write-host "$computerName OFFLINE" -ForegroundColor Red
        Get-Choice
    }
}

function Copy-DLL($computerName) {
    write-host "INSTALLING COMMON DLL's TO " -ForegroundColor Yellow -NoNewline; Write-Host $computerName -ForegroundColor Blue -NoNewline
    robocopy  "\\va.gov\cn\Salt Lake City\VHASLC\TechDrive\CPRS_Update_Script\Common Files"  "\\$computerName.v19.med.va.gov\c$\Program Files (x86)\VistA\Common Files" /E /R:1 /W:1 /TEE #using robocopy to see copy progress
    Start-Sleep -Seconds 3
}

function Add-GroupMembers ($computerName) {
    write-host "ADDING 'WORKSTAIONADMINS' AND 'VHASLCOI&T' TO REMOTE DESKTOP USERS GROUP"
    try {
        Invoke-Command -ComputerName $computerName -ErrorAction Stop -ScriptBlock {Add-LocalGroupMember -Group "Remote Desktop Users" -Member "VHASLCOI&T"}
    }
    catch {
        $theError = $_
        if ($theError.Exception.Message -like '*VHA19\vhaslcoi&t*already*') {
            Write-Warning 'VHASLCOI&T is already a member of the Remote Desktop Users Group'
            Start-Sleep -Seconds 2
        } 
        elseif ($theError.Exception.Message -like '*failed*') {
            Write-Host 'Cannot connect to the computer. Make sure computer is in the right Active Directory OU.' -ForegroundColor Red
            Start-Sleep -Seconds 5
        }
    }
    
    try {
        Invoke-Command -ComputerName $computerName -ErrorAction Stop -ScriptBlock {Add-LocalGroupMember -Group "Remote Desktop Users" -Member "VHASLCWorkstationAdmins"}
    }
    catch {
        $theError = $_
        if ($theError.Exception.Message -like '*VHA19\vhaslcworkstationadmins*already*') {
            Write-Warning 'VHASLCWorkstationAdmins is already a member of the Remote Desktop Users Group'
        Start-Sleep -Seconds 2
        } 
        elseif ($theError.Exception.Message -Like '*failed*') {
            Write-Host 'Cannot connect to the computer. Make sure computer is in the right Active Directory OU.' -ForegroundColor Red
            Start-Sleep -Seconds 5
        }
    }
    
    finally {
        Write-Host "`n"
        Write-Host "Current members for the Remote Desktop Users Group " -ForegroundColor Green
        Invoke-Command -ComputerName $computerName -ScriptBlock {Get-LocalGroupmember -Group "Remote Desktop Users"}
        Start-Sleep -Seconds 2
    }
}

<# function Optimize-Power ($computerName){
    $script:psexec = "C:\Windows\System32\PsExec.exe"
    $command1 = 'cmd /c powercfg /setacvalueindex scheme_current sub_buttons lidaction 0' # Set lid to do nothing when on plugged in.
    $command2 = 'cmd /c powercfg /setdcvalueindex scheme_current sub_buttons lidaction 0' # Set lid to do nothing when on battery in.
    $command3 = 'cmd /c powercfg /setactive scheme_current' # Apply changes
    Write-Host "APPLYING LAPTOP LID POWER SETTINGS " -NoNewline; Write-Host $computerName -ForegroundColor Blue -NoNewline
    Start-Process -FilePath $psexec -ArgumentList "\\$computerName $command1" -Wait
    Start-Process -FilePath $psexec -ArgumentList "\\$computerName $command2" -Wait
    Start-Process -FilePath $psexec -ArgumentList "\\$computerName $command3" -Wait
} #>

function Optimize-Power ($computerName) {
    # Create a scriptblock with all the power configuration commands
    $scriptBlock = {
        powercfg /setacvalueindex scheme_current sub_buttons lidaction 0
        powercfg /setdcvalueindex scheme_current sub_buttons lidaction 0
        powercfg /setactive scheme_current
    }

    Write-Host "APPLYING LAPTOP LID POWER SETTINGS " -NoNewline
    Write-Host $computerName -ForegroundColor Blue -NoNewline
    
    # Execute the commands on the remote computer using Invoke-Command
    Invoke-Command -ComputerName $computerName -ScriptBlock $scriptBlock
}
    
function Add-ADDescription ($computerName){
    $global:AD_Description = Read-Host "PLEASE ENTER A DESCRIPTION FOR ACTIVE DIRECTORY"
    Set-ADComputer $computerName -Description $global:AD_Description
}


function Get-ADDescriptionChoice ($computerName){
    #====== BEGIN - Get choice from User for install =====#
    Write-Host "DO YOU WANT TO ADD AN AD DESCRIPTION TO $computerName" -ForegroundColor Yellow
    Write-Output "Press 1 to add an AD Descripion"
    Write-Output "Press 2 to decline"
    $choice = Read-Host -Prompt "1 or 2"
    if (1 -eq $choice -or 2 -eq $choice) {
        switch ($choice ) {
            1 { $global:isADDescription = 1
                $global:AD_Description = Read-Host "PLEASE ENTER A DESCRIPTION FOR ACTIVE DIRECTORY"}
            2 { $global:isADDescription}
            }
    } else {
        Write-Host "Invalid choice" -ForegroundColor Red
        Start-Sleep -Seconds 1
        Clear-Host
        Get-ADDescriptionChoice ($computerName)
    }
    #====== END - Get choice from User for install =====#
}

function Get-ChoiceLaptopDesktop {
    #====== BEGIN - Get choice from User for install =====#
    Write-Host "Press " -NoNewline;
    Write-Host "1 " -ForegroundColor Green -NoNewline;
    Write-Host "to run Postimage on " -NoNewline;
    Write-Host "Desktop" -ForegroundColor Green;
    Write-Host "Press " -NoNewline;
    Write-Host "2 " -ForegroundColor Yellow -NoNewline;
    Write-Host "to run Postimage on " -NoNewline;
    Write-Host "Laptop" -ForegroundColor Yellow
    $choice = Read-Host -Prompt "1 or 2"
    if (1 -eq $choice -or 2 -eq $choice) {
        switch ($choice ) {
            1 { $global:isDesktop
                Clear-Host
                Get-Choice}
            2 { $global:isDesktop = 0
                Clear-Host
                Get-Choice}
            }
    } else {
        Write-Host "Invalid choice" -ForegroundColor Red
        Start-Sleep -Seconds 1
        Clear-Host
        Get-ChoiceLaptopDesktop
    }
    #====== END - Get choice from User for install =====#
}

function Get-Choice {
    #====== BEGIN - Get choice from User for install =====#
    Write-Output "Press 1 to install on a single computer"
    Write-Output "Press 2 to install on a multiple computers with different AD descriptions (leave desription blank to use no description)"
    Write-Output "Press 3 to install on a multiple computers with same AD description"
    $choice = Read-Host -Prompt "1 or 2"
    if (1 -eq $choice -or 2 -eq $choice -or 3 -eq $choice) {
        switch ($choice ) {
            1 {Install-Single}
            2 {Install-MultipleDifferentAD}
            3 {Install-MultipleSameAD}
            }
    } else {
        Write-Host "Invalid choice" -ForegroundColor Red
        Start-Sleep -Seconds 1
        Clear-Host
        Get-Choice
    }
    #====== END - Get choice from User for install =====#
}

Get-ChoiceLaptopDesktop ##### START PROGRAM ####
