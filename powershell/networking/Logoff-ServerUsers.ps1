#Requires -Version 5.1
<#
.SYNOPSIS
Warns and logs off all interactive/RDP users on a server and optionally closes SMB sessions.

.DESCRIPTION
- Remote mode: runs quser/msg/logoff on the target via WinRM (avoids RPC to TermService).
- Local mode: runs directly on the current machine (no remoting).
- Gracefully treats 'No User exists for *' as no sessions and falls back to qwinsta if needed.
- Optionally closes SMB open files and sessions (CIM/WMI).
- Supports -WhatIf and -Confirm.

.EXAMPLES
# Dry run (remote)
.\Logoff-ServerUsers.ps1 -Server vhaslcprt1 -WhatIf -Verbose

# If DNS requires FQDN and HTTPS listener:
.\Logoff-ServerUsers.ps1 -Server vhaslcprt1.yourdomain.va.gov -UseSSL -Port 5986 -Verbose

# Using IP and NTLM (add to TrustedHosts once):
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "10.1.2.3" -Concatenate -Force
$cred = Get-Credential
.\Logoff-ServerUsers.ps1 -Server 10.1.2.3 -Authentication Negotiate -Credential $cred -Verbose

# Run locally on the server:
.\Logoff-ServerUsers.ps1 -Local -GraceMinutes 1 -CloseSmb -Verbose
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(ParameterSetName = 'Remote', Mandatory = $false)]
    [string]$Server = 'vhaslcprt1',

    [Parameter(Mandatory = $false)]
    [int]$GraceMinutes = 2,

    [Parameter(Mandatory = $false)]
    [switch]$SkipWarning,

    [Parameter(Mandatory = $false)]
    [switch]$CloseSmb,

    [Parameter(ParameterSetName = 'Remote', Mandatory = $false)]
    [pscredential]$Credential,

    [Parameter(ParameterSetName = 'Remote', Mandatory = $false)]
    [ValidateSet('Default', 'Kerberos', 'Negotiate', 'CredSSP', 'Basic')]
    [string]$Authentication = 'Kerberos',

    [Parameter(ParameterSetName = 'Remote', Mandatory = $false)]
    [switch]$UseSSL,

    [Parameter(ParameterSetName = 'Remote', Mandatory = $false)]
    [int]$Port,

    [Parameter(ParameterSetName = 'Local', Mandatory = $true)]
    [switch]$Local
)

# If they passed the local machine name, treat as local for reliability
if (-not $Local) {
    $isLocalName = $Server -and (
        $Server -ieq 'localhost' -or
        $Server -ieq $env:COMPUTERNAME -or
        $Server -ieq "$env:COMPUTERNAME.$($env:USERDNSDOMAIN)"
    )
    if ($isLocalName) { $Local = $true }
}

function Invoke-Remote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @(),
        [pscredential]$Credential,
        [string]$Authentication = 'Kerberos',
        [switch]$UseSSL,
        [int]$Port
    )
    $params = @{
        ComputerName   = $Server
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $ArgumentList
        ErrorAction    = 'Stop'
        Authentication = $Authentication
    }
    if ($PSBoundParameters.ContainsKey('Credential') -and $Credential) { $params.Credential = $Credential }
    if ($UseSSL) { $params.UseSSL = $true }
    if ($PSBoundParameters.ContainsKey('Port') -and $Port) { $params.Port = $Port }

    Invoke-Command @params
}

function ConvertFrom-Quser {
    [CmdletBinding()]
    param([string[]]$Lines)

    $regex = '^\s*>?\s*(?<User>\S+)\s+(?:(?<Session>\S+)\s+)?(?<Id>\d+)\s+(?<State>\S+)\s+(?<Idle>\S+)\s+(?<Logon>.+)$'
    $sessions = @()
    foreach ($line in $Lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -match 'USERNAME\s+SESSIONNAME\s+ID\s+STATE') { continue }
        $m = [regex]::Match($line, $regex)
        if ($m.Success) {
            $sessions += [pscustomobject]@{
                Username    = $m.Groups['User'].Value
                SessionName = $m.Groups['Session'].Value
                Id          = [int]$m.Groups['Id'].Value
                State       = $m.Groups['State'].Value
                IdleTime    = $m.Groups['Idle'].Value
                LogonTime   = $m.Groups['Logon'].Value
            }
        }
    }
    $sessions
}

function ConvertFrom-Qwinsta {
    [CmdletBinding()]
    param([string[]]$Lines)

    # Robust parser for qwinsta aligned output
    $states = @('Active', 'Disc', 'Listen', 'Down', 'Idle', 'Conn', 'Init', 'Reset')
    $sessions = @()

    foreach ($line in $Lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -match 'SESSIONNAME\s+USERNAME\s+ID\s+STATE') { continue }

        $norm = ($line -replace '^\s*>\s*', '' -replace '\s{2,}', ' ').Trim()
        if (-not $norm) { continue }
        $parts = $norm.Split(' ')
        # Find numeric ID token
        $idIndex = ($parts | ForEach-Object { $_ }) |
        ForEach-Object -Begin { $i = 0 } -Process {
            if ($_ -match '^\d+$') { $script:id = $i }
            $i++
        } | Out-Null; $idIndex = $script:id
        if ($null -eq $idIndex) { continue }

        $id = [int]$parts[$idIndex]
        $state = if ($parts.Count -gt $idIndex + 1) { $parts[$idIndex + 1] } else { '' }

        $candidateUser = if ($idIndex - 1 -ge 0) { $parts[$idIndex - 1] } else { '' }
        if ($candidateUser -and $states -notcontains $candidateUser) {
            $username = $candidateUser
            $sessionName = if ($idIndex - 2 -ge 0) { ($parts[0..($idIndex - 2)] -join ' ') } else { '' }
        }
        else {
            $username = ''
            $sessionName = if ($idIndex - 1 -ge 0) { ($parts[0..($idIndex - 1)] -join ' ') } else { '' }
        }

        $sessions += [pscustomobject]@{
            Username    = $username
            SessionName = $sessionName
            Id          = $id
            State       = $state
            IdleTime    = $null
            LogonTime   = $null
        }
    }
    $sessions
}

function Get-RdpSessionsLocal {
    [CmdletBinding()]
    param()
    Write-Verbose "Querying local user sessions via quser ..."
    try {
        $out = & quser.exe 2>&1
        $text = ($out -join "`n")
        if ($text -match 'No User exists for \*') {
            Write-Verbose "quser indicates no users exist locally."
            return @()
        }
        if ($LASTEXITCODE -eq 0) {
            $parsed = ConvertFrom-Quser -Lines $out
            if ($parsed.Count -gt 0) {
                return $parsed | Where-Object { $_.Username -and $_.Username -notin @('>', 'SERVICES', 'SYSTEM') }
            }
        }
        throw "quser parsing produced no sessions."
    }
    catch {
        Write-Verbose "quser failed locally; falling back to qwinsta ..."
        $q = & qwinsta.exe 2>&1
        $qtext = ($q -join "`n")
        if ($qtext -match 'No User exists for \*') {
            Write-Verbose "qwinsta indicates no users exist locally."
            return @()
        }
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to query sessions locally. Output: $qtext"
        }
        return ConvertFrom-Qwinsta -Lines $q | Where-Object { $_.Username -and $_.Username -notin @('>', 'SERVICES', 'SYSTEM') }
    }
}

function Get-RdpSessionsRemote {
    [CmdletBinding()]
    param([string]$Server, [pscredential]$Credential, [string]$Authentication, [switch]$UseSSL, [int]$Port)

    Write-Verbose "Querying user sessions on $Server via quser over WinRM ..."
    try {
        $out = Invoke-Remote -Server $Server -Credential $Credential -Authentication $Authentication -UseSSL:$UseSSL -Port $Port -ScriptBlock { quser.exe 2>&1 }
        $text = ($out -join "`n")
        if ($text -match 'No User exists for \*') {
            Write-Verbose "quser on $Server indicates no users exist."
            return @()
        }
        $parsed = ConvertFrom-Quser -Lines $out
        if ($parsed.Count -gt 0) {
            return $parsed | Where-Object { $_.Username -and $_.Username -notin @('>', 'SERVICES', 'SYSTEM') }
        }
        throw "quser parsing produced no sessions."
    }
    catch {
        Write-Verbose "quser failed on $Server; falling back to qwinsta ..."
        $q = Invoke-Remote -Server $Server -Credential $Credential -Authentication $Authentication -UseSSL:$UseSSL -Port $Port -ScriptBlock { qwinsta.exe 2>&1 }
        $qtext = ($q -join "`n")
        if ($qtext -match 'No User exists for \*') {
            Write-Verbose "qwinsta on $Server indicates no users exist."
            return @()
        }
        return ConvertFrom-Qwinsta -Lines $q | Where-Object { $_.Username -and $_.Username -notin @('>', 'SERVICES', 'SYSTEM') }
    }
}

function Send-UserWarningLocal {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([int]$Seconds, [array]$Sessions)
    $mins = [math]::Ceiling($Seconds / 60)
    $msg = "Admin notice: You will be logged off from $($env:COMPUTERNAME) in $mins minute(s) for maintenance. Please save your work."
    foreach ($s in $Sessions) {
        $target = "local session $($s.Id) ($($s.Username))"
        if ($PSCmdlet.ShouldProcess($target, "Warn via msg.exe")) {
            try { & msg.exe $s.Id /TIME:$Seconds $msg | Out-Null }
            catch { Write-Warning ("Failed to send message to session {0} on {1}: {2}" -f $s.Id, $env:COMPUTERNAME, $_.Exception.Message) }
        }
    }
}

function Send-UserWarningRemote {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([string]$Server, [int]$Seconds, [array]$Sessions, [pscredential]$Credential, [string]$Authentication, [switch]$UseSSL, [int]$Port)
    $mins = [math]::Ceiling($Seconds / 60)
    $msg = "Admin notice: You will be logged off from $Server in $mins minute(s) for maintenance. Please save your work."
    foreach ($s in $Sessions) {
        $target = "$Server session $($s.Id) ($($s.Username))"
        if ($PSCmdlet.ShouldProcess($target, "Warn via msg.exe")) {
            try {
                Invoke-Remote -Server $Server -Credential $Credential -Authentication $Authentication -UseSSL:$UseSSL -Port $Port -ScriptBlock {
                    param($id, $sec, $message)
                    msg.exe $id /TIME:$sec $message | Out-Null
                } -ArgumentList @($s.Id, $Seconds, $msg)
            }
            catch {
                Write-Warning ("Failed to send message to session {0} on {1}: {2}" -f $s.Id, $Server, $_.Exception.Message)
            }
        }
    }
}

function Disconnect-UserSessionsLocal {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([array]$Sessions)
    foreach ($s in $Sessions) {
        $target = "local session $($s.Id) ($($s.Username))"
        if ($PSCmdlet.ShouldProcess($target, "Logoff")) {
            try { & logoff.exe $s.Id }
            catch { Write-Warning ("Failed to log off session {0} on {1}: {2}" -f $s.Id, $env:COMPUTERNAME, $_.Exception.Message) }
        }
    }
}

function Disconnect-UserSessionsRemote {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([string]$Server, [array]$Sessions, [pscredential]$Credential, [string]$Authentication, [switch]$UseSSL, [int]$Port)
    foreach ($s in $Sessions) {
        $target = "$Server session $($s.Id) ($($s.Username))"
        if ($PSCmdlet.ShouldProcess($target, "Logoff")) {
            try {
                Invoke-Remote -Server $Server -Credential $Credential -Authentication $Authentication -UseSSL:$UseSSL -Port $Port -ScriptBlock {
                    param($id)
                    logoff.exe $id
                } -ArgumentList @($s.Id)
            }
            catch {
                Write-Warning ("Failed to log off session {0} on {1}: {2}" -f $s.Id, $Server, $_.Exception.Message)
            }
        }
    }
}

function Close-SmbSessions {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([string]$TargetServer, [pscredential]$Credential)

    Write-Verbose "Closing SMB open files and sessions on $TargetServer via CIM ..."
    try {
        $cim = if ($Credential) {
            New-CimSession -ComputerName $TargetServer -Credential $Credential -ErrorAction Stop
        }
        else {
            New-CimSession -ComputerName $TargetServer -ErrorAction Stop
        }

        $openFiles = Get-SmbOpenFile -CimSession $cim -ErrorAction SilentlyContinue
        if ($openFiles) {
            foreach ($f in $openFiles) {
                $desc = "$TargetServer open file ID $($f.FileId)/$($f.SessionId) ($($f.Path))"
                if ($PSCmdlet.ShouldProcess($desc, "Close-SmbOpenFile -Force")) {
                    try { Close-SmbOpenFile -CimSession $cim -FileId $f.FileId -Force -ErrorAction Stop }
                    catch { Write-Warning ("Failed to close open file {0} (Session {1}): {2}" -f $f.Path, $f.SessionId, $_.Exception.Message) }
                }
            }
        }

        $smbSessions = Get-SmbSession -CimSession $cim -ErrorAction SilentlyContinue
        if ($smbSessions) {
            foreach ($sess in $smbSessions) {
                $desc = "$TargetServer SMB session $($sess.SessionId) ($($sess.ClientUserName) from $($sess.ClientComputerName))"
                if ($PSCmdlet.ShouldProcess($desc, "Close-SmbSession -Force")) {
                    try { Close-SmbSession -CimSession $cim -SessionId $sess.SessionId -Force -ErrorAction Stop }
                    catch { Write-Warning ("Failed to close SMB session {0}: {1}" -f $sess.SessionId, $_.Exception.Message) }
                }
            }
        }
        else {
            Write-Verbose "No SMB sessions found on $TargetServer."
        }
    }
    catch {
        Write-Warning "Could not manage SMB sessions on $TargetServer via CIM/WMI. Error: $($_.Exception.Message)"
    }
    finally {
        if ($cim) { $cim | Remove-CimSession -ErrorAction SilentlyContinue }
    }
}

# -------------------------
# Main
# -------------------------
try {
    $targetName = if ($Local) { $env:COMPUTERNAME } else { $Server }

    if ($Local) {
        $sessions = Get-RdpSessionsLocal
    }
    else {
        # Clearer error if WinRM/DNS has issues
        try {
            $twParams = @{ ComputerName = $Server; ErrorAction = 'Stop' }
            if ($UseSSL) { $twParams.UseSSL = $true }
            if ($PSBoundParameters.ContainsKey('Port') -and $Port) { $twParams.Port = $Port }
            Test-WSMan @twParams | Out-Null
        }
        catch {
            throw ("Unable to reach {0} over WinRM. If DNS fails, try the short name or the IP with -Authentication Negotiate (and add to TrustedHosts). Details: {1}" -f $Server, $_.Exception.Message)
        }
        $sessions = Get-RdpSessionsRemote -Server $Server -Credential $Credential -Authentication $Authentication -UseSSL:$UseSSL -Port $Port
    }

    if (-not $sessions -or $sessions.Count -eq 0) {
        Write-Host ("No interactive/RDP user sessions found on {0}." -f $targetName)
        return
    }

    Write-Host ("Found {0} session(s) on {1}:" -f $sessions.Count, $targetName)
    $sessions | Select-Object Username, SessionName, Id, State, IdleTime, LogonTime | Format-Table -AutoSize

    if (-not $SkipWarning -and $GraceMinutes -gt 0) {
        if ($Local) {
            Send-UserWarningLocal -Seconds ($GraceMinutes * 60) -Sessions $sessions
        }
        else {
            Send-UserWarningRemote -Server $Server -Seconds ($GraceMinutes * 60) -Sessions $sessions -Credential $Credential -Authentication $Authentication -UseSSL:$UseSSL -Port $Port
        }
        if ($PSBoundParameters.ContainsKey('WhatIf')) {
            Write-Verbose "WhatIf specified - skipping wait."
        }
        else {
            Write-Verbose ("Waiting {0} minute(s) before logging off users ..." -f $GraceMinutes)
            Start-Sleep -Seconds ($GraceMinutes * 60)
        }
    }

    if ($Local) {
        Disconnect-UserSessionsLocal -Sessions $sessions
    }
    else {
        Disconnect-UserSessionsRemote -Server $Server -Sessions $sessions -Credential $Credential -Authentication $Authentication -UseSSL:$UseSSL -Port $Port
    }

    if ($CloseSmb) {
        Close-SmbSessions -TargetServer $targetName -Credential $Credential
    }
}
catch {
    Write-Error $_
}