<#
.SYNOPSIS
    Interactively updates the Description attribute for a single Active Directory computer,
    automatically selecting the appropriate domain/DC (root or child) based on the computer input.

.DESCRIPTION
    Prompts for a computer name and a new Description value, resolves the computer in AD across
    the relevant domain(s), shows current details, asks for confirmation, updates the Description,
    and re-reads the object from the same DC to verify the change.

    Domain selection logic:
      - If the input is an FQDN (contains dots), the domain suffix after the first dot is used
        to discover a writable DC (e.g., "SLC-WS120077.v19.med.va.gov" -> "v19.med.va.gov").
      - If the input is a short name, the script searches:
          1) The operator’s current domain
          2) The forest root domain
          3) Any additional domains in the forest

.PARAMETER None
    Interactive; prompts for:
      - Computer name (short or FQDN)
      - New Description
      - Confirmation

.INPUTS
    None.

.OUTPUTS
    Writes status to the console and, on success, displays Name and Description.

.EXAMPLE
    PS C:\> .\update_ad_computer_description.ps1

.NOTES
    ┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
    │ ORIGIN STORY                                                                                │ 
    ├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
    │   DATE        : 2025-09-18                                                                  │
    │   AUTHOR      : Dallas Bleak  (Dallas.Bleak@va.gov)                                         │
    │   VERSION     : 1.0                                                                         │
    │   Run As      : Elevated PowerShell (Run as Administrator) recommended.                     │
    │   Requirements:                                                                             │
    │     - Domain-joined Windows machine                                                         │
    │     - Rights to modify computer objects in the target OU/domain                             │
    │     - RSAT AD DS and LDS Tools (ActiveDirectory module)                                     │
    └─────────────────────────────────────────────────────────────────────────────────────────────┘
#>

[CmdletBinding()]
param()

# =========================
# Module Handling
# =========================
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Error "ActiveDirectory module could not be loaded. Ensure RSAT AD DS and LDS Tools are installed. Error: $($_.Exception.Message)"
    return
}

# =========================
# Helper Functions
# =========================

function Convert-LdapEscapedValue {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Value
    )
    # Replace backslash first
    $v = $Value -replace '\\', '\5c'
    $v = $v -replace '\*', '\2a'
    $v = $v -replace '\(', '\28'
    $v = $v -replace '\)', '\29'
    $v = $v -replace "`0", '\00'
    if ($v.StartsWith(' ')) { $v = '\20' + $v.Substring(1) }
    if ($v.EndsWith(' ')) { $v = $v.Substring(0, $v.Length - 1) + '\20' }
    return $v
}

function Get-WritableDC {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string] $DomainName
    )
    try {
        $dc = (Get-ADDomainController -DomainName $DomainName -Discover -Writable -ErrorAction Stop).HostName
        Write-Verbose "Discovered writable DC '$dc' for domain '$DomainName'."
        return $dc
    }
    catch {
        Write-Verbose "Failed to discover writable DC for domain '$DomainName': $($_.Exception.Message)"
        return $null
    }
}

function Resolve-ADComputerInDomain {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [string] $ComputerInput,
        [Parameter(Mandatory = $true)]
        [string] $DomainName
    )

    $server = Get-WritableDC -DomainName $DomainName
    if (-not $server) {
        Write-Verbose "Skipping domain '$DomainName' because no DC was found."
        return $null
    }

    try {
        $baseDN = (Get-ADDomain -Server $server -ErrorAction Stop).DistinguishedName
    }
    catch {
        Write-Verbose "Could not get base DN from '$server' in domain '$DomainName': $($_.Exception.Message)"
        return $null
    }

    # Attempt 1: sAMAccountName direct (ensure trailing $)
    $samCandidate = if ($ComputerInput.EndsWith('$')) { $ComputerInput } else { "$ComputerInput$" }
    try {
        $comp = Get-ADComputer -Identity $samCandidate -Server $server -Properties Description -ErrorAction Stop
        if ($comp) {
            Write-Verbose "Resolved by sAMAccountName in '$DomainName' via '$server'."
            return [pscustomobject]@{ Computer = $comp; Server = $server; DomainName = $DomainName }
        }
    }
    catch {
        Write-Verbose "Not found by sAMAccountName in '$DomainName': $($_.Exception.Message)"
    }

    # Attempt 2: LDAP filter across sAMAccountName, cn/name, dnsHostName
    $escapedShort = Convert-LdapEscapedValue -Value $ComputerInput
    $filters = @(
        "(sAMAccountName=$escapedShort)"
        "(cn=$escapedShort)"
        "(name=$escapedShort)"
    )

    if ($ComputerInput -like "*.*") {
        $escapedFqdn = Convert-LdapEscapedValue -Value $ComputerInput
        $filters += "(dnsHostName=$escapedFqdn)"
    }

    $ldap = "(|$($filters -join ''))"

    try {
        $comp = Get-ADComputer -LDAPFilter $ldap -Server $server -SearchBase $baseDN -SearchScope Subtree -Properties Description -ErrorAction Stop |
        Select-Object -First 1
        if ($comp) {
            Write-Verbose "Resolved by LDAP filter in '$DomainName' via '$server'."
            return [pscustomobject]@{ Computer = $comp; Server = $server; DomainName = $DomainName }
        }
        else {
            Write-Verbose "LDAP filter returned no results in '$DomainName' via '$server'."
        }
    }
    catch {
        Write-Verbose "LDAP search failed in '$DomainName' via '$server': $($_.Exception.Message)"
    }

    return $null
}

function Resolve-ADComputerAuto {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [string] $ComputerInput
    )

    # If input has dots, derive domain from suffix after first dot
    if ($ComputerInput -like "*.*") {
        $domainName = $ComputerInput.Substring($ComputerInput.IndexOf('.') + 1)
        Write-Verbose "Input appears to be FQDN; attempting domain '$domainName' first."
        $resolved = Resolve-ADComputerInDomain -ComputerInput $ComputerInput -DomainName $domainName
        if ($resolved) { return $resolved }
        Write-Verbose "Not found in '$domainName'; falling back to current and forest domains."
    }

    # Try current domain
    try {
        $currentDomain = (Get-ADDomain -ErrorAction Stop).DNSRoot
        Write-Verbose "Trying current domain '$currentDomain'."
        $resolved = Resolve-ADComputerInDomain -ComputerInput $ComputerInput -DomainName $currentDomain
        if ($resolved) { return $resolved }
    }
    catch {
        Write-Verbose "Could not determine current domain: $($_.Exception.Message)"
        $currentDomain = $null
    }

    # Try forest root
    $rootDomain = $null
    try {
        $rootDomain = (Get-ADForest -ErrorAction Stop).RootDomain
        if ($rootDomain) {
            Write-Verbose "Trying forest root domain '$rootDomain'."
            $resolved = Resolve-ADComputerInDomain -ComputerInput $ComputerInput -DomainName $rootDomain
            if ($resolved) { return $resolved }
        }
    }
    catch {
        Write-Verbose "Could not determine forest root domain: $($_.Exception.Message)"
    }

    # Try all other domains in the forest
    try {
        $domains = (Get-ADForest -ErrorAction Stop).Domains
        foreach ($d in $domains) {
            if ($d -ne $currentDomain -and $d -ne $rootDomain) {
                Write-Verbose "Trying additional forest domain '$d'."
                $resolved = Resolve-ADComputerInDomain -ComputerInput $ComputerInput -DomainName $d
                if ($resolved) { return $resolved }
            }
        }
    }
    catch {
        Write-Verbose "Could not enumerate forest domains: $($_.Exception.Message)"
    }

    return $null
}

# =========================
# Main Logic
# =========================

Write-Output "Update AD Computer Description (single item)" -ForegroundColor Cyan

# Prompt for computer identifier (short name or FQDN)
$compInput = Read-Host "Enter the computer name (short or FQDN, e.g., SLC-WS120077 or SLC-WS120077.v19.med.va.gov)"
if ([string]::IsNullOrWhiteSpace($compInput)) {
    Write-Error "Computer name cannot be empty."
    return
}

# Prompt for new Description
$descInput = Read-Host "Enter the new Description"
if ($null -eq $descInput) {
    Write-Error "Input aborted."
    return
}

# Resolve computer across domains
$resolved = Resolve-ADComputerAuto -ComputerInput $compInput

if (-not $resolved) {
    Write-Error "Computer '$compInput' not found in current, root, or other forest domains accessible from this session."
    return
}

$computer = $resolved.Computer
$server = $resolved.Server
$domain = $resolved.DomainName

Write-Output "Using DC: $server (Domain: $domain)" -ForegroundColor DarkCyan
Write-Output ""
Write-Output "Target computer found:" -ForegroundColor Yellow
$computer | Select-Object Name, SamAccountName, DNSHostName, DistinguishedName, Description | Format-List

# Confirm operation
$confirm = Read-Host "Set Description to: '$descInput' ? (Y/N)"
if ($confirm -notin @('Y', 'y', 'Yes', 'YES')) {
    Write-Output "Operation cancelled."
    return
}

# Update and verify
try {
    Set-ADComputer -Identity $computer.DistinguishedName -Server $server -Description $descInput -ErrorAction Stop
    Start-Sleep -Milliseconds 750
    $updated = Get-ADComputer -Identity $computer.DistinguishedName -Server $server -Properties Description

    Write-Output ""
    Write-Output "Success. Updated Description:" -ForegroundColor Green
    $updated | Select-Object Name, Description | Format-Table -AutoSize
}
catch {
    Write-Error "Failed to update Description on DC '$server' (Domain '$domain'): $($_.Exception.Message)"
}