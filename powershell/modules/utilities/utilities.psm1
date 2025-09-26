function Invoke-Elevation {
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.Collections.IDictionary] $BoundParameters
    )

    if (-not $BoundParameters) {
        $BoundParameters = @{}
    }

    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($currentIdentity)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host 'Elevation required. Relaunching with administrative privileges...' -ForegroundColor Yellow

        # Relaunch with the same host (pwsh or powershell)
        $processPath = (Get-Process -Id $PID).Path

        $argumentList = @(
            '-NoProfile'
            '-ExecutionPolicy', 'Bypass'
            '-File', "`"$PSCommandPath`""
        )

        foreach ($entry in $BoundParameters.GetEnumerator()) {
            $key = "-$($entry.Key)"
            $value = $entry.Value

            if ($null -eq $value) { continue }

            # Switch parameters
            if ($value -is [System.Management.Automation.SwitchParameter]) {
                if ($value.IsPresent) { $argumentList += $key }
                continue
            }

            # Array parameters -> repeat key for each value
            if ($value -is [System.Array]) {
                foreach ($item in $value) {
                    if ($null -ne $item -and $item -ne '') {
                        $argumentList += $key
                        $argumentList += "`"$item`""
                    }
                }
                continue
            }

            # Scalar values
            $stringValue = [string]$value
            if ($stringValue.Length -eq 0) { continue }
            $argumentList += $key
            $argumentList += "`"$stringValue`""
        }

        Start-Process -FilePath $processPath -ArgumentList $argumentList -Verb RunAs | Out-Null
        exit
    }
}