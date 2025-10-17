#Original from Ronal Arp - Ronald.Arp@va.gov
# Modified by Donnie Bartley - Donnie.Bartley@va.gov
# 03/24/2025 - Swapped YubiKey slots.  1 (short press) is now user & pass.  User & Pass is used more frequently.  Makes sense to save time with short press.
# 03/24/2025 - Set 10ms delay during user/pass send.  This helps to stop invalid login attempts over slow remote desktop connections.
# 08/25/2025 - Set 20ms delay during user/pass send.

# convert a password to yubikey modhex
function Convert-ToModHex {
    # parameters
    param (
        # string to convert to modhex
        [string]$InputString
    )

    # convert the string to a string of hex values
    $HexString = ($InputString.ToLower() | Format-Hex).HexBytes.Replace(' ', '')

    # our output modhex string
    $ModHexString = ""
    # loop through our hex string and convert the values to mod hex
    for ($i = 0; $i -lt $HexString.Length; $i++) {
        $ModHexString += (Get-ModHexCharacter -Char $HexString[$i])
    }

    # return our mod hex string
    return $ModHexString
}

# yubikey manager cli
$YubiCli = 'C:\Program Files\Yubico\YubiKey Manager CLI\ykman.exe'

if (-not (Test-Path -Path $YubiCli)) {
    Write-Output "YubiKey Manager CLI (ykman.exe) missing."
    return
}

# get our passwords to save to the yubikey
$Zero = Get-Credential -Message "Enter zero account username and password in NMEA@va.gov format"

# ask user if they want to send an enter key after yubikey password
$SendEnter = Read-Host "Do you want ENTER sent after your YubiKey is activated? (Y/N)"
# the arg to add/remove from the yubikey arguments
$EnterArgString = '--no-enter'
# check user's input for anything that is a Y or YES
if (($Null -ne $SendEnter) -and (($SendEnter.ToUpper() -eq "Y") -or ($SendEnter.ToUpper() -eq "YES"))) {
    $EnterArgString = $Null
    Write-Output "Enter will be sent after key activation." -ForegroundColor Cyan
}
else {
    Write-Output "Enter will NOT be sent after key activation." -ForegroundColor Cyan
}

# check if we got a 0 password from the user
if (($Null -ne $Zero) -and ($Zero.Length -gt 0)) {
    # get our mod hex version of our password
    $ZeroPass = ConvertFrom-SecureString $Zero.Password -AsPlainText
    $ZeroUser = "$($Zero.UserName)`t$(ConvertFrom-SecureString $Zero.Password -AsPlainText)"

    # send 0 account (NMEA) username and password to short press key
& $YubiCli otp static $EnterArgString --force --keyboard-layout US 1 $ZeroUser
    # send 0 account (NMEA) password only
& $YubiCli otp static $EnterArgString --force --keyboard-layout US 2 $ZeroPass
	# slow down key press time for both slots.  This is to stop failed password attempts over remote desktop sessions (RDP, SCCM Remote Client, BigFix, etc...)
& $YubiCli otp settings -f -p 20 1
& $YubiCli otp settings -f -p 20 2
    Write-Output "Done."
}
else {
    Write-Output "Credentials were null.`nYubiKey not updated."
}