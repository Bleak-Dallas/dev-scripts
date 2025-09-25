@ECHO off
:: Name:  addPrinterForAllUsers.bat
:: Purpose:  This script adds a single printer to the default user profile.
::           NOTE:  Printer names with spaces will NOT be accepted.  
::           Usage: run addPrinterForAllUsers.bat and follow onscreen directions
:: Author:   Dallas Bleak
:: Revision:  May 2020 - initial version



Title Get Admin Rights
:: BatchGotAdmin (Run as Admin code starts)
REM --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    ECHO.
    COLOR 0E
    ECHO REQUESTING ADMIN PRIVILEGES...
goto UACPrompt
) else ( goto gotAdmin )
:UACPrompt
ECHO Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
ECHO UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
"%temp%\getadmin.vbs"
exit /B
:gotAdmin
if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
pushd "%CD%"
CD /D "%~dp0"
:: BatchGotAdmin (Run as Admin code ends)
:: Your codes should start from the following line
GOTO INITIATE

:INITIATE
cls
ECHO.
COLOR 0F
ECHO This script adds the specified local or network printer  
ECHO to the deafult account for all existing/new users.
ECHO *IMPORTATNT* Printer names with spaces will NOT be accepted.
ECHO *IMPORTATNT* If you get a pop up ERROR box close script and try again.
ECHO ************************************************************

SET /P computerName=Enter target computer name: 
SET /P printerName=Enter Printserver/Printername (do not include \\):
FOR /F "usebackq" %%i IN (`hostname`) DO SET localHostName=%%i
    IF %localHostName%==%computerName% (GOTO ADDPRINTER
    ) ELSE ( GOTO START_PING )

:START_PING
ECHO.
ECHO Attempting to ping %computerName%
CD C:\Windows\System32
PING %computerName% -n 3
    IF ERRORLEVEL 1 ( GOTO PING_FAILED
    ) ELSE ( GOTO ADDPRINTER )

:ADDPRINTER
CLS
rundll32 printui.dll,PrintUIEntry /ga /c \\%computerName% /n \\%printerName%
ECHO.
ECHO Attempting to add %printerName% for all users on %computerName%

:START_SPOOLY
CLS
COLOR 1F
ECHO.
ECHO New printers will NOT appear until spooler is restarted.
SET /P reset=Reset print spooler Y/N? 
    IF %reset%==y ( GOTO SPOOLY 
    ) ELSE ( GOTO CHOICE )

:SPOOLY
REM stop the print spooler on the specified computer and wait until the sc command finishes
start /wait sc \\%computerName% stop spooler
REM start the print spooler on the specified computer and wait until the sc command finishes
start /wait sc \\%computerName% start spooler
ECHO.
ECHO Print Spooler Service restarted
PAUSE
GOTO CHOICE

:CHOICE
CLS
COLOR 0F
ECHO.
ECHO Would you like to add another printer to %computerName% or from a differnt computer?
SET /P restart=Add another printer Y/N?
if %restart%==y ( GOTO INITIATE 
) ELSE ( GOTO END )

:PING_FAILED
CLS
COLOR 04
ECHO.
ECHO Computer %computerName% does not ping.
ECHO Would you like to try again?
SET /P restart=Run script again Y/N? Y/N?
if %restart%==y ( GOTO INITIATE 
) ELSE ( GOTO END )

:END
