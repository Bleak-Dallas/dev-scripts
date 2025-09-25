:: Name:     power.bat
:: Purpose:  Gets battery status from a remote laptop computer       
:: Author:   Dallas Bleak
:: Revision: November 2020 - initial version

@echo off
Title Laptop Startup
:: BatchGotAdmin (Run as Admin code starts)
REM --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
echo Requesting administrative privileges...
goto UACPrompt
) else ( goto gotAdmin )
:UACPrompt
echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
"%temp%\getadmin.vbs"
exit /B
:gotAdmin
if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
pushd "%CD%"
CD /D "%~dp0"
:: BatchGotAdmin (Run as Admin code ends)
:: Your codes should start from the following line
======================================================================================
Color 0F
======================================================================================
ECHO.
ECHO.
ECHO [96m********************************************[0m
ECHO [96m******* Get Battery Status ******[0m
ECHO [96m********************************************[0m
ECHO[
ECHO[
ECHO[
ECHO[
REM **********************************************************************************
GOTO STARTER

REM ####### GET COMPUTER NAME AND CONNECT TO COMPUTER ######
:STARTER
Title Get Battery Status
CLS
ECHO.
ECHO WHAT IS THE COMPUTER NAME?
SET COMPUTERNAME=
SET /P COMPUTERNAME=COMPUTER NAME: 
IF NOT DEFINED COMPUTERNAME (
    ECHO PLEASE ENTER A COMPUTER NAME
    TIMEOUT /T 3 >NUL
    GOTO STARTER
) ELSE (
    GOTO GET_BATTERY_STATUS
)

REM ####### GET BATTERY STATUS ######
:GET_BATTERY_STATUS
psexec \\%COMPUTERNAME% -e cmd /c ( WMIC Path Win32_Battery Get BatteryStatus )
ECHO.
ECHO.
ECHO [96m********************************************[0m
echo [96m DISCHARGING (ON BATTERY POWER) [0m
echo [96m     Other (1)                  [0m
echo [96m     Low (4)                    [0m
echo [96m     Critical (5)               [0m
echo [96m CHARGING (ON AC POWER)         [0m
echo [96m     Unknown (2)                [0m
echo [96m     Charging (6)               [0m
echo [96m     Charging and High (7)      [0m
echo [96m     Charging and Low (8)       [0m
echo [96m     Charging and Critical (9)  [0m
echo [96m     Partially Charged (11)     [0m
echo [96m FULL                           [0m
echo [96m     Fully Charged (3)          [0m
echo [96m NOT PRESENT                    [0m
echo [96m     Undefined (10)             [0m
ECHO [96m********************************************[0m

PAUSE