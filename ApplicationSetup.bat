@REM 
@REM Application Setup.bat
@REM 
@REM =========================

@echo off
@REM Set title of window
title Application Installation Script
goto checkPermissions

@REM Check permissions that the user is currently running
:checkPermissions
echo Administrative permissions are required to run this script. Detecting permissions...

net session >nul 2>&1
if %errorLevel% == 0 (
	color 9
    echo Success: Administrative permissions confirmed.
    goto AdminAccess
) else (
   color CE
   echo Failure: Current permissions are inadequate.
   echo Please run as administrator and try again.
   TIMEOUT 7
   Exit
)

pause >nul

:AdminAccess
echo Success: Administrative permissions confirmed.
echo.
echo.

ewfmgr.exe C: -disable

set today=%date:~10,4%-%date:~7,2%-%date:~4,2%
cls
echo ============================================
echo ====   Application Installation Script  ====
echo ============================================
echo =======     Created By: Ben Robson   =======
echo ============================================
echo.
echo.

cls
@REM Installing Standard Applications
echo ============================================
echo =======   Installing Standard Apps   =======
echo =======        Please wait...        =======
echo ============================================

@REM Download Chocolatey to download standard applications
@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"

choco install teamviewer --file64 -y
choco install adobereader --file64 -y
choco install googlechrome --file64 -y

goto :exit
:exit
powercfg.exe /h off
powercfg.exe -change -standby-timeout-ac 0
powercfg.exe -change -monitor-timeout-ac 60
cls
color a
echo ============================================
echo =======      Setup is complete.      =======
echo ======= Please press any key to exit =======
echo ============================================
pause
