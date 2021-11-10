@REM 
@REM Personalise Computer Script
@REM ===============================
@REM Features
@REM * Hides Cortana Button.
@REM * Hides Task View Button.
@REM * Hides People Button.
@REM * Removes News and Interests from Taskbar.
@REM * Clear all Taskbar items.
@REM * Adds This PC icon to Desktop.
@REM
@REM To-Do
@REM * Removes all Start Menu Icons/Tiles.
@REM * Add User's Files to Desktop.
@REM ===============================
@REM 

@echo off
@REM Set title of window
title Personalise Computer Script
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
echo ====   Personalise   Computer   Script  ====
echo ============================================
echo =======     Created By: Ben Robson   =======
echo ============================================
echo.
echo.

cls

@REM Hide Cortana Button
REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v AllowCortana /t REG_DWORD /d 0 /f

@REM Hide Task View Button
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v ShowTaskViewButton /t REG_DWORD /d 0 /f

@REM Hide People Button
REG ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People" /v PeopleBand /t REG_DWORD /d 0 /f

@REM Hide News and Interests
REG ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Feeds" /V ShellFeedsTaskbarViewMode /T REG_DWORD /D 0 /F

@REM Clear all Taskbar items
DEL /F /S /Q /A "%AppData%\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*"
REG DELETE HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband /F

@REM Add This PC to Desktop
Set KeyToSet=HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel\

Set ThisPCGuid={20D04FE0-3AEA-1069-A2D8-08002B30309D}

Set ShowIconValue=0
Set HideIconValue=1

REG ADD %KeyToSet% /v %ThisPCGuid% /t REG_DWORD /d %ShowIconValue% /f

@REM Restart Explorer to take effect
taskkill /F /IM explorer.exe & start explorer