@echo off

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

NET stop spooler
NET start spooler

goto completed
:completed
powercfg.exe /h off
powercfg.exe -change -standby-timeout-ac 0
powercfg.exe -change -monitor-timeout-ac 60
cls
color a
echo ============================================
echo =======   Operation is complete.     =======
echo ======= Please press any key to exit =======
echo ============================================
pause