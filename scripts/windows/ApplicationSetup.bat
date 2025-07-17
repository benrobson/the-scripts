@echo off
title Application Installation Script
goto checkPermissions

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
    timeout 7
    exit
)

:AdminAccess
echo.
echo Success: Administrative permissions confirmed.

set today=%date:~10,4%-%date:~7,2%-%date:~4,2%

echo ============================================
echo ====   Application Installation Script  ====
echo ============================================
echo =======     Created By: Ben Robson   =======
echo ============================================
echo.
echo Installing standard applications...

:: Install Applications using Winget
echo Installing TeamViewer...
winget install --id=TeamViewer.TeamViewer --accept-source-agreements --accept-package-agreements

echo Installing Adobe Acrobat Reader...
winget install --id=Adobe.Acrobat.Reader.64-bit --accept-source-agreements --accept-package-agreements

echo Installing Google Chrome...
winget install --id=Google.Chrome --accept-source-agreements --accept-package-agreements

echo Installing Microsoft 365 Apps for Business...
winget install --id=Microsoft.Office --accept-source-agreements --accept-package-agreements

:: Prepare icon directory
echo Downloading support icon...
mkdir "C:\ProgramData\ReliableIT" >nul 2>&1
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Reliable-IT/the-scripts/refs/heads/master/scripts/assets/favicon.ico' -OutFile 'C:\ProgramData\ReliableIT\support.ico' -UseBasicParsing"

:: Create Shortcut to https://reliableit.au/support with custom icon
echo Creating desktop shortcut to support page...
powershell -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut('C:\Users\Public\Desktop\ReliableIT Support.lnk');$s.TargetPath='https://reliableit.au/support';$s.IconLocation='C:\ProgramData\ReliableIT\support.ico';$s.Save()"

goto exit

:exit
echo.
color a
echo ============================================
echo =======      Setup is complete.      =======
echo ======= Please press any key to exit =======
echo ============================================
pause >nul
