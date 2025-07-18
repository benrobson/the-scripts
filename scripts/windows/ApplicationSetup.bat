@echo off
setlocal
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
echo What would you like to do?
echo.
echo 1. Install standard applications
echo 2. Check for Windows updates
echo.
set /p choice="Enter your choice: "
if /i "%choice%"=="1" goto installStandardApps
if /i "%choice%"=="2" goto checkWindowsUpdates
goto exit

:installStandardApps
echo Installing standard applications...
call :installApp "TeamViewer" "TeamViewer.TeamViewer"
call :installApp "Adobe Acrobat Reader" "Adobe.Acrobat.Reader.64-bit"
call :installApp "Google Chrome" "Google.Chrome"
call :installApp "Microsoft 365 Apps for Business" "Microsoft.Office"
goto postInstall

:checkWindowsUpdates
echo Checking for Windows updates...
powershell -Command "$updateSession = New-Object -ComObject Microsoft.Update.Session; $updateSearcher = $updateSession.CreateUpdateSearcher(); $searchResult = $updateSearcher.Search('IsInstalled=0'); $searchResult.Updates"
echo.
set /p installUpdates="Install all updates? (y/n): "
if /i "%installUpdates%"=="y" powershell -Command "$updateSession = New-Object -ComObject Microsoft.Update.Session; $updateSearcher = $updateSession.CreateUpdateSearcher(); $searchResult = $updateSearcher.Search('IsInstalled=0'); $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl; foreach ($update in $searchResult.Updates) { $updatesToInstall.Add($update) }; $installer = $updateSession.CreateUpdateInstaller(); $installer.Updates = $updatesToInstall; $installer.Install()"
goto exit

:postInstall
:: Prepare icon directory
echo Downloading support icon...
mkdir "C:\ProgramData\ReliableIT" >nul 2>&1
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Reliable-IT/the-scripts/refs/heads/master/scripts/assets/favicon.ico' -OutFile 'C:\ProgramData\ReliableIT\support.ico' -UseBasicParsing"

:: Create Shortcut to https://reliableit.au/support with custom icon
echo Creating desktop shortcut to support page...
powershell -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut('C:\Users\Public\Desktop\ReliableIT Support.lnk');$s.TargetPath='https://reliableit.au/support';$s.IconLocation='C:\ProgramData\ReliableIT\support.ico';$s.Save()"

goto exit

:installApp
echo Installing %~1...
winget install --id=%~2 --accept-source-agreements --accept-package-agreements
goto :eof

:exit
echo.
color a
echo ============================================
echo =======      Setup is complete.      =======
echo ======= Please press any key to exit =======
echo ============================================
pause >nul
endlocal
