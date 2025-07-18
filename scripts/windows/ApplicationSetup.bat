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
    goto checkWinget
) else (
    color CE
    echo Failure: Current permissions are inadequate.
    echo Please run as administrator and try again.
    timeout 7
    exit
)

:checkWinget
echo Checking for winget...
winget --version >nul 2>&1
if %errorLevel% == 0 (
    echo Winget is already installed.
    goto installApps
) else (
    echo Winget is not installed. Attempting to install...
    goto installWinget
)

:installWinget
echo "Winget is not installed. Attempting to install..."
echo "Downloading dependencies..."
powershell -Command "Invoke-WebRequest -Uri 'https://github.com/microsoft/vclibs/releases/download/v14.0.30704.0/Microsoft.VCLibs.x64.14.00.Desktop.appx' -OutFile '%TEMP%\vclibs.appx' -UseBasicParsing"
powershell -Command "Invoke-WebRequest -Uri 'https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.8.6' -OutFile '%TEMP%\xaml.zip' -UseBasicParsing"
powershell -Command "Expand-Archive -Path '%TEMP%\xaml.zip' -DestinationPath '%TEMP%\xaml' -Force"
echo "Installing dependencies..."
powershell -Command "Add-AppxPackage -Path '%TEMP%\vclibs.appx'"
powershell -Command "Add-AppxPackage -Path '%TEMP%\xaml\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.8.appx'"
echo "Downloading winget..."
powershell -Command "Invoke-WebRequest -Uri 'https://github.com/microsoft/winget-cli/releases/download/v1.11.400/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle' -OutFile '%TEMP%\winget.msixbundle' -UseBasicParsing"
echo "Installing winget..."
powershell -Command "Add-AppxPackage -Path '%TEMP%\winget.msixbundle'"
winget --version >nul 2>&1
if %errorLevel% == 0 (
    echo Winget installed successfully.
    goto installApps
) else (
    color CE
    echo Failed to install winget.
    timeout 7
    exit
)

:installApps
echo.
echo ============================================
echo ====   Application Installation Script  ====
echo ============================================
echo =======     Created By: Ben Robson   =======
echo ============================================
echo.
echo What would you like to do?
echo.
echo 1. Install standard applications
echo 2. Check for updates
echo.
set /p choice="Enter your choice: "
if /i "%choice%"=="1" goto installStandardApps
if /i "%choice%"=="2" goto checkUpdates
goto exit

:installStandardApps
echo Installing standard applications...
call :installApp "TeamViewer" "TeamViewer.TeamViewer"
call :installApp "Adobe Acrobat Reader" "Adobe.Acrobat.Reader.64-bit"
call :installApp "Google Chrome" "Google.Chrome"
call :installApp "Microsoft 365 Apps for Business" "Microsoft.Office"
goto postInstall

:checkUpdates
echo Checking for updates...
winget upgrade
echo.
set /p installUpdates="Install all updates? (y/n): "
if /i "%installUpdates%"=="y" winget upgrade --all --accept-package-agreements --accept-source-agreements
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
