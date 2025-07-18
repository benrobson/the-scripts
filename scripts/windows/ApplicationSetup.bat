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
    goto main
) else (
    color CE
    echo Failure: Current permissions are inadequate.
    echo Please run as administrator and try again.
    timeout 7
    exit
)

:main
echo.
echo ============================================
echo ====   Application Installation Script  ====
echo ============================================
echo =======     Created By: Ben Robson   =======
echo ============================================
echo.
echo Checking for updates...
winget upgrade
powershell -Command "$updateSession = New-Object -ComObject Microsoft.Update.Session; $updateSearcher = $updateSession.CreateUpdateSearcher(); $searchResult = $updateSearcher.Search('IsInstalled=0'); if ($searchResult.Updates.Count -gt 0) { $searchResult.Updates } else { echo 'No Windows updates available.' }"
echo.
set /p installUpdates="Install all updates? (y/n): "
if /i "%installUpdates%"=="y" (
    winget upgrade --all --accept-package-agreements --accept-source-agreements
    powershell -Command "$updateSession = New-Object -ComObject Microsoft.Update.Session; $updateSearcher = $updateSession.CreateUpdateSearcher(); $searchResult = $updateSearcher.Search('IsInstalled=0'); if ($searchResult.Updates.Count -gt 0) { $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl; foreach ($update in $searchResult.Updates) { $updatesToInstall.Add($update) }; $installer = $updateSession.CreateUpdateInstaller(); $installer.Updates = $updatesToInstall; $installer.Install() }"
)
goto postInstall

:postInstall
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
endlocal
