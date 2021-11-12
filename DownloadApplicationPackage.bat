@ECHO OFF
@REM Inital Setup
@REM Set Dimensions
MODE 60,30
TITLE Script Repo Download Library

:START
COLOR 9
CLS
ECHO ========================================================
ECHO =======        Script Repo Download Library      =======
ECHO ======= https://github.com/benrobson/the-scripts =======
ECHO ========================================================
ECHO.
ECHO.

ECHO ============================================
ECHO =============== Categories =================
ECHO ============================================
ECHO 1. Application Setup
ECHO 2. Personalise Computer
ECHO.
ECHO.

ECHO ============================================
ECHO ================ Bundles ===================
ECHO ============================================
ECHO 3. Computer Run Up Bundle
ECHO.
ECHO.

ECHO ============================================
ECHO ================ Utilites ==================
ECHO ============================================
ECHO 4. Restart Print Spooler Service
ECHO.
ECHO.

CHOICE /C 12345 /M "Enter your choice:"

IF ERRORLEVEL 1 GOTO ApplicationSetup
IF ERRORLEVEL 2 GOTO PersonaliseComputer
IF ERRORLEVEL 3 GOTO ComputerRunUpBundle
IF ERRORLEVEL 4 GOTO RestartPrintSpoolerService

@REM Application Setup
:ApplicationSetup
CLS
ECHO Downloading ApplicationSetup Package
curl -LJO https://github.com/benrobson/the-scripts/raw/main/ApplicationSetup.bat
GOTO End

@REM Personalise Computer
:Shutdown
CLS
ECHO Downloading PersonaliseComputer Package
curl -LJO https://github.com/benrobson/the-scripts/raw/main/PersonaliseComputer.bat
GOTO End

@REM Computer Run Up Bundle
:ComputerRunUpBundle
CLS
ECHO Downloading ApplicationSetup and PersonaliseComputer Package
curl -LJO https://github.com/benrobson/the-scripts/raw/main/ApplicationSetup.bat
curl -LJO https://github.com/benrobson/the-scripts/raw/main/PersonaliseComputer.bat
GOTO End

@REM Restart Print Spooler Service
:RestartPrintSpoolerService
CLS
ECHO Downloading RestartPrintSpoolerService Package
curl -LJO https://github.com/benrobson/the-scripts/raw/main/RestartPrintSpoolerService.bat
GOTO End

:END
COLOR A
ECHO.
ECHO.
ECHO ============================================
ECHO ================ Complete ==================
ECHO ============================================
PAUSE
GOTO START