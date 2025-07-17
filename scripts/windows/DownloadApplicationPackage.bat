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
ECHO ===============  Packages  =================
ECHO ============================================
ECHO 1. Application Setup
ECHO 2. Personalise Computer
ECHO.
ECHO.

ECHO ============================================
ECHO ================ Utilites ==================
ECHO ============================================
ECHO 3. Restart Print Spooler Service
ECHO.
ECHO.

CHOICE /C 12345 /M "Enter your choice:"

IF ERRORLEVEL 3 GOTO RestartPrintSpoolerService
IF ERRORLEVEL 2 GOTO PersonaliseComputer
IF ERRORLEVEL 1 GOTO ApplicationSetup

@REM Application Setup
:ApplicationSetup
CLS
ECHO Downloading ApplicationSetup Package
curl -LJO https://raw.githubusercontent.com/benrobson/the-scripts/main/ApplicationSetup.bat
GOTO END

@REM Personalise Computer
:PersonaliseComputer
CLS
ECHO Downloading PersonaliseComputer Package
curl -LJO https://raw.githubusercontent.com/benrobson/the-scripts/main/PersonaliseComputer.bat
GOTO END

@REM Restart Print Spooler Service
:RestartPrintSpoolerService
CLS
ECHO Downloading RestartPrintSpoolerService Package
curl -LJO https://raw.githubusercontent.com/benrobson/the-scripts/main/utilities/RestartPrintSpoolerService.bat
GOTO END

:END
COLOR A
ECHO.
ECHO.
ECHO ============================================
ECHO ================ Complete ==================
ECHO ============================================
PAUSE
GOTO START