@echo off
REM ============================================
REM Build and Package OAS CAHPS Auditor
REM ============================================
REM Builds the executable with PyInstaller, then
REM compiles the NSIS installer (Setup.exe).
REM
REM Prerequisites:
REM   - Python 3.8+ with dependencies from requirements.txt
REM   - NSIS 3.x (makensis.exe must be on PATH)
REM   - NSIS EnVar plugin installed
REM
REM SIDs.csv is NOT included - users must download it from the
REM shared OneDrive folder and place it in the install directory.
REM ============================================

REM Change to project root directory
cd /d "%~dp0.."

REM Read version from .env file
for /f "tokens=2 delims==" %%a in ('findstr "VERSION" .env') do set VERSION=%%a

echo.
echo ========================================
echo Building OAS CAHPS Auditor Package
echo Version: %VERSION%
echo ========================================
echo.

REM Step 1: Build the executable
echo [1/2] Building executable...
call scripts\build_exe.bat
if errorlevel 1 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

REM Step 2: Build the installer
echo.
echo [2/2] Building NSIS installer...

REM Check that makensis is available
set "MAKENSIS="
where makensis >nul 2>&1
if not errorlevel 1 (
    set "MAKENSIS=makensis"
) else if exist "C:\Program Files (x86)\NSIS\makensis.exe" (
    set "MAKENSIS=C:\Program Files (x86)\NSIS\makensis.exe"
) else if exist "C:\Program Files\NSIS\makensis.exe" (
    set "MAKENSIS=C:\Program Files\NSIS\makensis.exe"
) else (
    echo ERROR: makensis.exe not found!
    echo Please install NSIS from https://nsis.sourceforge.io/
    pause
    exit /b 1
)

set SETUP_NAME=OAS-CAHPS-Auditor-v%VERSION%-Setup.exe
"%MAKENSIS%" /DVERSION=%VERSION% installer\audit_installer.nsi

if exist "dist\%SETUP_NAME%" (
    echo.
    echo ========================================
    echo SUCCESS!
    echo ========================================
    echo.
    echo Installer created: dist\%SETUP_NAME%
    echo.
    echo NOTE: SIDs.csv is NOT included in the installer.
    echo Users must download it separately from the shared OneDrive folder.
    echo ========================================
) else (
    echo ERROR: Failed to create installer!
)

echo.
pause
