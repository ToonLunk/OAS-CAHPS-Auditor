@echo off
REM ============================================
REM Build and Package OAS CAHPS Auditor
REM ============================================
REM SIDs.csv is NOT included - it lives on a shared OneDrive folder
REM and the app reads it from there automatically at runtime.
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
echo [1/3] Building executable...
call scripts\build_exe.bat
if errorlevel 1 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

REM Step 2: Prepare distribution folder
echo.
echo [2/3] Preparing distribution package...

REM Update all files in distribution folder from their source of truth
copy /Y "dist\audit.exe" "distribution\" >nul
copy /Y "docs\Installation Instructions.txt" "distribution\" >nul
copy /Y "LICENSE" "distribution\" >nul
copy /Y "cpt_codes.json" "distribution\" >nul
copy /Y "scripts\register_context_menu.ps1" "distribution\" >nul
copy /Y "scripts\unregister_context_menu.ps1" "distribution\" >nul
copy /Y "scripts\deploy.bat" "distribution\" >nul
copy /Y "docs\About SIDs.csv.txt" "distribution\" >nul

REM Remove any leftover sensitive files from distribution folder
if exist "distribution\SIDs.csv" del "distribution\SIDs.csv"
if exist "distribution\hospital_names.csv" del "distribution\hospital_names.csv"

echo Distribution folder updated.

REM Step 3: Create ZIP file
echo.
echo [3/3] Creating ZIP archive...

set ZIPNAME=OAS-CAHPS-Auditor-v%VERSION%.zip
if exist "%ZIPNAME%" del "%ZIPNAME%"
powershell -Command "Compress-Archive -Path 'distribution\*' -DestinationPath '%ZIPNAME%' -Force"

if exist "%ZIPNAME%" (
    echo.
    echo ========================================
    echo SUCCESS!
    echo ========================================
    echo.
    echo Package created: %ZIPNAME%
    echo.
    echo This ZIP contains:
    echo   - audit.exe
    echo   - deploy.bat
    echo   - register_context_menu.ps1
    echo   - unregister_context_menu.ps1
    echo   - Installation Instructions.txt
    echo   - LICENSE
    echo   - cpt_codes.json
    echo   - About SIDs.csv.txt
    echo.
    echo NOTE: SIDs.csv is NOT included.
    echo The app reads it from the shared OneDrive folder automatically.
    echo ========================================
) else (
    echo ERROR: Failed to create ZIP file
)

echo.
pause
