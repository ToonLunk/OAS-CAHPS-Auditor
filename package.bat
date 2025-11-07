@echo off
REM ============================================
REM Build and Package OAS CAHPS Auditor
REM ============================================

echo.
echo ========================================
echo Building OAS CAHPS Auditor Package
echo ========================================
echo.

REM Step 1: Build the executable
echo [1/3] Building executable...
call build_exe.bat
if errorlevel 1 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

REM Step 2: Create distribution folder
echo.
echo [2/3] Preparing distribution package...
if exist "distribution" rmdir /s /q distribution
mkdir distribution

REM Copy files to distribution folder
copy "dist\audit.exe" "distribution\" >nul
copy "deploy.bat" "distribution\" >nul
copy "INSTALL.txt" "distribution\" >nul

REM Step 3: Create ZIP file
echo.
echo [3/3] Creating ZIP archive...

REM Use PowerShell to create the ZIP
powershell -Command "Compress-Archive -Path 'distribution\*' -DestinationPath 'OAS-CAHPS-Auditor.zip' -Force"

if exist "OAS-CAHPS-Auditor.zip" (
    echo.
    echo ========================================
    echo SUCCESS!
    echo ========================================
    echo.
    echo Package created: OAS-CAHPS-Auditor.zip
    echo.
    echo This ZIP contains:
    echo   - audit.exe
    echo   - deploy.bat
    echo   - INSTALL.txt
    echo.
    echo Ready to share!
    echo ========================================
) else (
    echo ERROR: Failed to create ZIP file
)

echo.
pause
