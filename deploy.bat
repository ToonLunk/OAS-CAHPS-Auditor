@echo off
echo ====================================
echo Quick Deploy - OAS-CAHPS Auditor
echo ====================================
echo.
echo This script will:
echo   1. Copy audit.exe to C:\OAS-CAHPS-Auditor
echo   2. Add it to your PATH
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script requires administrator privileges.
    echo Right-click this file and select "Run as administrator"
    echo.
    pause
    exit /b 1
)

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Check if audit.exe exists
if not exist "%SCRIPT_DIR%audit.exe" (
    echo ERROR: audit.exe not found in this directory!
    echo Looking in: %SCRIPT_DIR%
    echo Make sure audit.exe is in the same folder as this script.
    echo.
    pause
    exit /b 1
)

REM Create installation directory
set "INSTALL_DIR=C:\OAS-CAHPS-Auditor"
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo Copying audit.exe to %INSTALL_DIR%...
copy /Y "%SCRIPT_DIR%audit.exe" "%INSTALL_DIR%\audit.exe" >nul

echo Adding %INSTALL_DIR% to system PATH...

REM Get current system PATH
for /f "skip=2 tokens=3*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYSTEM_PATH=%%a %%b"

REM Check if already in PATH
echo %SYSTEM_PATH% | find /i "%INSTALL_DIR%" >nul
if errorlevel 1 (
    setx /M PATH "%SYSTEM_PATH%;%INSTALL_DIR%" >nul
    echo Added to PATH successfully!
) else (
    echo Already in PATH!
)

echo.
echo ====================================
echo INSTALLATION COMPLETE!
echo ====================================
echo.
echo The audit tool is now installed at:
echo   %INSTALL_DIR%\audit.exe
echo.
echo IMPORTANT: Close and reopen any Command Prompt or PowerShell windows
echo for the PATH change to take effect.
echo.
echo After reopening your terminal, you can run:
echo   audit --all
echo   audit filename.xlsx
echo.
echo from any directory!
echo.
pause
