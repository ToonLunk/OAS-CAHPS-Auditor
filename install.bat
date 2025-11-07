@echo off
setlocal EnableDelayedExpansion

echo ====================================
echo OAS-CAHPS Auditor Installer
echo ====================================
echo.

REM Get the directory where this script is located
set "INSTALL_DIR=%~dp0"
set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"

echo This will install the OAS-CAHPS Auditor to:
echo   %INSTALL_DIR%
echo.
echo And add it to your PATH so you can run "audit" from anywhere.
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges...
    set "ADMIN=1"
) else (
    echo NOTE: Not running as administrator.
    echo The installer will add to your USER PATH only.
    echo For system-wide installation, right-click and "Run as administrator"
    set "ADMIN=0"
)
echo.

pause
echo.

REM Check if audit.exe exists
if not exist "%INSTALL_DIR%\audit.exe" (
    echo ERROR: audit.exe not found in this directory!
    echo Make sure you've built the executable first using build_exe.bat
    echo.
    pause
    exit /b 1
)

echo Adding %INSTALL_DIR% to PATH...
echo.

if %ADMIN%==1 (
    REM Add to system PATH (requires admin)
    for /f "skip=2 tokens=3*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYSTEM_PATH=%%a %%b"
    echo !SYSTEM_PATH! | find /i "%INSTALL_DIR%" >nul
    if errorlevel 1 (
        setx /M PATH "%SYSTEM_PATH%;%INSTALL_DIR%" >nul
        echo Added to SYSTEM PATH successfully!
    ) else (
        echo Already in SYSTEM PATH!
    )
) else (
    REM Add to user PATH (no admin needed)
    for /f "skip=2 tokens=3*" %%a in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USER_PATH=%%a %%b"
    if "!USER_PATH!"=="" set "USER_PATH=%PATH%"
    echo !USER_PATH! | find /i "%INSTALL_DIR%" >nul
    if errorlevel 1 (
        setx PATH "!USER_PATH!;%INSTALL_DIR%" >nul
        echo Added to USER PATH successfully!
    ) else (
        echo Already in USER PATH!
    )
)

echo.
echo ====================================
echo INSTALLATION COMPLETE!
echo ====================================
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
