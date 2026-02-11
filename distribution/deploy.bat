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

echo Copying LICENSE to %INSTALL_DIR%...
copy /Y "%SCRIPT_DIR%LICENSE" "%INSTALL_DIR%\LICENSE" >nul

REM Handle cpt_codes.json installation
if not exist "%INSTALL_DIR%\cpt_codes.json" (
    echo Installing default cpt_codes.json...
    copy /Y "%SCRIPT_DIR%cpt_codes.json" "%INSTALL_DIR%\cpt_codes.json" >nul
    goto :cpt_done
)

echo.
echo Existing cpt_codes.json configuration found.
:ask_overwrite
set /p "OVERWRITE=Do you want to install the new version? (Y/N): "
if /i "%OVERWRITE%"=="Y" (
    echo Installing new cpt_codes.json...
    copy /Y "%SCRIPT_DIR%cpt_codes.json" "%INSTALL_DIR%\cpt_codes.json" >nul
    goto :cpt_done
)
if /i "%OVERWRITE%"=="N" (
    echo Keeping existing cpt_codes.json configuration...
    goto :cpt_done
)
echo Invalid input. Please enter Y or N.
goto :ask_overwrite
:cpt_done

REM Handle SIDs.csv installation (optional file)
if exist "%SCRIPT_DIR%SIDs.csv" (
    if not exist "%INSTALL_DIR%\SIDs.csv" (
        echo Installing default SIDs.csv...
        copy /Y "%SCRIPT_DIR%SIDs.csv" "%INSTALL_DIR%\SIDs.csv" >nul
        goto :sids_done
    )

    echo.
    echo Existing SIDs.csv found.
    :ask_overwrite_sids
    set /p "OVERWRITE_SIDS=Do you want to install the new version? (Y/N): "
    if /i "%OVERWRITE_SIDS%"=="Y" (
        echo Installing new SIDs.csv...
        copy /Y "%SCRIPT_DIR%SIDs.csv" "%INSTALL_DIR%\SIDs.csv" >nul
        goto :sids_done
    )
    if /i "%OVERWRITE_SIDS%"=="N" (
        echo Keeping existing SIDs.csv...
        goto :sids_done
    )
    echo Invalid input. Please enter Y or N.
    goto :ask_overwrite_sids
    :sids_done
)

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
echo Installing context menu integration...
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%register_context_menu.ps1"

echo.
echo ====================================
echo INSTALLATION COMPLETE!
echo ====================================
echo.
echo The audit tool is now installed at:
echo   %INSTALL_DIR%\audit.exe
echo.
echo Context menus installed:
echo   - Right-click inside folders: "Audit All OAS Files"
echo   - Right-click Excel files: "Audit This OAS File"
echo.
echo IMPORTANT: Close and reopen any Command Prompt or PowerShell windows
echo for the PATH change to take effect.
echo.
echo You can also run from command line:
echo   audit --all
echo   audit filename.xlsx
echo.
pause
