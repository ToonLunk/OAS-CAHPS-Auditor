@echo off
REM Change to project root directory
cd /d "%~dp0.."

REM Read version from .env file
for /f "tokens=2 delims==" %%a in ('findstr "VERSION" .env') do set VERSION=%%a

echo ====================================
echo Building OAS-CAHPS Auditor Executable
echo Version: %VERSION%
echo ====================================
echo.

REM Update version in audit.py
powershell -Command "(Get-Content audit.py) -replace '__version__ = \".*\"', '__version__ = \"%VERSION%\"' | Set-Content audit.py"

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>NUL
if errorlevel 1 (
    echo PyInstaller not found. Installing dependencies...
    pip install -r requirements.txt
    echo.
)

echo Cleaning old builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo.

echo Building executable using audit.spec...
echo This may take a few minutes...
echo.

python -m PyInstaller audit.spec

if errorlevel 1 (
    echo.
    echo BUILD FAILED!
    echo Check the error messages above.
    pause
    exit /b 1
)

echo.
echo ====================================
echo BUILD SUCCESSFUL!
echo ====================================
echo.
echo Your executable is located at:
echo   dist\audit.exe
echo.
echo Next steps:
echo   1. Copy `audit.exe` and `deploy.bat` to the same folder
echo   2. Run deploy.bat to add it to your PATH
echo   3. Run the auditor from anywhere!
echo.
echo After installation, run:
echo   audit --all
echo   audit filename.xlsx
echo from anywhere!
echo.
pause
