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

REM Update version in audit.py from .env file
powershell -Command "(Get-Content audit.py) -replace '__version__ = \".*\"', '__version__ = \"%VERSION%\"' | Set-Content audit.py"

REM Update version in version_info.txt from .env file
REM Splits VERSION (e.g. 1.2.3) into filevers/prodvers tuple and string fields
powershell -Command "$v='%VERSION%'; $parts=($v -split '\.'); $tuple=\"($($parts[0]), $($parts[1]), $($parts[2]), 0)\"; (Get-Content version_info.txt) -replace 'filevers=\(.*?\)', \"filevers=$tuple\" -replace 'prodvers=\(.*?\)', \"prodvers=$tuple\" -replace \"FileVersion', u'.*?'\", \"FileVersion', u'$v'\" -replace \"ProductVersion', u'.*?'\", \"ProductVersion', u'$v'\" | Set-Content version_info.txt"

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
echo   Run scripts\package.bat to build the installer
echo   Or copy dist\audit.exe to C:\OAS-CAHPS-Auditor manually
echo.
echo After installation, run:
echo   audit --all
echo   audit filename.xlsx
echo from anywhere!
echo.
pause
