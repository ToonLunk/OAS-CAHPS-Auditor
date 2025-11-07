@echo off
echo ====================================
echo Building OAS-CAHPS Auditor Executable
echo ====================================
echo.

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>NUL
if errorlevel 1 (
    echo PyInstaller not found. Installing dependencies...
    pip install -r requirements.txt
    echo.
)

echo Building executable...
echo This may take a few minutes...
echo.

python -m PyInstaller --onefile --name "audit" --add-data "audit_report.css;." --console audit.py

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
echo   1. Copy audit.exe to a permanent location (or keep in dist folder)
echo   2. Run install.bat to add it to your PATH
echo   3. Share audit.exe + install.bat with coworkers
echo.
echo After installation, you and your coworkers can run:
echo   audit --all
echo   audit filename.xlsx
echo from anywhere!
echo.
pause
