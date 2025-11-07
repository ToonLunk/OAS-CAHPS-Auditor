# Packaging OAS-CAHPS Auditor as an Executable

## Quick Start (Command-Line Tool Style)

To create a command-line tool that works like `audit --all` or `audit <filename.xlsx>`:

1. **Build the executable:**

   - Double-click `build_exe.bat`
   - Wait for the program to build

2. **Install for yourself:**

   - Copy `audit.exe` from the `dist` folder to where you want it
   - Copy `deploy.bat` to be next to `audit.exe`
   - Run `deploy.bat` as administrator (right click it, then click "run as admin")
   - Done! You can now use `audit --all` from anywhere

## What Gets Created

After building, you'll have:

- `dist/audit.exe` - The standalone executable
- `build/` folder - Temporary build files (can be deleted)
- `audit.spec` - Build configuration (keep this)

## Installation Options

### Option 1: Simple Install (Recommended)

1. Copy `audit.exe` and `deploy.bat` to your computer
2. Right-click `deploy.bat` → "Run as administrator"
3. Installs to `C:\OAS-CAHPS-Auditor` and adds to system PATH

### Option 2: Custom Location Install

1. Place `audit.exe` wherever you want (e.g., `C:\Tools`)
2. Copy `install.bat` to that same folder
3. Run `install.bat` (adds that folder to your PATH)

### Option 3: Manual Install

1. Place `audit.exe` in a folder (e.g., `C:\Tools\audit.exe`)
2. Add that folder to PATH manually:
   - Search Windows for "Environment Variables"
   - Edit the "Path" variable
   - Add your folder path
   - Click OK
3. Restart your terminal

## Important Notes

### Building the .exe

- You only need to rebuild when you make changes to the Python code
- The first build takes longer; subsequent builds are faster
- Make sure `audit_report.css` is in the same folder as your Python files

### Using the Command

After installation, you can use it like this:

```
audit --all
audit filename.xlsx
audit C:\path\to\file.xlsx
```

- Works from any directory
- No Python installation required
- No need to navigate to the tool's location

### File Size

- The .exe will be around 20-30 MB (it includes Python and all dependencies)
- This is normal for PyInstaller executables

## Troubleshooting

**"PyInstaller not found" error:**

- The script will automatically install it
- Or manually run: `pip install -r requirements.txt`

**Missing CSS file error:**

- Make sure `audit_report.css` is in the same folder as `audit.py`

**Antivirus warnings:**

- Some antivirus software flags PyInstaller executables
- This is a false positive - you can add an exception

## Updating the Executable

When you make changes to this Python code:

1. Run `build_exe.bat` again
2. Share the new .exe from the `dist` folder
