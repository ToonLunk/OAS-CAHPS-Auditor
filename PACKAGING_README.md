# Packaging OAS-CAHPS Auditor as an Executable

## Quick Start (Command-Line Tool Style)

To create a command-line tool that works like `audit --all` or `audit filename.xlsx`:

1. **Build the executable:**

   - Double-click `build_exe.bat`
   - Wait for the build to complete (may take a few minutes the first time)

2. **Install for yourself:**

   - Copy `audit.exe` from the `dist` folder to where you want it
   - Double-click `install.bat` (in the folder with `audit.exe`)
   - OR use `deploy.bat` for system-wide installation (requires admin)

3. **Share with coworkers:**
   - Give them `audit.exe` and `deploy.bat`
   - They run `deploy.bat` as administrator
   - Done! They can now use `audit --all` from anywhere

## What Gets Created

After building, you'll have:

- `dist/audit.exe` - The standalone executable
- `build/` folder - Temporary build files (can be deleted)
- `audit.spec` - Build configuration (keep this)

## Installation Options

### Option 1: Simple Install (Recommended for Coworkers)

1. Copy `audit.exe` and `deploy.bat` to their computer
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

### For You (Building the .exe)

- You only need to rebuild when you make changes to the Python code
- The first build takes longer; subsequent builds are faster
- Make sure `audit_report.css` is in the same folder as your Python files

### For Your Coworkers (Using the Command)

After installation, they can use it just like you do:

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

## What to Share with Coworkers

**Easy way (recommended):**

- `audit.exe` (from the `dist` folder)
- `deploy.bat`
- Tell them: "Run deploy.bat as administrator"

**They'll then be able to run:**

```
audit --all
audit myfile.xlsx
```

## How It Works

1. The executable is named `audit.exe` (not a long name)
2. The installer adds its location to the Windows PATH
3. Windows can now find `audit` from any directory
4. Just like how you can run `python` or `git` from anywhere

## Troubleshooting

**"PyInstaller not found" error:**

- The script will automatically install it
- Or manually run: `pip install -r requirements.txt`

**Missing CSS file error:**

- Make sure `audit_report.css` is in the same folder as `audit.py`

**Antivirus warnings:**

- Some antivirus software flags PyInstaller executables
- This is a false positive - you can add an exception

**The .exe doesn't work on coworker's computer:**

- Make sure they're using Windows
- The .exe is only compatible with Windows OS
- They may need to allow it through Windows Defender

## Updating the Executable

When you make changes to your Python code:

1. Run `build_exe.bat` again
2. Share the new .exe from the `dist` folder
