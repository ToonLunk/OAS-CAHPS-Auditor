# Building and Packaging

Instructions for building the executable and creating distribution packages.

## Prerequisites

- Python 3.8 or higher
- All dependencies from `requirements.txt`

## Build Executable

```cmd
build_exe.bat
```

**Output:**

- `dist/audit.exe` - Standalone executable (~20-30 MB)
- `build/` - Temporary build files (safe to delete)
- `audit.spec` - PyInstaller configuration (keep this)

**Build time:** 2-5 minutes (first build), ~30 seconds (subsequent builds)

## Create Distribution Package

```cmd
package.bat
```

**Output:** `OAS-CAHPS-Auditor.zip`

**Contains:**

- `audit.exe` - The executable
- `deploy.bat` - System-wide installation script
- `INSTALL.txt` - Installation instructions

This ZIP file is ready to share. Recipients run `deploy.bat` as administrator to install.

## Installation Methods

**Automated (Recommended):**

- Run `deploy.bat` as administrator
- Installs to `C:\OAS-CAHPS-Auditor`
- Adds installation directory to system PATH
- Requires terminal restart to take effect

**Manual:**

1. Place `audit.exe` in any directory
2. Add that directory to Windows PATH environment variable
3. Restart terminal

**Alternative (User PATH only):**

- Run `install.bat` in same folder as `audit.exe`
- No administrator rights required
- Adds to user PATH instead of system PATH

## Development Workflow

1. Make changes to Python source files
2. Test: `python audit.py <file>`
3. Build: `build_exe.bat`
4. Package: `package.bat` (when ready to distribute)

## Technical Notes

- PyInstaller bundles Python interpreter and all dependencies
- `audit_report.css` is automatically included in the executable
- Antivirus software may flag PyInstaller executables (false positive)
- Users must restart terminal after installation for PATH changes
- The executable is Windows-only (built on Windows, for Windows)
