# Building and Packaging

Instructions for building the executable and creating distribution packages.

## Prerequisites

- Python 3.8 or higher
- All dependencies from `requirements.txt`

## Version Management

Version is stored in `.env` file at project root:

```
VERSION=0.50-rc1
```

This version is used by:

- The executable (shown in `audit --version`)
- The package script (creates `OAS-CAHPS-Auditor-v{VERSION}.zip`)

To update: Edit `.env` and change the `VERSION=` line.

## Build Executable

```cmd
scripts\build_exe.bat
```

**Output:**

- `dist/audit.exe` - Standalone executable (~20-30 MB)
- `build/` - Temporary build files (safe to delete)
- `audit.spec` - PyInstaller configuration (keep this)

**Build time:** 2-5 minutes (first build), ~30 seconds (subsequent builds)

## Create Distribution Package

```cmd
scripts\package.bat
```

This automatically runs `build_exe.bat` first, then packages everything.

**Output:** `OAS-CAHPS-Auditor-v{VERSION}.zip`

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
2. Update version in `.env` if needed
3. Test: `python audit.py <file>`
4. Build: `scripts\build_exe.bat`
5. Package: `scripts\package.bat` (when ready to distribute)

## Project Structure

```
project/
├── audit.py                    # Main source
├── audit_lib_funcs.py          # Validation functions
├── audit_printer.py            # Report generation
├── audit_report.css            # Report styling
├── .env                        # Version configuration
├── requirements.txt            # Python dependencies
├── readme.md                   # Main documentation
├── scripts/                    # Build & deployment scripts
│   ├── package.bat
│   ├── build_exe.bat
│   ├── deploy.bat
│   └── install.bat
└── docs/                       # Documentation
    ├── PACKAGING_README.md
    └── INSTALL.txt
```

## Technical Notes

- PyInstaller bundles Python interpreter and all dependencies
- `audit_report.css` is automatically included in the executable
- Antivirus software may flag PyInstaller executables (false positive)
- Users must restart terminal after installation for PATH changes
- The executable is Windows-only (built on Windows, for Windows)
