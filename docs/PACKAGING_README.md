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

As well as in audit.py as a constant.

This version is used by:

- The executable (shown in `audit --version`)
- The package script (creates `OAS-CAHPS-Auditor-v{VERSION}.zip`)

To update: Edit `.env` and change the `VERSION=` line, then update the `VERSION` constant in `audit.py` to match.

## Build Executable (for development)

```cmd
scripts\build_exe.bat
```

**Output:**

- `dist/audit.exe` - Standalone executable (~20-30 MB)
- `build/` - Temporary build files (safe to delete)
- `audit.spec` - PyInstaller configuration (keep this)

**Build time:** 2-5 minutes (first build), ~30 seconds (subsequent builds)

## Create Distribution Package (for users; reccommended)

```cmd
scripts\package.bat
```

This automatically runs `build_exe.bat` first, then packages everything. It's recommended to use `package.bat` instead of just `build_exe.bat` when you're ready to share the tool, as it ensures the executable is built and packaged correctly, as well as including installation instructions and other necessary files.

**Output:** `OAS-CAHPS-Auditor-v{VERSION}.zip`

**Contains:**

- `audit.exe` - The executable
- `deploy.bat` - System-wide installation script
- `Installation Instructions.txt` - Installation instructions

This ZIP file is ready to share. Recipients run `deploy.bat` as administrator to install.

## Installation Methods

**Automated (Recommended):**

- Run `deploy.bat` as administrator
- Installs to `C:\OAS-CAHPS-Auditor`
- Adds installation directory to system PATH
- Requires terminal restart to take effect

**Manual:**

1. Place `audit.exe` and all files in `C:\OAS-CAHPS-Auditor` (or any directory you choose)
2. Add that directory to Windows PATH environment variable
3. Restart terminal

## Development Workflow

1. Make changes to Python source files
2. Update version in `.env` if needed
3. Test: `python audit.py <file>`
4. Build: `scripts\build_exe.bat`
5. Package: `scripts\package.bat` (when ready to distribute)

## Technical Notes

- PyInstaller bundles Python interpreter and all dependencies - no need for users to install Python or libraries
- `audit_report.css` is automatically included in the executable
- Antivirus software may flag PyInstaller executables (false positive)
- The executable is Windows-only. This could be built for other platforms with adjustments to the build process, but currently only Windows is supported
