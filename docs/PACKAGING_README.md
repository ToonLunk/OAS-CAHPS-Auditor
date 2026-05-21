# Building and Packaging

Instructions for building the executable and creating distribution packages.

## Prerequisites

- Python 3.8 or higher
- All dependencies from `requirements.txt`
- [NSIS 3.x](https://nsis.sourceforge.io/) with the [EnVar plugin](https://nsis.sourceforge.io/EnVar_plug-in) installed
- `makensis.exe` must be on your system PATH

## Version Management

Version is stored in `.env` file at project root:

```
VERSION=0.50-rc1
```

As well as in audit.py as a constant, and version_info.txt.

<!-- note: should probably figure out a way to onlu have this in one place -->

This version is used by:

- The executable (shown in `audit --version`)
- The package script (creates `OAS-CAHPS-Auditor-v{VERSION}-Setup.exe`)

## Create Distribution Package (for users; reccommended)

```cmd
scripts\package.bat
```

This automatically runs `build_exe.bat` first, then compiles the NSIS installer.

**Output:** `dist/OAS-CAHPS-Auditor-v{VERSION}-Setup.exe`

The installer bundles:
- `audit.exe` - The executable
- `cpt_codes.json` - CPT code configuration
- `Installation Instructions.txt`
- `About SIDs.csv.txt` - Instructions for downloading SIDs.csv
- `LICENSE`

> **Note:** `SIDs.csv` is NOT included in the installer. Users download it
> separately from the [shared OneDrive folder](https://jlm353-my.sharepoint.com/:f:/g/personal/dcdata_jlm-solutions_com/IgBhYR7tt6YTRbgNTDEh9M7xAc5HSCC3KSaJt6ImfJV65kg?e=hKp0ZU).
> The app shows the download link in the terminal and in the HTML report's info tooltip.

Recipients run the Setup.exe wizard to install.

## Installation Methods

**Setup Wizard:**

- Run `OAS-CAHPS-Auditor-v{VERSION}-Setup.exe`
- Installs to `C:\OAS-CAHPS-Auditor`
- Adds installation directory to system PATH
- Optionally registers right-click context menu entries
- Creates an uninstaller accessible via Add/Remove Programs

## Development Workflow

1. Make changes to Python source files
2. Update version if needed (reserve for somewhat substantial changes/bug fixes since this will popup on everyone's audits)
3. Test: `python audit.py <file>`
4. Package: `scripts\package.bat` (when ready to distribute)
5. Upload the installer to GitHub Releases
6. Update `SIDs.csv` on the shared OneDrive when clients change (monthly)

## Technical Notes

- PyInstaller bundles Python interpreter and all dependencies - users don't need Python or libraries
- Antivirus software may flag PyInstaller executables (false positive)
- The executable is Windows-only. This could be built for other platforms with adjustments to the build process, but currently only Windows is supported
