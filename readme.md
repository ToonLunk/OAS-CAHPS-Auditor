# OAS CAHPS Auditor

Command-line tool for auditing OAS CAHPS Excel files. Validates headers, sample sizes, addresses, CPT codes, cross-tab consistency, and data quality. Outputs HTML reports.

## Usage

```cmd
audit filename.xlsx    # Audit a specific file
audit --all            # Audit all Excel files in current directory
audit --version        # Show version number
```

## Installation

If you have the distribution package (ZIP file):

1. Extract the ZIP file
2. Right-click `deploy.bat` and select "Run as administrator"
3. Restart your terminal
4. Verify installation by running: `audit --help`

The tool installs to `C:\OAS-CAHPS-Auditor` and is added to your system PATH, so you can run `audit` from anywhere.

See `INSTALL.txt` in the package for detailed instructions and troubleshooting.

## Building from Source

**Requirements:** Python 3.8+

**Build executable:**

```cmd
pip install -r requirements.txt
build_exe.bat
```

Output: `dist/audit.exe`

**Create distribution package:**

```cmd
package.bat
```

Output: `OAS-CAHPS-Auditor.zip` (ready to share)

**Development:**

- Edit source files: `audit.py`, `audit_lib_funcs.py`, `audit_printer.py`, `audit_report.css`
- Test changes: `python audit.py <file>`
- Rebuild: `build_exe.bat`

See `PACKAGING_README.md` for build details.
