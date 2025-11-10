# OAS CAHPS Auditor

Command-line tool for auditing OAS CAHPS Excel files. Validates headers, sample sizes, addresses, CPT codes, cross-tab consistency, and data quality. Outputs HTML reports.

## Usage

```cmd
audit filename.xlsx    # Audit a specific file
audit --all            # Audit all Excel files in current directory
audit --version        # Show version number
```

## Validation Checks

**Header Validation:**

- Ensures required headers are present and correctly named in OASCAPHS and UPLOAD tabs
- Validates header formatting and order

**Number Validation:**

- Checks that the header's submitted # matches the total of INEL, Frame INEL, and Reported rows
- Checks that Emails (E) and Mailings (M) sum to the sample size as well as the sum of rows where CMS=1
- Checks that eligible is the same as submitted - all ineligible rows
- and more!

**Data Validation:**

- Validates address fields (State, ZIP, City) for correct formatting
- Validates CPT codes against a predefined list
- Checks DOB and SERVICE DATE columns for valid date formats
- Ensures no duplicate rows based on MRN
- and more!

## TODO

- Check for "REPEAT" to the right of rows in INEL where nothing else is highlighted/marked
- Add option "--6month" to check the 6month repeat file and make sure all data is there and in the right place.
  -- When this option is used, the program should ask the user to input the directory where the 6month files are located. Then it should be saved in the .env file for future use.

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

**Build executable (for development):**

```cmd
pip install -r requirements.txt
scripts\build_exe.bat
```

Output: `dist/audit.exe`

**Create distribution package (for users):**

```cmd
scripts\package.bat
```

Output: `OAS-CAHPS-Auditor-v{VERSION}.zip` (ready to share)

**Development:**

- Edit source files: `audit.py`, `audit_lib_funcs.py`, `audit_printer.py`, `audit_report.css`
- Update version: Edit `.env` file
- Test changes: `python audit.py <file>`
- Rebuild: `scripts\build_exe.bat`

See `docs/PACKAGING_README.md` for build details.

- Rebuild: `build_exe.bat`

See `PACKAGING_README.md` for build details.
