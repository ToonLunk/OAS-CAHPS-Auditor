# OAS CAHPS Auditor

Command-line tool for auditing OAS CAHPS Excel files. Validates headers, sample sizes, addresses, CPT codes, cross-tab consistency, and data quality. Outputs HTML reports.

## Usage

**Command Line:**
```cmd
audit filename.xlsx    # Audit a specific file
audit --all            # Audit all Excel files in current directory
audit --version        # Show version number
```

**Context Menu (Right-Click):**

If you installed the context menu during setup, you can:
Audit an entire folder:
- **Right-click** inside any folder → **"Audit All OAS Files"** - Audits all OAS .xlsx files in that folder
  - **Windows 10:** Regular right-click works
  - **Windows 11:** Hold **SHIFT** while right-clicking (opens extended menu), or right-click and select **"Show more options"**

Audit a single file:
- **Right-click** on any `.xlsx` file → **"Audit OAS File"** - Audits just that one file
  - **Windows 10:** Regular right-click works
  - **Windows 11:** Hold **SHIFT** while right-clicking (opens extended menu), or right-click and select **"Show more options"**

No need to use the command line at all!

## Output

This software generates an HTML report summarizing validation results, including errors and warnings found during the audit process.

Example Audit Report: [docs/SAMPLE_AUDIT.png](docs/SAMPLE_AUDIT.png)

## Validation Checks

**Header Validation:**

- Ensures required headers are present and correctly named in OASCAPHS and UPLOAD tabs
- Validates header formatting and order
- Validates sample size and other key header values
- Checks SIDs for correct formatting and ranges

**Number Validation:**

- Checks that Emails (E) and Mailings (M) sum to the sample size as well as the sum of rows where CMS=1
- Checks that eligible is the same as submitted minus all ineligible rows
- and more!

**Data Validation:**

- Validates address fields (State, ZIP, City) for correct formatting
- Validates CPT codes against a predefined list
- Checks DOB and SERVICE DATE columns for valid date formats
- Ensures no duplicate rows based on MRN
- and more!

## Installation

If you have the distribution package (ZIP file):

1. Extract the ZIP file
2. Right-click `deploy.bat` and select "Run as administrator"
3. Choose whether to install context menu integration (recommended!)
4. Restart your terminal (if using command line)
5. Verify installation by running: `audit --help` or right-clicking in a folder

The tool installs to `C:\OAS-CAHPS-Auditor` and is added to your system PATH, so you can run `audit` from anywhere.

**Context Menu:** During installation, you'll be asked if you want to add right-click integration. This is highly recommended for ease of use - you can audit folders without opening a terminal!

See `Installation Instructions.txt` in the package for detailed instructions and troubleshooting.

**To Remove Context Menu:** Run `scripts\unregister_context_menu.ps1` as administrator.

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

Output: `OAS-CAHPS-Auditor-v{VERSION}.zip` (ready to share; includes installer, powershell files, installation instructions, and license)

**Development:**

See `PACKAGING_README.md` for build details.

## TODO

High Priority:
- Add option "--6month" to check the 6month repeat file and make sure all data is there and in the right place.
  -- When this option is used, the program should ask the user to input the directory where the 6month files are located. Then it should be saved in the .env file for future use.
- Add summary report showing expected vs actual SID range at the top of the validation section for better visibility
- Fix the bug where if there are more than 2 columns next to each other in the FRAME inel (like if you have to make a temporary column or something), the program doesn't read the columns correctly and thinks there are zero FRAME inel rows.
- Add the ability to look for header columns even if they aren't in the first row (for files where there is a title or extra info at the top), and when checking
for the differences between the POP and UPLOAD tabs.


Low Priority:
- Add charts/graphs showing validation metrics using matplotlib or plotly
- Add a GUI using Tkinter or PyQt for users who prefer not to use command line
- Add place for user to put notes or comments onto the HTML report after the audit is complete

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for details on changes and updates.

## License and Credit

This software is licensed under the **GNU General Public License v3.0 (GPL-3.0)**.

Copyright © 2025 Tyler Brock.

Developed for J.L. Morgan & Associates, Inc. Code written by Tyler Brock.

See [LICENSE](LICENSE) file for full legal text or visit: https://www.gnu.org/licenses/gpl-3.0.en.html
