# OAS CAHPS Auditor

Command-line tool for auditing OAS CAHPS Excel files. Validates headers, sample sizes, addresses, CPT codes, cross-tab consistency, and data quality. Outputs HTML reports in a user-friendly format.

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

- Checks that Emails (E) and Mailings (M) sum up to the sample size as well as the sum of rows where CMS = 1
- Checks that eligible is the same as submitted minus all ineligible rows
- and more!

**Data Validation:**

- Validates address fields (State, ZIP, City) for correct formatting
- Validates CPT codes against a customizable list of valid codes
- Checks DOB and SERVICE DATE columns for valid date formats
- Ensures no duplicate rows based on MRN
- Validates client names against a customizable list of valid client names (SIDs.csv)
- and more!

## Installation

Download the latest release from the [Releases page](https://github.com/ToonLunk/OAS-CAHPS-Auditor/releases).

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

## Updating the CPT and SID Lists

The CPT and SID lists are stored in `CPT_CODES.json` and `SIDs.csv` respectively. You can typically find these files in the default installation folder, which is `C:\OAS-CAHPS-Auditor`. You can edit these files with any text editor or spreadsheet software to add/remove valid codes and client names. The auditor will use these updated lists during validation.

If you updated these lists and later download a new version of the auditor, be sure to make a backup of your custom `CPT_CODES.csv` and `SIDs.csv` files before updating, as the installer may overwrite them with the default versions. Then simply copy your custom files back into the installation directory after updating.

## Updating this Software

When a new update is available, you will get a notification when running the auditor. You can also check for updates manually by visiting the [Releases page](https://github.com/ToonLunk/OAS-CAHPS-Auditor/releases).

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

Output: `OAS-CAHPS-Auditor-v{VERSION}.zip` (includes installer, powershell files, installation instructions, CPT and SID lists, and license)

**Development:**

See [docs/PACKAGING_README.md](docs/PACKAGING_README.md) for detailed instructions on how to set up a development environment, build the executable, and create distribution packages.

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for details on changes and updates.

## ToDo

See [todo.md](todo.md) for planned features and improvements. None of these features are guaranteed to be implemented.

## License and Credit

This software is licensed under the **GNU General Public License v3.0 (GPL-3.0)**.

Copyright © 2026 Tyler Brock.

Developed for J.L. Morgan & Associates, Inc. Code written by Tyler Brock.

See [LICENSE](LICENSE) file for full legal text or visit: https://www.gnu.org/licenses/gpl-3.0.en.html
