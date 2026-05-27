# CAHPS Auditor

Command-line tool for auditing OAS CAHPS Excel files. Validates headers, sample sizes, addresses, CPT codes, cross-tab consistency, and data quality. Outputs HTML reports in a user-friendly format.

## Usage

**Command Line:**
```cmd
audit filename.xlsx    # Audit a specific file
audit --all            # Audit all Excel files in current directory
audit --version        # Show version number
```

**Context Menu (Reccommended):**

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

### Validation Checks

**Header Validation:**

- Ensures required headers are present and correctly named in OASCAPHS and UPLOAD tabs
- Validates header formatting and order

**Number Validation:**

- Checks that Emails (E) and Mailings (M) sum up to the sample size as well as the number of rows where CMS = 1
- Runs a number of checks to ensure that the submitted, INEL, and sample size work out logically
- and more!

**Data Validation:**

- Validates address fields (State, ZIP, City) for correct formatting and likelihood of being invalid
- Validates CPT codes against a user-customizable list of valid codes
- Checks DOB and SERVICE DATE columns for valid date ranges and formats
- Ensures no duplicate rows based on MRN
- Validates client names against a customizable list of valid client names (SIDs.csv)
- If telephone numbers or addresses are missing, populates a search query to search for their information according to CMS guidelines
- Checks SIDs for correct formatting and ranges
- and more!

## Installation

Download the latest release from the [Releases page](https://github.com/ToonLunk/OAS-CAHPS-Auditor/releases).

1. Download and run `OAS-CAHPS-Auditor-v{VERSION}-Setup.exe`
2. Follow the setup wizard
3. Verify installation by running: `audit --help` or right-clicking in a folder

The setup wizard installs to `C:\OAS-CAHPS-Auditor`, adds it to your system PATH, and optionally registers right-click context menu entries for Explorer.

**Context Menu:** During installation, you can choose to add right-click integration. This is highly recommended for ease of use - you can audit folders without opening a terminal!

**To Uninstall:** Use Add/Remove Programs (Settings → Apps → Installed Apps → CAHPS Auditor → Uninstall). This removes all files, PATH entries, and context menu registrations.

## Updating the CPT and SID Lists

The CPT and SID lists are stored in `CPT_CODES.json` and `SIDs.csv` respectively. They are stored in the default path (`%localappdata%OAS-CAHPS-Auditor`). For information on SIDs.csv, see the `About SIDs.csv.txt` file included in the distribution package. For first-time users, you will need to download `SIDs.csv` from the shared OneDrive folder and place it in the installation directory to enable SID registry lookup functionality.

If you updated these lists and later download a new version of the auditor, be sure to make a backup of your custom `CPT_CODES.csv` file before installing the new version, as the installer may overwrite it with the default version. After installing the new version, you can replace the default `CPT_CODES.csv` with your backup to retain your custom codes.

## Updating this Software

When a new update is available, you will get a notification when running the auditor. You can also check for updates manually by visiting the [Releases page](https://github.com/ToonLunk/OAS-CAHPS-Auditor/releases).

**Development:**

See [docs/PACKAGING_README.md](docs/PACKAGING_README.md) for detailed instructions on how to set up a development environment, build the executable, and create distribution packages.

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for details on changes and updates.

## ToDo

See [todo.md](todo.md) for planned features and improvements. None of these features are guaranteed to be implemented.

## License and Credit

Copyright (C) 2026 HST Pathways. All rights reserved. Developed by Tyler Brock.

This software is proprietary and confidential. Unauthorized copying, distribution, modification, or use without the express written permission of HST Pathways is strictly prohibited. This copyright notice must be preserved in all copies, modifications, and derivative works.

See [LICENSE](LICENSE) file for full legal text.
