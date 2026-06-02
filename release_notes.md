# Version 2.1.3 - Misc HCAHPS Fixes and Improvements

**Full Changelog**: https://github.com/ToonLunk/OAS-CAHPS-Auditor/compare/v2.1.2...v2.1.3

## What's New

- Suppressed the duplicate phone number warning in the "General Issues" section of the report since it's already covered in the Contact Lookup section
- Added more aliases for various column names
- Added check to make sure D.DATE is in ascending order
- Auditor CMD window now stays open after the audit completes if there are any warnings or errors
- Added a section for found CMS columns 
- Added AS as a required column for HCAHPS audits

---

## How do I install this?

1. Download **`OAS-CAHPS-Auditor-v2.1.3-Setup.exe`** below.
2. Run the installer - it will upgrade in place if you already have a previous version.
3. You're done! You can now start using the auditor.

Default install location: `%LOCALAPPDATA%\OAS-CAHPS-Auditor`

---

## How to Use

**Audit a single file**: hold **Shift**, right-click the Excel file, and select **"Audit this OAS file"**.

**Audit an entire folder**: hold **Shift**, right-click empty space inside the folder, and select **"Audit All OAS Files"**.

Reports are saved in an **AUDITS** folder next to your files and open automatically in your browser.

---

## SIDs

In order for the auditor to find SIDs, the SIDs.csv file must be downloaded and placed in the same folder as the auditor.

1. Hover over the blue 🛈 icon in the report, then click the link.
2. Download the SIDs.csv file and save it to your computer.
3. Move the SIDs.csv file to the same folder where the OAS-CAHPS-Auditor is installed (e.g., `%LOCALAPPDATA%\OAS-CAHPS-Auditor`).

---

## Feedback & Support

If you have any questions, run into issues, or have suggestions for improvement, please send an email to the project maintainer or submit an issue on GitHub.

---
