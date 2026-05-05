# Version 1.4.0 - Report Improvements

**Full Changelog**: https://github.com/ToonLunk/OAS-CAHPS-Auditor/compare/v1.3.5...v1.4.0

## What's New

- The VALIDATION SUMMARY is now collapsible — collapsed automatically when everything passes, open when there are failures
- Age validation improved: ages ≤ 0, under 18, or over 110 are now flagged separately with distinct issue types
- Console warning shown on startup if cpt_codes.json can't be loaded

---

## How do I install this?

1. Download **`OAS-CAHPS-Auditor-v1.4.0-Setup.exe`** below.
2. Run the installer - it will upgrade in place if you already have a previous version.
3. You're done! You can now start using the auditor.

Default install location: `C:\OAS-CAHPS-Auditor`

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
3. Move the SIDs.csv file to the same folder where the OAS-CAHPS-Auditor is installed (e.g., `C:\OAS-CAHPS-Auditor`).

---

## Feedback & Support

If you have any questions, run into issues, or have suggestions for improvement, please send an email to the project maintainer or submit an issue on GitHub.

---