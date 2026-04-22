# Version 1.3.2 - Better Name Search 🔎

**Full Changelog**: https://github.com/ToonLunk/OAS-CAHPS-Auditor/compare/v1.3.1...v1.3.2

## What's New

- The Contact Lookup section now includes a collapsible **Try reversed** expander under each row's search links, for when the patient name export order won't work for people-search sites (e.g. if the export is LAST FIRST but the site expects FIRST LAST). Clicking it will open a new tab with the name reversed in the search query.

- Added "EST." to the sample size % to clarify that it's an estimate based on what was submitted, not directly pulling from the log sheet.
---

## How do I install this?

1. Download **`OAS-CAHPS-Auditor-v1.3.2-Setup.exe`** below.
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
    