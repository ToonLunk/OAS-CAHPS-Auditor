# Version 1.3.5 - Facility Name Fixes & Report Polish

**Full Changelog**: https://github.com/ToonLunk/OAS-CAHPS-Auditor/compare/v1.3.4...v1.3.5

## What's New

- Fixed a bug where the Facility/Location columns section was only showing 1 result (or sometimes none). The auditor was incorrectly treating common words inside column headers as stop signals — for example, `"id"` was matching inside `"provider"` and `"residential"`. This is now fixed.

- The Facility/Location columns section now always appears in the report, even when SID lookup fails.

- Added support for additional facility column name aliases: `agency name`, `client id`, `revenue location` (and underscore/no-space variants).

- "Possible" and "Potentially" flagged issues now appear with a subtle yellow background in the Issues table, making them easier to distinguish from hard errors at a glance.

- The CONTACT LOOKUP section heading no longer appears when there are no contact issues to show.

- Placeholder name flag wording softened — now reads "Possible Placeholder Name" with a note to verify, rather than treating it as a definite error.

---

## How do I install this?

1. Download **`OAS-CAHPS-Auditor-v1.3.5-Setup.exe`** below.
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

# Version 1.3.4 - Contact Lookup Tweaks

**Full Changelog**: https://github.com/ToonLunk/OAS-CAHPS-Auditor/compare/v1.3.3...v1.3.4

## What's New

- The Contact Lookup section now shows a name-order picker above the table. It displays 2–3 real patient names from the file in both orderings so you can click whichever looks correct — all search links update instantly to match.

- Removed FastPeopleSearch links. Links are now WhitePages and TruePeopleSearch only.

- The auditor now checks to make sure the year in the filename matches the year(s) found in the patient data.

---

## How do I install this?

1. Download **`OAS-CAHPS-Auditor-v1.3.4-Setup.exe`** below.
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
    