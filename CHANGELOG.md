# Changelog (Releases)

All notable changes to this project will be documented in this file.

## Version 1.3.4 - Contact Lookup Improvements

- Contact Lookup section now shows a name-order picker above the table — displays real patient names from the file in both orderings (as stored vs. first/last swapped) so the user can select whichever is correct; all search links update to match
- Removed FastPeopleSearch links; Contact Lookup now shows WhitePages and TruePeopleSearch only
- Minor internal fixes, including some tweaks to where certain issues get printed

## Version 1.3.3

- Minor internal fixes

## Version 1.3.2 - Reversed Name Search

- Added collapsible "Try reversed name" expander to people-search links in the Contact Lookup section, covering cases where the patient name export order is ambiguous (e.g. LAST FIRST vs FIRST LAST)
- Added "EST." to the estimated log sheet line to clarify that it's an estimate. The auditor can't read the log sheet...

## Version 1.3.1 - Facility Recognition Fix

- Fixed facility recognition logic to properly identify and match facility names in reports

## Version 1.3.0 - Email Quality Checks

- Replaced the basic email format regex with the `email-validator` library — now catches leading/trailing dots, consecutive dots, and other RFC-violating patterns that previously slipped through
- Added detection for potentially invalid emails: opt-out/placeholder local parts (e.g. `optout@`, `noreply@`, `declined@`, `test@`), single-character or all-numeric local parts (e.g. `0@gmail.com`), disposable/throwaway domains (e.g. `mailinator.com`), and very short addresses
- CMS=1 rows with potentially invalid emails are flagged in the main Issues table
- CMS=2 rows with potentially invalid emails are shown in a new collapsible section (closed by default) at the bottom of the report
- Added `email-validator` to dependencies

## Version 1.2.0 - Contact Lookup

- Added a **CONTACT LOOKUP** section to the HTML report for CMS=1 patients with contact issues. Appears automatically when issues are found — no flag or extra step needed
- Patients with no valid phone number (both fields blank, invalid, or missing) are listed with pre-built people-search links (WhitePages, TruePeopleSearch, FastPeopleSearch) that open pre-populated on click — nothing is fetched until clicked
- Patients with an invalid email address are also included with search links
- Patients who have at least one valid phone but an invalid entry in the other field are listed as reference rows (reason shown, no search links needed)
- Phone validation uses the `phonenumbers` library — a number like `1234567890` that passes format checks but is not a real assignable US number will be flagged


- CMS=2 rows no longer report invalid address or phone number issues, since these patients are contacted by email only
- Added warning when a SID is found on a non-CMS=1 row (too many SIDs added by mistake)
- Added duplicate phone number detection (flags possible copy/paste errors, noted as potentially a shared family number)
- Math error (Eligible + INEL ≠ Submitted) is now a visible ✓/✗ row in the ADDITIONAL VALIDATIONS table with highlighted numbers and an edge-case tooltip
- Blank gender is now accepted as valid (no longer flagged as an issue)
- Update badge now reads "↓ Download vX.X.X" with a tooltip clarifying it reflects the version available at audit time, not a live check
- SID registry name comparison now strips 3-part dates (e.g. 1/1/26) in addition to 2-part dates (e.g. 1/1)
- Fixed FRAME INEL count returning 0 when the patient identifier is not in column B (e.g. when a temporary column is added) — now counts any non-empty value in columns B onwards, ignoring column A where RATSTATS numbers may be placed
- Added warning when a CMS=2 row has a blank email address (email is the only contact method for these patients)
- E/M=E rows no longer report invalid address issues, and are now checked for missing email instead
- Duplicate phone detection now requires at least two CMS=1 appearances before flagging (avoids false positives when a shared number spans a CMS=1 and CMS=2 row)
- UPLOAD/OASCAPHS value-by-value comparison now flags mismatches in the row-level issues table (previously was bugged); if row counts differ, comparison is skipped to avoid a cascade

## Version 1.0.2 - More Post-Release Fixes/Additions

- Added a check to make sure the month in the filename is spelled correctly, and that the # exists in the filename

## Version 1.0.1 - Post-Release Fixes/Additions

- **IMPORTANT**: Fixed a bug where if the address column was missing/not named correctly, the report would throw an error instead of just skipping the address validation
- Added a comparison of the OASCAPHS and UPLOAD tabs to check for mismatches between the two, with a warning in the report if any are found
- Tweaked the estimated log sheet line to look cleaner and added a note about how the fields are determined
- Enchanced the alias list to catch more variations of common column names, especially for email and MRN columns
- Made address validation more robust using the `usaddress` library, which can handle a wider variety of address formats and is more accurate
- Patched the list of valid genders to include U and O
- Added more tooltips
- Added a feature to grab the facility name(s) and append them below the SID comparison

## Version 1.0.0 - The Final Release?

- Changed from installing via a batch file to using NSIS installer for a more polished installation experience
- Prompts user about context menu integration during installation, with an option to enable it
- Added an update banner to the HTML report when a new version is available because no one reads the terminal output
- Updated documentation and installation instructions
- Added an icon to the executable and installer
- Added a check to the installer for users who don't already have the SIDs.csv file, with a link to the OneDrive folder to download it

## Version 0.65.6

- The terminal and HTML report now include a link to the shared OneDrive folder for downloading SIDs.csv
- Added info tooltip next to "SID Registry Check" in reports. Contains info on how to get the SIDs.csv file and what to do if you don't have it
- Added "About SIDs.csv.txt" to the distribution package explaining where to get and install the file
- Removed hospital_names.csv and fuzzy matching feature
- Removed `thefuzz` dependency as hospital name matching is no longer needed

## Version 0.65.1

- Improved the report layout
- Added a section to show if SID registry check couldn't be performed and why

## Version 0.64.9

- Added help button

## Version 0.64.6

- Better info for adding SID registry data

## Version 0.64.5

- Improved documentation for SIDs.csv and CPT codes management and added info to the help message

## Version 0.64.4

- Skip CPT/surgical category validation when both fields are blank

## Version 0.64.3

- Added more email and service date column aliases
- MRN validation now searches more columns
- Massive performance improvements for larger files
- Added tooltips to warnings in the HTML report
- INEL tab now shows highlighted row count in reports

## Version 0.64.1

- Added service date column matching
- Finished implementing parallel processing for batch audits

## Version 0.64

- Added parallel processing to batch audits for improved performance

## Version 0.63.6

- Better SID handling and improved SID prefix extraction
- Refined SID client name comparison logic
- Report styling improvements

## Version 0.63.5

- Client name matching and summary report added
- Better registry name normalization and date formatting
- Output filenames now include months

## Version 0.63.4

- Updated installation instructions
- Minor fixes

## Version 0.63.3

- Service date validation
- Reports now auto-open after completion
- Updated HTML header display
- Expanded invalid CPT code ranges
- Show registry client name in tables instead of file client name

## Version 0.63

- Added SID registry lookup functionality (optional SIDs.csv file)
- Can now cross-reference client info with registry data

## Version 0.62

- Better error handling in audit functions
- INEL validation added
- SID prefix extraction
- Fixed version checking

## Version 0.61

- Fixed release checking
- Removed some invalid entries from cpt_codes.json

## Version 0.60

- Auto-check for updates on GitHub
- Added prompt before overwriting cpt_codes.json

## Version 0.59

- Update notifications
- Better CPT code management

## Version 0.58

- Can now run on individual files via context menu (right-click on Excel file)
- Context menu integration improvements
- Deploy script updates

## Version 0.57rc1

- Context menu integration for Windows
- Right-click an Excel file to audit it directly

## Version 0.55.1

- Simplified DOB validation (no regex): added parse_dob
- Improved handling of blank rows in CPT/SURGICAL CATEGORY check

## Version 0.54-rc1

- Began changelog file
