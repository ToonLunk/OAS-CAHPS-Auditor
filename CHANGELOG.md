# Changelog (Releases)

All notable changes to this project will be documented in this file.

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
