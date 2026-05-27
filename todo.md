# TODO

## Technical
### High Priority
- [ ] Add HCAHPS file auditing
- [ ] Add OAS SIDs to .env, and add HCAHPS SIDs
- [ ] Add validation checks for the HCAHPS-specific tabs like INEL, EXCLU, etc

### Low Priority
None at the moment

## Graphical

### High Priority
- [ ] "OAS" in the title for OAS and "HCAHPS" in the title for HCAHPS

### Low Priority
None at the moment

---

## Archive
- [x] Separate the OAS and generic functions into different files to prepare for HCAHPS auditing
- [x] Fix the bug where if there are more than 2 columns next to each other in the FRAME inel (like if you have to make a temporary column or something), the program doesn't read the columns correctly and thinks there are zero FRAME inel rows
- [x] Make sure to remove not just "1/1" or "5/1" from the SIDs name for the comparison, but also "1/1/26" (i.e. 3-number dates) as well.
- [x] Add None (blank string) as a valid gender.
- [x] Add "click here to download" or something like that to the message about new versions being available
- [x] Make the math error more noticable (give it a red checkmark or highlight the numbers or something)
- [x] Look for phone numbers that appear more than once in the UPLOAD tab so we can catch accidental copies of phone numbers.
- [x] Give a warning if there are more SIDs than patients with CMS=1 (if too many were added by mistake, for example). Right now it only shows if there aren't enough SIDs
- [x] Be more selective with which issues are listed if CMS=2. For example, don't show invalid addresses or telephones numbers since they are only contacted via email.

### Decided Against
- [ ] Allow user to customize the HTML report with different color schemes and layouts (would require too much time)
- [ ] Add a GUI using Tkinter or PyQt (would be nice but it's easier to just right-click the file)
- [ ] Add option to let user change the audit output directory during installation and in the settings (for now it's just in %LOCALAPPDATA%, which is fine. Code is already there to support this if I want to add it later though.)
