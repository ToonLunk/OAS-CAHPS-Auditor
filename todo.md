# TODO

## Technical
### High Priority
### Low Priority
- [ ] Add option to let user change the audit output directory during installation and in the settings

## Graphical

### High Priority
None at the moment

### Low Priority
- [ ] Allow user to customize the HTML report with different color schemes and layouts
- [ ] Add a GUI using Tkinter or PyQt

---

## Archive
- [x] Fix the bug where if there are more than 2 columns next to each other in the FRAME inel (like if you have to make a temporary column or something), the program doesn't read the columns correctly and thinks there are zero FRAME inel rows
- [x] Make sure to remove not just "1/1" or "5/1" from the SIDs name for the comparison, but also "1/1/26" (i.e. 3-number dates) as well.
- [x] Add None (blank string) as a valid gender.
- [x] Add "click here to download" or something like that to the message about new versions being available
- [x] Make the math error more noticable (give it a red checkmark or highlight the numbers or something)
- [x] Look for phone numbers that appear more than once in the UPLOAD tab so we can catch accidental copies of phone numbers.
- [x] Give a warning if there are more SIDs than patients with CMS=1 (if too many were added by mistake, for example). Right now it only shows if there aren't enough SIDs
- [x] Be more selective with which issues are listed if CMS=2. For example, don't show invalid addresses or telephones numbers since they are only contacted via email.
