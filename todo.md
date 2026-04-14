# TODO

## Technical
### High Priority
- Be more selective with which issues are listed if CMS=2. For example, don't show invalid addresses or telephones numbers since they are only contacted via email.
- Give a warning if there are more SIDs than patients with CMS=1 (if too many were added by mistake, for example). Right now it only shows if there aren't enough SIDs
- Look for phone numbers that appear more than once in the UPLOAD tab so we can catch accidental copies of phone numbers. This will be a false positive if a family member shares a phone or something, but it's worth noting
- Make the math error more noticable (give it a red checkmark or highlight the numbers or something)
- Add None (blank string) as a valid gender.

### Low Priority
- Add option to let user change the audit output directory during installation and in the settings
- Fix the bug where if there are more than 2 columns next to each other in the FRAME inel (like if you have to make a temporary column or something), the program doesn't read the columns correctly and thinks there are zero FRAME inel rows

## Graphical

### High Priority
None at the moment

### Low Priority
- Allow user to customize the HTML report with different color schemes and layouts
- Add "click here to download" or something like that to the message about new versions being available
- Add a GUI using Tkinter or PyQt
