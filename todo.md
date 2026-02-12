# TODO

## High Priority
- Add option "--6month" to check the 6month repeat file and make sure all data is there and in the right place.
  -- When this option is used, the program should ask the user to input the directory where the 6month files are located. Then it should be saved in the .env file for future use.
- Fix the bug where if there are more than 2 columns next to each other in the FRAME inel (like if you have to make a temporary column or something), the program doesn't read the columns correctly and thinks there are zero FRAME inel rows.
- Add the ability to look for header columns even if they aren't in the first row (for files where there is a title or extra info at the top), and when checking for the differences between the POP and UPLOAD tabs.

## Low Priority
- Add charts/graphs showing validation metrics using matplotlib or plotly
- Add a GUI using Tkinter or PyQt
- Add place for user to put notes or comments onto the HTML report after the audit is complete