# Hubitat-Sheets Lock Automation for Short-Term Rentals
Lock automation for short-term rentals using Hubitat and Google Sheets.

1. Set up your locks in Hubitat.
2. Install "Maker API" with remote access on your Hubitat.
3. Create a new Google Sheet.
4. Rename the individual tab to match the name of your lock device.
5. Add additional tabs for additional locks.
6. Create an Apps Script by clicking "Exetensions --> Apps Script"
7. Copy and paste this code over the Code.gs file.
8. Replace HUBITAT_ACCESS_TOKEN and HUBITAT_URL_STUB with entries from the Maker API setup screen.
9. In the PROPERTIES constant below, enter the Google Sheets tab names in the 'sheetname' field and
   links to your .ics calendars in the 'link' field.
10. Save this script.
11. Go back to the Google Sheet and refresh.
12. On each lock tab, clieck "Rental Lock Automator --> Initialize Sheet"
13. Add permanent codes to your locks with the word "permanent" (all lower case) in the "Type" column.
14. Test various functions by running them in Apps Script: updateCalendars() and updateLockStatus() and reviewSheetsForChanges()
15. Once you're confident everything works, add automation triggers: Run everyMinute() on a time-based trigger. Run onChange() on edits.
