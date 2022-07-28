Some parts of this script are not efficient or not working 100% properly. It's worked good enough for what I've needed it to do, I don't work with javascript much, and with time being scarce, I've just left it as is, as it gets the job done.

# Requirements
- A Google workspace domain with managed chromebooks
- Privileges to access the Google Admin SDK API (Super Admin or other with appropriate permissions)

# Installation Instructions

- Create a new Google Sheet
- In the top row create the following in cell's A-P

orgUnitPath, notes, annotatedLocation, annotatedAssetId, annotatedUser, serialNumber, status, recentUsers, activeTimes, lastSync, lastEnrollmentTime, deviceId, model, platformVersion, firmwareVersion, supportEndDate

I ususally give the first row a background color and 'View > Freeze' the 1st row and up through column E. Columns A-E can be updated. F-P are data points that cannot be updated back into the admin console.

- Note the Sheet's ID
- Click Extensions > Apps Scripts
- Copy and paste the code into the Apps script, replacing the function that created by default.
- Paste your sheet's ID into the first line of the code
- If changing the name of the sheet tab, update line 2
- Save Apps Script
- Drop Down Services on the left, add the Google Sheets API and the Admin SDK API
- Back in your Sheet, navigate to Extensions > Macros > Import Macro
- Add the functions updateAdminConsole and updateSheet

updateSheet will update your Google sheet with data from the Admin console

updateAdminConsole will update the admin console with the data within the sheet - This process is currently slow and inefficient, so I would recommend deleting all rows that you aren't planning on updating.

- Navigate to Extensions > Macros > updateSheet
- Click "Continue" on authorization required prompt.
- Select your account with Admin priveleges and click Allow.
- After permissions have been given, attempt to run the updateSheet macro again and your Google sheet should start populating in chunks of 100 until all of your devices have been added to the sheet. This can take some time if you have a lot of chromebooks
- Manipulate data as you wish and if wanting to update back into the Admin Console, run the updateAdminConsole macro
