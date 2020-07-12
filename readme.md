# Mentorship

## How to Use

### Setup Google Drive Spreadsheets

1. Log in Google Drive with the account with access to the spreadsheets

2. At the Google Drive root: 

   1. Create a folder for the current year (e.g. 2019)

   2. Create a file with the spreadsheet template for the year: e.g. Template-Weekly-Tracking-2019

3. Go to https://developers.google.com/sheets/api/quickstart/python and click on the button "Enable the Google Sheets API" 

4. Select "Desktop app" to configure your OAuth client

5. Click on "Create"

6. Click on "Download Client Configuration"

7. Save the file "credentials.json" then click on "Done"

### Setup Python Script

6. Go to the command line and navigate to the folder where the scripts are located

7. Execute: pip install -r requirements.txt 

8. On the file create_tracking_sheets.py: Change the ID of the file containing the full list and the tab name on by changing the variables 'id_spreadsheet' and 'year'. 
sheet_mentors_mentees = get_spreadsheet(service, '(id_spreadsheet)', year)

9. On the file create_tracking_sheets.py: Substitute the ID of the template file by changing the variable template_spreadsheet_id and substitute the ID of the folder for the year "2019" by changing the variable folder_id. The ID is in the file URL.

10. On the file create_tracking_sheets.py: Change the IDs of the tabs in the method copy_sheet().

11. Substitute the tab name (named '2019') and verify all the tab names on the method update_cell(service,)

12. Delete token-drive.json and token-spreadsheet.json

13. Run create_tracking_sheets.py to create these two files.

14. Run create_tracking_sheets.py to create the spreadsheets.

### Setup Consolidation Script

File: list_consolidation.py

1. Substitute the ID of the file containing the full list by changing the variable 
LIST_SPREADSHEET_ID 
 
2. Substitute the tab name on by changing the variable 'year'
 
3. Substitute the tab name with the summary in the mentor/mentees spreadsheet (named 'Summary') on the line sheet_one_mentor_mentee = google_files.get_spreadsheet

4. Substitute the tab name that has the list of mentor/mentee (named '2019') on the lines reading the total hours, percentage and status.

5. Run list_consolidation.py to update the data from the spreadsheets.
