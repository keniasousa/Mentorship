# Mentorship

## How to Use

### Setup Google Drive Spreadsheets

1. Log in Google Drive with the account with access to the spreadsheets

2. At the Google Drive root: 

   1. Create a folder for the current year (e.g. 2020)

   2. Create a file with the spreadsheet template for the year. Reference the file "template-weekly-tracking-year.xlsx" for the format.
   
   3. Create a file with the list of mentors and mentees for the year. Reference the file "index-year.xlsx" for the format.

3. Go to https://developers.google.com/sheets/api/quickstart/python and click on the button "Enable the Google Sheets API" 

4. Select "Desktop app" to configure your OAuth client

5. Click on "Create"

6. Click on "Download Client Configuration"

7. Save the file "credentials.json" then click on "Done"

### Setup Python Script

8. Go to the command line and navigate to the folder where the scripts are located

9. Execute: pip install -r requirements.txt  (For info on how to install pip: https://www.shellhacks.com/python-install-pip-mac-ubuntu-centos/)

File: create_tracking_sheets.py: 

10. Change the ID of the file containing the full list by changing the variable 'id_spreadsheet' 

11. Change the variable 'year' with the current year, e.g. 2020. 

12. Substitute the ID of the template file by changing the variable template_spreadsheet_id. THe ID is in the file URL. 

13. Substitute the ID of the folder with the current year, e.g. "2020", by changing the variable folder_id. The ID is in the folder URL.

14. Change the IDs of the tabs in the method copy_sheet(), the third parameter. Look at gid in the file URL after selecting the tab in the template file.

15. Substitute the tab name (named '2019') and verify all the tab names (if you changed them) on the method update_cell(service,)

Location: folder where the scripts are located

16. Delete token-drive.json and token-spreadsheet.json

17. Run create_tracking_sheets.py to create these two files. To run it, go to the command line and type python create_tracking_sheets.py 

18. Run create_tracking_sheets.py to create the spreadsheets.

### Setup Consolidation Script

File: list_consolidation.py

19. Substitute the ID of the file containing the full list by changing the variable LIST_SPREADSHEET_ID 
 
20. Change the variable 'year' with the current year, e.g. 2020. 
 
21. Substitute the tab name that has the list of mentor/mentee (named '2019') on the lines reading the total hours, percentage and status and verify all the tab names (if you changed the tab named 'Summary' on the line sheet_one_mentor_mentee = get_spreadsheet()

22. Run list_consolidation.py to update the data from the spreadsheets.
