import time
import datetime

from google_files import get_service, get_spreadsheet

LIST_SPREADSHEET_ID = "id" #The ID is in the file URL - file containing the list of all mentors and mentees
year = '2019'

def consolidate():
    service = get_service()

    sheet_mentors_mentees = get_spreadsheet(service, LIST_SPREADSHEET_ID, year)
    sheet_mentors_mentees_content = sheet_mentors_mentees.get('values', [])

    # inc is one number above row number
    inc = 2
    # row number starts at 1 to ignore header
    for rec in sheet_mentors_mentees_content[inc-1:75]: #1 unti the last row
        print(datetime.datetime.now())
        file_name = "TRK{}".format(rec[0])

        url = rec[13]
        url_id = url.split("/d/", 1)[1]
        url_id = url_id[:-5]

        sheet_one_mentor_mentee = get_spreadsheet(service, url_id, 'Summary') #name of tab with summary in the mentor/mentees spreadsheet
        sheet_one_mentor_mentee_content = sheet_one_mentor_mentee.get('values', [])
        row_total_hours = '{}{}'.format('2019!O', inc)
        row_percentage_progress = '{}{}'.format('2019!P', inc)
        row_status = '{}{}'.format('2019!Q', inc)

        if len(sheet_one_mentor_mentee_content[12]) > 1:
            percentage_progress = sheet_one_mentor_mentee_content[12][1]
        else:
            percentage_progress = ''

        update_cell(service, LIST_SPREADSHEET_ID, row_total_hours, sheet_one_mentor_mentee_content[11][1])  #total hours
        update_cell(service, LIST_SPREADSHEET_ID, row_percentage_progress, percentage_progress)  # % progress
        update_cell(service, LIST_SPREADSHEET_ID, row_status, sheet_one_mentor_mentee_content[13][1])  # status

        inc += 1
        print("Processed {}".format(file_name))
        print(datetime.datetime.now())
        time.sleep(20)


def update_cell(service, spreadsheet_id, range, value):
    value_input_option = 'USER_ENTERED'
    value_range_body = {
        "range": "",
        "values": [[value]]
    }
    request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id,
                                                     range=range,
                                                     valueInputOption=value_input_option, body=value_range_body)
    return request.execute()


if __name__ == '__main__':
    consolidate()
