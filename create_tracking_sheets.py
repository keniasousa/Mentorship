import time
import datetime

from google_files import get_service, get_drive_service, get_spreadsheet, create_spreadsheet


def main():
    service = get_service()
    drive_service = get_drive_service()


    year = '2020'
    id_spreadsheet = 'id' #The ID is in the file URL - file containing the list of all mentors and mentees

    sheet_mentors_mentees = get_spreadsheet(service, id_spreadsheet, year) 
    sheet_mentors_mentees_content = sheet_mentors_mentees.get('values', [])

    template_spreadsheet_id = 'id' # ID of the template spreadsheet
    folder_id = 'id' # ID of the folder

    inc = 2 #first row, excluding header
    for rec in sheet_mentors_mentees_content[inc-1:77]: #1 unti the last row
        print(datetime.datetime.now())
        file_name = "TRK{}".format(rec[0])
        body = {
          "properties": {
            "title": file_name
          }
        }

        new_spreadsheet = create_spreadsheet(service, body)

        copy_sheet(service, template_spreadsheet_id, 0, 'Weekly log', new_spreadsheet.get('spreadsheetId', '')) #Weekly log
        copy_sheet(service, template_spreadsheet_id, 0, 'Summary', new_spreadsheet.get('spreadsheetId', '')) #Summary
        copy_sheet(service, template_spreadsheet_id, 0, 'Guidelines', new_spreadsheet.get('spreadsheetId', '')) #Guidelines

        delete_sheet(service, new_spreadsheet.get('spreadsheetId', ''))

        mentor_full_name = '{} {}'.format(rec[5], rec[6])
        mentee_full_name = '{} {}'.format(rec[1], rec[2])

        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B3', rec[0]) #id
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B4', mentor_full_name) #mentor name
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B5', rec[7])  #mentor email
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B6', rec[8])  # mentor phone
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B7', mentee_full_name)  # mentee name
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B8', rec[3])  # mentee email
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B9', rec[4])  # mentee phone
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B10', rec[9])  # sector
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B12', "='Weekly log'!C33")  # total hours
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B13', "=INDEX('Weekly log'!D: D, COUNTA('Weekly log'!D: D), 1)")  # percentage progress
        update_cell(service, new_spreadsheet.get('spreadsheetId', ''), 'Summary!B14',"=INDEX('Weekly log'!E:E,COUNTA('Weekly log'!E:E),1)")  # status
        update_cell(service, id_spreadsheet, '2019!N{}'.format(inc), new_spreadsheet.get('spreadsheetUrl', ''))

        move_to_folder(drive_service, new_spreadsheet.get('spreadsheetId', ''), folder_id)
        share_spreadsheet(drive_service, new_spreadsheet.get('spreadsheetId', ''), rec[7]) #email mentor
        share_spreadsheet(drive_service, new_spreadsheet.get('spreadsheetId', ''), rec[3])  # email mentee

        inc += 1
        print("Processed {}".format(file_name))
        print(datetime.datetime.now())
        time.sleep(20)


def copy_sheet(service, origin_id, origin_sheet_id, origin_sheet, destination_id):
    copy_sheet_to_another_spreadsheet_request_body = {
        'destinationSpreadsheetId': destination_id
    }

    request = service.spreadsheets().sheets().copyTo(spreadsheetId=origin_id, sheetId=origin_sheet_id,
                                                     body=copy_sheet_to_another_spreadsheet_request_body)

    response = request.execute()
    sheet_id = response.get('sheetId', '')
    update_sheet_name_body = {
      "requests": [
        {
          "updateSheetProperties": {
            "properties": {
              "sheetId": sheet_id,
              "title": origin_sheet
            },
            "fields": "title"
          }
        }
      ]
    }
    request = service.spreadsheets().batchUpdate(spreadsheetId=destination_id, body=update_sheet_name_body)
    return request.execute()


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


def delete_sheet(service, spreadsheet_id):
    delete_sheet_body = {
        "requests": [
            {
                "deleteSheet": {
                    "sheetId": 0
                }
            }
        ]
    }
    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=delete_sheet_body)
    return request.execute()


def move_to_folder(service, spreadsheet_id, folder_id):
    file = service.files().get(fileId=spreadsheet_id,
                               fields='parents').execute()

    previous_parents = ",".join(file.get('parents'))
    service.files().update(fileId=spreadsheet_id,
                           addParents=folder_id,
                           removeParents=previous_parents,
                           fields='id, parents').execute()


def share_spreadsheet(service, spreadsheet_id, email):
    def callback(request_id, response, exception):
        if exception:
            print(exception)
        else:
            print("Permission Id: %s" % response.get('id'))

    batch = service.new_batch_http_request(callback=callback)
    user_permission = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': email
    }
    batch.add(service.permissions().create(
        fileId=spreadsheet_id,
        body=user_permission,
        fields='id',
    ))
    batch.execute()


if __name__ == '__main__':
    main()
