import gspread
from oauth2client.service_account import ServiceAccountCredentials

def authenticate_gsheets(credentials_path):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(creds)
    return client

def append_to_spreadsheet(client, spreadsheet_url, data, worksheet_index=0):
    spreadsheet = client.open_by_url(spreadsheet_url)
    worksheet = spreadsheet.get_worksheet(worksheet_index)
    rows_to_append = data.to_dict('records')
    worksheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
