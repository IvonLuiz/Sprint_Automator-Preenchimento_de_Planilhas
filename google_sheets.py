import gspread
from oauth2client.service_account import ServiceAccountCredentials

# def authenticate_gsheets(credentials_path):
#     scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
#     creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
#     client = gspread.authorize(creds)
#     return client

# def append_to_spreadsheet(client, spreadsheet_url, data, worksheet_index=0):
#     spreadsheet = client.open_by_url(spreadsheet_url)
#     worksheet = spreadsheet.get_worksheet(worksheet_index)
#     rows_to_append = data.to_dict('records')
#     worksheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')

def clear_and_append_worksheet(worksheet, clear_matrix=['A2:ZZ300'], data_to_append=0):
        worksheet.batch_clear(clear_matrix)
        worksheet.append_rows(data_to_append)


def find_column_index(worksheet, header_name):
        df = pd.DataFrame(worksheet.get_all_records())
        return df.columns.get_loc(header_name) + 1


def apply_formula_to_column(worksheet, formula, column, start_row, end_row):
        # Apply the formula to each cell in the specified range
        for row in range(start_row, end_row + 1):
                formula_string = formula.format(column, row)
                cell_range = f"{column}{row}"
                worksheet.update(cell_range, formula_string, raw=False)


# Updated function to apply the formula to a specific column and range
def apply_formula_to_column_name(worksheet, formula, column, start_row, end_row):
        # Find the column index based on the header name
        column_index = find_column_index(worksheet, column)
        print(column_index)
        # Apply the formula to each cell in the specified range
        for row in range(start_row, end_row + 1):
                formula_string = formula.format(column_index, row)
                cell_range = f"{chr(65 + column_index - 1)}{row}"  # Convert column index to letter
                worksheet.update(cell_range, formula_string, raw=False)


def find_related(dataframe, search_column, target_word, return_column):
        contains_word = dataframe[search_column].str.contains(target_word, case=False)

        # Get the rows where the target_word is present
        rows_with_word = dataframe[contains_word]

        if not rows_with_word.empty:
                # Get the values from the specified return_column where the target_word is present
                values_from_return_column = rows_with_word[return_column]

                return values_from_return_column
        else:
                return pd.DataFrame(), None