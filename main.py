import pandas as pd
# from openpyxl import Workbook, load_workbook
import gspread
# from google_sheets import authenticate_gsheets, append_to_spreadsheet
# worksheet.format('A1:B1', {'textFormat': {'bold': True}})

# Variables
path_to_file = 'sheets//tasks.xls'
sprint_number = 39

tasks_file = pd.read_html(path_to_file)[0]  # Use [0] to get the first DataFrame returned by read_html()

# Access the 'Task' column of the DataFrame
task_column = tasks_file['Task']

# Check if 'Sprint 39' is present in any task
contains_sprint = task_column.str.contains('Sprint {:d}'.format(sprint_number))
rows_from_sprint = tasks_file[contains_sprint]

# Display the rows where 'Sprint 39' is present
print(f"Tarefas da Sprint: {sprint_number}")
print(rows_from_sprint)

credentials_path = 'credentials/credentials.json'

gc = gspread.service_account(credentials_path)
# Establish the connection
database = gc.open('Teste sprints Edenred CSC')
wks = database.worksheet('Tarefas')
data = pd.DataFrame(wks.get_all_records())
print(data)

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


print(find_related(data, 'Responsavel', 'Adriana Chaves', 'Função'))

# Input_Bitrix worksheet
input_wks = database.worksheet('Input_Bitrix')
data = pd.DataFrame(input_wks.get_all_records())
print(data)

# Clear cells except header
worksheet_range = input_wks.range("A2:ZZ1000")  

for cell in worksheet_range:
        cell.value = ""

input_wks.update_cells(worksheet_range)
input_wks.update(rows_from_sprint.fillna('').values.tolist()[1:])

formula = '=indirect("Sheet1!B2")<>"hello"'

# Output_Bitrix worksheet
output_wks = database.worksheet('Output_Bitrix')



# input_wks.update([rows_from_sprint.columns.values.tolist()] + rows_from_sprint.values.tolist())


# # spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1Pcmqd2LUlJofH9eqX6YVvBAY5S5W_1JZhzfaUpdo3YU/edit#gid=0'
# spreadsheet_url = 'https://docs.google.com/spreadsheets/d/18sleRWSJaPr_aeaz2gmxOMlRdgCECda43Gu5FLKxKE8/edit#gid=1952530042'
# client = authenticate_gsheets(credentials_path)

# # Adiciona as linhas à planilha
# append_to_spreadsheet(client, spreadsheet_url, rows_from_sprint)


# wb = load_workbook('[Edenred CSC] Controle de projeto ATIVO.xlsx')
# ws = wb.active

# print(ws)

