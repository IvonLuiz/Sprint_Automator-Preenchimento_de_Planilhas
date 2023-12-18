import pandas as pd
import gspread
from google_sheets import *
# worksheet.format('A1:B1', {'textFormat': {'bold': True}})

# Variables
path_to_file = 'sheets//tasks.xls'
sprint_number = 35

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

print(find_related(data, 'Responsavel', 'Adriana Chaves', 'Função'))

# Input_Bitrix worksheet

input_wks = database.worksheet('Input_Bitrix')

# Clear cells except header (Adjust if necessary)
clear_and_append_worksheet(input_wks, ['A2:ZZ300'], rows_from_sprint.fillna('').values.tolist()[1:])


# Output_Bitrix worksheet



output_wks = database.worksheet('Output_Bitrix')

# Update the worksheet with the formula in column A from row 2 to the end of rows_from_sprint
apply_formula_to_column(worksheet = output_wks, 
                        formula = "=EXT.TEXTO(Input_Bitrix!B{1};23;2)",
                        column = 'A', 
                        start_row = 2, 
                        end_row = len(rows_from_sprint))

apply_formula_to_column(worksheet = output_wks, 
                        formula = "=Input_Bitrix!A{1}",
                        column = 'B', 
                        start_row = 2, 
                        end_row = len(rows_from_sprint))

apply_formula_to_column_name(worksheet = output_wks, 
                        formula = "=EXT.TEXTO(Input_Bitrix!B{1};27;5000)",
                        column = 'Task', 
                        start_row = 2, 
                        end_row = len(rows_from_sprint))

apply_formula_to_column_name(worksheet = output_wks, 
                        formula = "=Input_Bitrix!G{1}",
                        column = 'Responsavel', 
                        start_row = 2, 
                        end_row = len(rows_from_sprint))


# size = len(rows_from_sprint)
# # Assuming 'Sprint' is a column header in your worksheet
# sprint_column_header = 'Sprint'
# sprint_values = [sprint_number] * len(rows_from_sprint)

# # Find the column index corresponding to the 'Sprint' header
# sprint_column_index = output_wks.columns.get_loc(sprint_column_header) + 1  # Adding 1 to convert from 0-based to 1-based index

# # Update the entire column with the sprint value
# input_wks.update_cell(1, sprint_column_index, sprint_values)

# # Add formula to column A for each row (starting from the second row)
# formula_range = output_wks.range('A2:A' + str(len(data_to_append) + 1))
# formula = '=EXT.TEXTO(Input_Bitrix!B2;23;2)'  # Your specified formula
# for cell in formula_range:
#         cell.value = formula

# output_wks.update_cells(formula_range)



# input_wks.update([rows_from_sprint.columns.values.tolist()] + rows_from_sprint.values.tolist())


# # spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1Pcmqd2LUlJofH9eqX6YVvBAY5S5W_1JZhzfaUpdo3YU/edit#gid=0'
# spreadsheet_url = 'https://docs.google.com/spreadsheets/d/18sleRWSJaPr_aeaz2gmxOMlRdgCECda43Gu5FLKxKE8/edit#gid=1952530042'
# client = authenticate_gsheets(credentials_path)

# # Adiciona as linhas à planilha
# append_to_spreadsheet(client, spreadsheet_url, rows_from_sprint)


# wb = load_workbook('[Edenred CSC] Controle de projeto ATIVO.xlsx')
# ws = wb.active

# print(ws)

