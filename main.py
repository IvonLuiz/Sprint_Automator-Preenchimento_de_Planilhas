import pandas as pd
from google_sheets import authenticate_gsheets, append_to_spreadsheet

def is_word_in_text(text, word):
        for possible_name in set(findnames.findall(text)):
            if possible_name in word:
                return possible_name
        return False
    
# Variables
path_to_file = 'tasks.xls'
sprint_number = 39

tasks_file = pd.read_html(path_to_file)[0]  # Use [0] to get the first DataFrame returned by read_html()

# Access the 'Task' column of the DataFrame
task_column = tasks_file['Task']

# Check if 'Sprint 39' is present in any task
contains_sprint_39 = task_column.str.contains('Sprint {:d}'.format(sprint_number))
rows_from_sprint = tasks_file[contains_sprint_39]

# Display the rows where 'Sprint 39' is present
print(rows_from_sprint)

credentials_path = 'path/to/your/credentials.json'
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1Pcmqd2LUlJofH9eqX6YVvBAY5S5W_1JZhzfaUpdo3YU/edit#gid=0'
