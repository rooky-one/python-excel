from openpyxl.workbook import Workbook
from openpyxl import load_workbook

print(f"<Start of hello.py>\n")
# Create workbook object
# wb = Workbook()

# load existing spreadsheet
wb = load_workbook('hello.xlsx')

# Create an active worksheet
ws = wb.active

# Use a variable to hold spreadsheet call value
name = ws['A2'].value
colour = ws['B2'].value

# Print something from our spreadsheet
print(ws['A2'].value)
print(ws['B2'].value)
print(f'{name}: {colour}\n')

# Grab a whole column
column_a = ws['A']  # Returns immutable tuple
column_b = ws['B']  # Returns immutable tuple

print(f'{column_a} \n')  # Prints the immutable tuple list
print(f'{column_b} \n')  # Prints the immutable tuple list

# Loop through the tuple
for spreadsheet_cells in column_a:
    print(spreadsheet_cells.value)  # Print tuple contents as list

print('')

# Loop through the tuple
for spreadsheet_cells in column_b:
    print(spreadsheet_cells.value)  # Print tuple contents as list

print(f'\n{"<End of hello.py>"}')  # Add newline to start line
