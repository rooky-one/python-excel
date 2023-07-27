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
name = (ws['A2'].value)
colour = (ws['B2'].value)

# Print something from our spreadsheet
print(ws['A2'].value)
print(ws['B2'].value)
print(f'{name}: {colour}\n')

print("<End of hello.py>")
