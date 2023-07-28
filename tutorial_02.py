from openpyxl.workbook import Workbook
from openpyxl import load_workbook
#  openpyxl ia a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files

print(f"<Start of tutorial_01.py>\n")
# Create workbook object
# wb = Workbook()

# load existing spreadsheet
wb = load_workbook('tutorial_01.xlsx')

# Create an active worksheet
ws = wb.active

# Lesson 12 - Change Cells and Save Spreadsheet
print(f"Lesson 12 - Change Cells and Save Spreadsheet\n")

# Change a cell

ws["A11"] = "Judas"

# Save the spreadsheet (two options)
# wb.save('C:\code\python-excel\tutorial.xlsx')  # Absolute path
wb.save('tutorial_02.xlsx')  # Relative save, note save under new file name

print("File was saved...")
