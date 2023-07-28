from openpyxl.workbook import Workbook
# from openpyxl import load_workbook
#  openpyxl ia a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files

print(f"<Start of tutorial_03.py>\n")

# Lesson 15 - Create A Spreadsheet Workbook
print(f"Lesson 15 - Create A Spreadsheet Workbook\n")

# Create workbook object
wb = Workbook()

# Create an active worksheet
ws = wb.active

# Create worksheet title
ws.title = "Names and Colours"

# Save Spreadsheet
wb.save('tutorial_03.xlsx')

print(f'\n{"<End of tutorial_03.py>"}')  # Add newline to start line


