from openpyxl import Workbook, load_workbook

wb = load_workbook(filename="Allocation Sorting.xlsx")

sheet = wb.active

for table in sheet:
    print(type(table))
    