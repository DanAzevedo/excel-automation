from openpyxl import load_workbook

wb = load_workbook('exemplo.xlsx')

sheet = wb["Sheet1"]
print(sheet['A4'].value)
print(sheet)
