import openpyxl

wb = openpyxl.Workbook()

sheet = wb.active

sheet.title = "test_sheet_01"

wb.save("test01.xlsx")

print(wb.sheetnames)

