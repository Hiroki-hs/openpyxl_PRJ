import openpyxl
import pprint

path = "D:\\sample_datas\\Shiseido_financialStatement.xlsx"
path2 = "D:\\sample_datas\\Shiseido_financialStatement2.xlsx"

wb = openpyxl.load_workbook(path)
ws1 = wb.worksheets[0]
ws2 = wb.worksheets[1]
ws3 = wb.worksheets[2]

