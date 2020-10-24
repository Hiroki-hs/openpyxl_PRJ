import openpyxl as xp

wb = xp.load_workbook("kinou.xlsx")


def copy_and_paste(sheetname, num_min_row, num_max_row, num_min_column, num_max_column, sheetname_paste):
    """
    docstring
    """
    ws = wb[str(sheetname)]
    paste_ws = wb[sheetname_paste]
    
    for j in range(num_min_column, num_max_column):
        for i in range(num_min_row, num_max_row):
            copy = ws.cell(row=i, column=j).value
            print("セル" + str(i) + "," + str(j) + ":" + str(copy))
            
            i += 1
        j += 1


#copy_and_paste("Sheet1", 3, 27, 2, 6)
copy_and_paste("Sheet2", 1, 15, 1, 6, "Sheet3")

wb.save("kinou3.xlsx")
