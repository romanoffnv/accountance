from win32com.client.gencache import EnsureDispatch
import os
# Get the Excel Application COM object
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\acc.xls")
        # Sheets = wb.Sheets.Count
ws1 = wb.Worksheets(1)

row = 14
while True:
    print(ws1.Cells(row, 2).Value)
    row +=1
    if ws1.Cells(row, 2).Value == None:
        break