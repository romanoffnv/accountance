from win32com.client.gencache import EnsureDispatch
import os
import re
from pprint import pprint
import pandas as pd

class ParseAcc:
    def __init__(self):
        # self.xl = xl
        # self.wb = wb
        # self.ws1 = ws1
        # Get the Excel Application COM object
        self.xl = EnsureDispatch('Excel.Application')
        self.wb = self.xl.Workbooks.Open(f"{os.getcwd()}\\acc.xls")
        # Sheets = wb.Sheets.Count
        self.ws1 = self.wb.Worksheets(1)
        
        # inserting row and copying stuff or repsonsible persons, while clearing off original column
        rangeObj = self.ws1.Range("C1:C5")
        rangeObj.EntireColumn.Insert()

    def transCol(self):
        
        row = 14
        while self.ws1.Cells(row, 2).Value != None:
            if self.ws1.Cells(row, 2).Font.Bold == False:
                # print(ws1.Cells(row, 2))
                self.ws1.Cells(row, 3).Value = self.ws1.Cells(row, 2).Value
                self.ws1.Cells(row, 2).Value = None
            row += 1
        endrow = row
        row = 14
        while row != endrow:
            if self.ws1.Cells(row, 2).Font.Bold == True and ("вич" in self.ws1.Cells(row, 2).Value or "вна" in self.ws1.Cells(row, 2).Value):
                # re.findall("(вич)|(вна)", ws1.Cells(row, 2).Value):
                print(self.ws1.Cells(row, 2).Value)
            row += 1
        self.wb.Close(True)
        self.xl.Quit()

parseAcc = ParseAcc()
transfer = parseAcc.transCol()
    
