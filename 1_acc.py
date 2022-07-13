from win32com.client.gencache import EnsureDispatch
import os
import re
from pprint import pprint
import pandas as pd
from functools import reduce

class ParseAcc:
    def __init__(self):
        # Get the Excel Application COM object
        self.xl = EnsureDispatch('Excel.Application')
        self.wb = self.xl.Workbooks.Open(f"{os.getcwd()}\\acc.xls")
        self.ws1 = self.wb.Worksheets(1)
        

    def transCol(self):
        # getting mols, items with breaks (for mols) into lists
        row = 14
        L_mols_scratch, L_items = [], []
        while True:
            if self.ws1.Cells(row, 2).Font.Bold == True and ("вич" in self.ws1.Cells(row, 2).Value or "вна" in self.ws1.Cells(row, 2).Value):
                print(self.ws1.Cells(row, 2).Value)
                L_mols_scratch.append(self.ws1.Cells(row, 2).Value)
                L_items.append('****')
            elif self.ws1.Cells(row, 2).Font.Bold != True:
                L_items.append(self.ws1.Cells(row, 2).Value)
            if self.ws1.Cells(row, 2).Value == None:
                break
            row += 1
        
        L_items = L_items[1:]
        L_items.append('****')
        
        L_counts = []
        counter = 0
        
        for i in L_items:
            if i == '****':
                L_items.remove(i)
                L_counts.append(counter)
                counter = 0
            counter += 1    
                # continue
        # L_counts = [int(x) - 1 for x in L_counts if x == L_counts[-1:]]
        
        sumL_counts = reduce(lambda x, y: x + y, L_counts)
        
        # print(L_items[:15])
        # print(len(L_items))
        print(f'this is mols len {len(L_mols_scratch)}')
        print(f'this is items len {len(L_items)}')
        print(f'this is sum of items  {sumL_counts}')
        print(f'this is the list of ranges {L_counts}')
        # the last item of the L_counts list is 24, when should be 23 - problem needs to be addressed
        print(f'this is the length of ranges list {len(L_counts)}')
        
        self.wb.Close(True)
        self.xl.Quit()
        
        # after mols_scratch, items and counts lists are collected, L_mols should be populated by
        # mulitplying each mols_scratch element by counts
        L_mols = [((i + '**').split('**') * j).filter('', L_mols) for i, j in (zip(L_mols_scratch, L_counts))]
        print(L_mols)

parseAcc = ParseAcc()
transfer = parseAcc.transCol()
    
