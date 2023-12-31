import pandas as pd
import openpyxl
from openpyxl import Workbook , load_workbook
import Poiskdata
wb = load_workbook("Svodka.xlsx", data_only=True)
sh = wb["Daily"]
i = Poiskdata.NomerStroki
cell_range1 = sh['B'+i:'J'+i]
cell_range2 = sh['K'+i:'S'+i]
cell_range3 = sh['AL'+i:'AT'+i]
cell_range4 = sh['AU'+i:'BC'+i]
cell_range5 = sh['BD'+i:'BL'+i]
cell_range6 = sh['BM'+i:'BU'+i]
cell_range7 = sh['BV'+i:'CD'+i]
cell_range8 = sh['CE'+i:'CM'+i]
DATA = sh["A"+i].value
print(DATA)
for a in range (1, 9):
    for cell1, cell2, cell3, cell4, cell5, cell6, cell7, cell8, cell9 in globals()['cell_range%s' % a]:
        print(cell1.value, cell2.value, cell3.value, cell4.value, cell5.value, cell6.value, cell7.value, cell8.value, cell9.value)
    globals()['PRICE%s' % a] = round(cell1.value, 2)
    globals()['CH_DAY%s' % a] = round(cell2.value, 2)
    globals()['CH_DAY_PR%s' % a] = round(cell3.value, 2)
    globals()['CH_W%s' % a] = round(cell5.value, 2)
    globals()['CH_W_PR%s' % a] = round(cell6.value, 2)
    globals()['CH_M%s' % a] = round(cell8.value, 2)
    globals()['CH_M_PR%s' % a] = round(cell9.value, 2)
print("Good")