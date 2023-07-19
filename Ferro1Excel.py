import pandas as pd
import openpyxl
from openpyxl import Workbook , load_workbook
import Poiskdata
wb = load_workbook("Svodka.xlsx", data_only=True)
sh = wb["Daily"]
i = Poiskdata.NomerStroki
cell_range1 = sh['GF'+i:'GN'+i]
cell_range2 = sh['GO'+i:'GW'+i]
cell_range3 = sh['GX'+i:'HF'+i]
cell_range4 = sh['HH'+i:'HP'+i]
cell_range5 = sh['HQ'+i:'HY'+i]
cell_range6 = sh['HZ'+i:'IH'+i]
cell_range7 = sh['IJ'+i:'IR'+i]
cell_range8 = sh['IS'+i:'JA'+i]
cell_range9 = sh['JB'+i:'JJ'+i]
cell_range10 = sh['JK'+i:'JS'+i]
DATA = sh["A"+i].value
print(DATA)
for a in range (1, 11):
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