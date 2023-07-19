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
cell_range9 = sh['CN'+i:'CV'+i]
cell_range10 = sh['CW'+i:'DE'+i]
cell_range11 = sh['DF'+i:'DN'+i]
cell_range12 = sh['DQ'+i:'DY'+i]
cell_range13 = sh['DZ'+i:'EH'+i]
cell_range14 = sh['EI'+i:'EQ'+i]
cell_range15 = sh['ET'+i:'FB'+i]
cell_range16 = sh['FC'+i:'FK'+i]
cell_range17 = sh['FL'+i:'FT'+i]
cell_range18 = sh['FU'+i:'GC'+i]
cell_range19 = sh['GF'+i:'GN'+i]
cell_range20 = sh['GO'+i:'GW'+i]
cell_range21 = sh['GX'+i:'HF'+i]
cell_range22 = sh['HH'+i:'HP'+i]
cell_range23 = sh['HQ'+i:'HY'+i]
cell_range24 = sh['HZ'+i:'IH'+i]
cell_range25 = sh['IJ'+i:'IR'+i]
cell_range26 = sh['IS'+i:'JA'+i]
cell_range27 = sh['JB'+i:'JJ'+i]
cell_range28 = sh['JK'+i:'JS'+i]
cell_range29 = sh['JT'+i:'KB'+i]
cell_range30 = sh['KC'+i:'KK'+i]
cell_range31 = sh['KL'+i:'KT'+i]
cell_range32 = sh['KU'+i:'LC'+i]
cell_range33 = sh['LD'+i:'LL'+i]
cell_range34 = sh['LM'+i:'LU'+i]
cell_range35 = sh['LV'+i:'MD'+i]
cell_range36 = sh['ME'+i:'MM'+i]
cell_range37 = sh['MN'+i:'MV'+i]
DATA = sh["A"+i].value
print(DATA)
for a in range (1, 38):
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