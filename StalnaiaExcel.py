import pandas as pd
import openpyxl
from openpyxl import Workbook , load_workbook
import Poiskdata
wb = load_workbook("Svodka.xlsx", data_only=True)
sh = wb["Daily"]
i = Poiskdata.NomerStroki
cell_range1 = sh['CN'+i:'CV'+i]
cellFE = sh["CN1"].value
cell_range2 = sh['CW'+i:'DE'+i]
cellUg = sh["CW1"].value
cell_range3 = sh['DF'+i:'DN'+i]
cellKo = sh["DF1"].value
cell_range4 = sh['DQ'+i:'DY'+i]
cellLo = sh["DO1"].value
cell_range5 = sh['DZ'+i:'EH'+i]
cellRu = sh["DZ1"].value
cell_range6 = sh['EI'+i:'EQ'+i]
cellIn = sh["EI1"].value
cell_range7 = sh['ET'+i:'FB'+i]
cellCh = sh["ER1"].value
cell_range8 = sh['FC'+i:'FK'+i]
cellFob = sh["FC1"].value
cell_range9 = sh['FL'+i:'FT'+i]
cellFR = sh["FL1"].value
cell_range10 = sh['FU'+i:'GC'+i]
cellFAC = sh["FU1"].value
DATA = sh["A"+i].value
for cell1, cell2, cell3, cell4, cell5, cell6, cell7, cell8, cell9 in cell_range1:
    print(cell1.value, cell2.value, cell3.value, cell4.value, cell5.value, cell6.value, cell7.value, cell8.value, cell9.value)
PRICE1 = round(cell1.value, 2)
CH_DAY1 = round(cell2.value, 2)
CH_DAY_PR1 = round(cell3.value, 2)
CH_W1 = round(cell5.value, 2)
CH_W_PR1 = round(cell6.value, 2)
CH_M1 = round(cell8.value, 2)
CH_M_PR1 = round(cell9.value, 2)
for cell11, cell12, cell13, cell14, cell15, cell16, cell17, cell18, cell19 in cell_range2:
    print(cell11.value, cell12.value, cell13.value, cell14.value, cell15.value, cell16.value, cell17.value, cell18.value, cell19.value)
PRICE2 = round(cell11.value, 2)
CH_DAY2 = round(cell12.value, 2)
CH_DAY_PR2 = round(cell13.value, 2)
CH_W2 = round(cell15.value, 2)
CH_W_PR2 = round(cell16.value, 2)
CH_M2 = round(cell18.value, 2)
CH_M_PR2 = round(cell19.value, 2)
for cell21, cell22, cell23, cell24, cell25, cell26, cell27, cell28, cell29 in cell_range3:
    print(cell21.value, cell22.value, cell23.value, cell24.value, cell25.value, cell26.value, cell27.value, cell28.value, cell29.value)
PRICE3 = round(cell21.value, 2)
CH_DAY3 = round(cell22.value, 2)
CH_DAY_PR3 = round(cell23.value, 2)
CH_W3 = round(cell25.value, 2)
CH_W_PR3 = round(cell26.value, 2)
CH_M3 = round(cell28.value, 2)
CH_M_PR3 = round(cell29.value, 2)
for cell31, cell32, cell33, cell34, cell35, cell36, cell37, cell38, cell39 in cell_range4:
    print(cell31.value, cell32.value, cell33.value, cell34.value, cell35.value, cell36.value, cell37.value, cell38.value, cell39.value)
PRICE4 = round(cell31.value, 2)
CH_DAY4 = round(cell32.value, 2)
CH_DAY_PR4 = round(cell33.value, 2)
CH_W4 = round(cell35.value, 2)
CH_W_PR4 = round(cell36.value, 2)
CH_M4 = round(cell38.value, 2)
CH_M_PR4 = round(cell39.value, 2)
for cell41, cell42, cell43, cell44, cell45, cell46, cell47, cell48, cell49 in cell_range5:
    print(cell41.value, cell42.value, cell43.value, cell44.value, cell45.value, cell46.value, cell47.value, cell48.value, cell49.value)
PRICE5 = round(cell41.value, 2)
CH_DAY5 = round(cell42.value, 2)
CH_DAY_PR5 = round(cell43.value, 2)
CH_W5 = round(cell45.value, 2)
CH_W_PR5 = round(cell46.value, 2)
CH_M5 = round(cell48.value, 2)
CH_M_PR5 = round(cell49.value, 2)
for cell51, cell52, cell53, cell54, cell55, cell56, cell57, cell58, cell59 in cell_range6:
    print(cell51.value, cell52.value, cell53.value, cell54.value, cell55.value, cell56.value, cell57.value, cell58.value, cell59.value)
PRICE6 = round(cell51.value, 2)
CH_DAY6= round(cell52.value, 2)
CH_DAY_PR6 = round(cell53.value, 2)
CH_W6 = round(cell55.value, 2)
CH_W_PR6 = round(cell56.value, 2)
CH_M6 = round(cell58.value, 2)
CH_M_PR6 = round(cell59.value, 2)
for cell61, cell62, cell63, cell64, cell65, cell66, cell67, cell68, cell69 in cell_range7:
    print(cell61.value, cell62.value, cell63.value, cell64.value, cell65.value, cell66.value, cell67.value, cell68.value, cell69.value)
PRICE7 = round(cell61.value, 2)
CH_DAY7= round(cell62.value, 2)
CH_DAY_PR7 = round(cell63.value, 2)
CH_W7 = round(cell65.value, 2)
CH_W_PR7 = round(cell66.value, 2)
CH_M7 = round(cell68.value, 2)
CH_M_PR7 = round(cell69.value, 2)
for cell71, cell72, cell73, cell74, cell75, cell76, cell77, cell78, cell79 in cell_range8:
    print(cell71.value, cell72.value, cell73.value, cell74.value, cell75.value, cell76.value, cell77.value, cell78.value, cell79.value)
PRICE8 = round(cell71.value, 2)
CH_DAY8= round(cell72.value, 2)
CH_DAY_PR8 = round(cell73.value, 2)
CH_W8 = round(cell75.value, 2)
CH_W_PR8 = round(cell76.value, 2)
CH_M8 = round(cell78.value, 2)
CH_M_PR8 = round(cell79.value, 2)
for cell81, cell82, cell83, cell84, cell85, cell86, cell87, cell88, cell89 in cell_range9:
    print(cell81.value, cell82.value, cell83.value, cell84.value, cell85.value, cell86.value, cell87.value, cell88.value, cell89.value)
PRICE9 = round(cell81.value, 2)
CH_DAY9= round(cell82.value, 2)
CH_DAY_PR9 = round(cell83.value, 2)
CH_W9 = round(cell85.value, 2)
CH_W_PR9 = round(cell86.value, 2)
CH_M9 = round(cell88.value, 2)
CH_M_PR9 = round(cell89.value, 2)
for cell91, cell92, cell93, cell94, cell95, cell96, cell97, cell98, cell99 in cell_range10:
    print(cell91.value, cell92.value, cell93.value, cell94.value, cell95.value, cell96.value, cell97.value, cell98.value, cell99.value)
PRICE10 = round(cell91.value, 2)
CH_DAY10= round(cell92.value, 2)
CH_DAY_PR10 = round(cell93.value, 2)
CH_W10 = round(cell95.value, 2)
CH_W_PR10 = round(cell96.value, 2)
CH_M10 = round(cell98.value, 2)
CH_M_PR10 = round(cell99.value, 2)
print("Найдено")