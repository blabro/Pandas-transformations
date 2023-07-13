import pandas as pd
import re
from datetime import datetime
now = datetime.now()

####################arkyusz 10009017
file_name = 'current.xlsx'  # path to file + file name
sheet = 'input'  # sheet name or sheet number or list of sheet numbers and names

data = pd.read_excel(io=file_name, sheet_name=sheet, skiprows=list(range(1)), index_col=0)
data.to_csv("input.txt")
f = open("input.txt", "r", encoding="utf8")

out17 = open("NA_" + now.strftime("%Y%m%d_%H_%M_%S_") + "10000006.trn", "w")
out17.write('H|' + now.strftime("%Y-%m-%dT0%H:%M:%S") + '|10000006\n')
i = 0
for line in f:
    print('R|' + line.split(',')[1] + '|||||||RefundedPoints|' + line.split(',')[2] + '|||' + line.split(',')[5])
    out17.write('R|' + line.split(',')[1] + '|||||||RefundedPoints|' + line.split(',')[2] + '|||' + line.split(',')[5]+'\n')
    i += 1

out17.write('F|'+str(i))

print('------------------END OF 10009017---------------')
