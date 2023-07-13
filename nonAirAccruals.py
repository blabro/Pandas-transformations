import pandas as pd
import re
from datetime import datetime
now = datetime.now()

file_name = 'current.xlsx'  # path to file + file name


def czytaj_plik(sheetExela):
    sheet = sheetExela  # sheet name or sheet number or list of sheet numbers and names
    return pd.read_excel(io=file_name, sheet_name=sheet, skiprows=list(range(1)), index_col=0)


def przerob_pliki(sheetExela, plikTxt):

    data = czytaj_plik(sheetExela)

    data.to_csv(plikTxt)
    with open(plikTxt, "r", encoding="utf8") as f:
        partner_code = sheetExela[6:14]
        # print(partnerCode)
        with open("NA_" + now.strftime("%Y%m%d_%H_%M_%S_") + partner_code + ".trn", "w") as out:
            out.write('H|' + now.strftime("%Y-%m-%dT0%H:%M:%S") + '|' + partner_code + '\n')
            i = 0
            for line in f:
                # print('R|' + line.split(',')[1] + '||||||||' + line.split(',')[2] + '|||' + line.split(',')[4])
                out.write('R|' + line.split(',')[1] + '||||||||' + line.split(',')[2] + '|||' + line.split(',')[4]+'\n')
                i += 1

            out.write('F|'+str(i))


przerob_pliki('Input 10009017', "in17.txt")
# przerob_pliki('Input 10009018', "in18.txt")
