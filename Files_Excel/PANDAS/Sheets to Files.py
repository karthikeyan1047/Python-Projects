import pandas as pd
import _functions as cfx
import os, shutil

ifile_path = cfx.ifile()
odirectory = os.path.dirname(ifile_path)
ofolder_path = os.path.join(odirectory, "Sheets_Files")

if os.path.exists(ofolder_path):
    shutil.rmtree(ofolder_path)
os.makedirs(ofolder_path)

sheet_data = {}

excel_file = pd.ExcelFile(ifile_path)
for sht in excel_file.sheet_names:
    df = pd.read_excel(ifile_path, sheet_name=sht)
    if not sht in sheet_data:
        sheet_data[sht] = []
    sheet_data[sht].append(df)

for sht, data in sheet_data.items():
    combined_df = pd.concat(data, ignore_index=True)
    outputfiles = os.path.join(ofolder_path, f"{sht}.xlsx")
    combined_df.to_excel(outputfiles, sheet_name=sht, index=False)