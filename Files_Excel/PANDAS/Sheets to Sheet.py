import pandas as pd
import os, shutil, sys
import _functions as cfx

ifile_path = cfx.ifile()
odirectory = os.path.dirname(ifile_path)
ofolder_path = os.path.join(odirectory, "Sheets_Sheet")

if os.path.exists(ofolder_path):
    shutil.rmtree(ofolder_path)

os.makedirs(ofolder_path)

all_data = []
excel_file = pd.ExcelFile(ifile_path)

for sht in excel_file.sheet_names:
    df = pd.read_excel(ifile_path, sheet_name=sht)
    all_data.append(df)

combined_df = pd.concat(all_data, ignore_index=True)
output_file = os.path.join(ofolder_path, "Sheets_Sheet.xlsx")
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    combined_df.to_excel(writer, sheet_name="All_Data", index=False)