import pandas as pd
import os, shutil
import _functions as cfx

ifolderpath = cfx.ifolder()
ofolderpath = os.path.join(ifolderpath, "Files_to_Sheets")
if os.path.exists(ofolderpath):
    shutil.rmtree(ofolderpath)
os.makedirs(ofolderpath)

combined_data = []
for files in os.listdir(ifolderpath):
    file_path = os.path.join(ifolderpath, files)
    if os.path.isfile(file_path):
        excel_file = pd.ExcelFile(file_path)
        for sht in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sht)
            combined_data.append(df)

combined_output = pd.concat(combined_data, ignore_index=True)
ofilepath = os.path.join(ofolderpath, "Combined_output.xlsx")
with pd.ExcelWriter(ofilepath, engine='openpyxl') as writer:
    combined_output.to_excel(writer,sheet_name='FullData', index=False)