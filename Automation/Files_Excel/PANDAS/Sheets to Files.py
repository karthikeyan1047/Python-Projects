import pandas as pd
import _functions as cfx
import os, shutil


choice = int(cfx.inputbox(title='Method', prompt='Choose the suitable method\n   1. Single File (File) \n   2. Multiple Files (Folder) '))

if choice == 1:
    ifile_path = cfx.ifile()
    odirectory = os.path.dirname(ifile_path)
    ofolder_path = os.path.join(odirectory, "Sheets_Files")

    if os.path.exists(ofolder_path):
        shutil.rmtree(ofolder_path)
    os.makedirs(ofolder_path)

    sheet_data = {}

    if os.path.isfile(ifile_path):
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

elif choice == 2:
    ifolder_path = cfx.ifolder()
    odirectory = os.path.dirname(ifolder_path)
    ofolder_path = os.path.join(odirectory, "Sheets_Files")

    if os.path.exists(ofolder_path):
        shutil.rmtree(ofolder_path)
    os.makedirs(ofolder_path)

    sheet_data = {}

    for files in os.listdir(ifolder_path):
        ifile_path = os.path.join(ifolder_path, files)
        if os.path.isfile(ifile_path):
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