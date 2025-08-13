import pandas as pd
import os, shutil, sys
import _functions as cfx


choice = int(cfx.inputbox(title='Method', prompt='Choose the suitable method\n   1. Single File (File) \n   2. Multiple Files (Folder) '))
if choice == 1:
    ifile_path = cfx.ifolder()
    odirectory = os.path.dirname(ifile_path)
    ofolder_path = os.path.join(odirectory, "Sheets_Sheet")

    if os.path.exists(ofolder_path):
        shutil.rmtree(ofolder_path)

    os.makedirs(ofolder_path)

    all_data = []

    if os.path.isfile(ifile_path):
        excel_file = pd.ExcelFile(ifile_path)
        for sht in excel_file.sheet_names:
            df = pd.read_excel(ifile_path, sheet_name=sht)
            all_data.append(df)

    combined_df = pd.concat(all_data, ignore_index=True)
    output_file = os.path.join(ofolder_path, "Sheets_Sheet.xlsx")

    combined_df.to_excel(output_file, sheet_name="All_Data", index=False)

elif choice == 2:
    ifolder_path = cfx.ifolder()
    odirectory = os.path.dirname(ifolder_path)
    ofolder_path = os.path.join(odirectory, "Sheets_Sheet")

    if os.path.exists(ofolder_path):
        shutil.rmtree(ofolder_path)

    os.makedirs(ofolder_path)

    all_data = []

    for files in os.listdir(ifolder_path):
        ifile_path = os.path.join(ifolder_path, files)
        if os.path.isfile(ifile_path):
            excel_file = pd.ExcelFile(ifile_path)
            for sht in excel_file.sheet_names:
                df = pd.read_excel(ifile_path, sheet_name=sht)
                all_data.append(df)

    combined_df = pd.concat(all_data, ignore_index=True)
    output_file = os.path.join(ofolder_path, "Sheets_Sheet.xlsx")

    combined_df.to_excel(output_file, sheet_name="All_Data", index=False)