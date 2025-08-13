import pandas as pd
import os, shutil, sys
import _functions as cfx

choice = int(cfx.inputbox(title='Method', prompt='Output as \n   1. Seperate Sheets\n   2. Combined Sheets'))


########################################################################################
### same sheet names in different files will be in a FileName_SheetName in a workbook
########################################################################################
if choice == 1:
    ifolderpath = cfx.ifolder()
    ofolderpath = os.path.join(ifolderpath, "Files_Sheets")
    if os.path.exists(ofolderpath):
        shutil.rmtree(ofolderpath)
        
    os.makedirs(ofolderpath)

    ofilepath = os.path.join(ofolderpath, "Combined_Sheets.xlsx")

    with pd.ExcelWriter(ofilepath, engine='openpyxl') as writer:
        for files in os.listdir(ifolderpath):
            file_path = os.path.join(ifolderpath, files)
            if os.path.isfile(file_path):
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                for sht in excel_file.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sht)
                    unique_sheet_name = f"{os.path.splitext(files)[0]}_{sht}"
                    df.to_excel(writer,sheet_name=unique_sheet_name, index=False)
    os.startfile(ofolderpath)
    
#####################################################################################
### if same sheet names in different files combined in a single sheet in a workbook
#####################################################################################
elif choice == 2:
    ifolderpath = cfx.ifolder()
    ofolderpath = os.path.join(ifolderpath, "Files_Sheets")
    os.makedirs(ofolderpath, exist_ok=True)
    ofilepath = os.path.join(ofolderpath, "Combined_Sheets.xlsx")

    sheet_data = {}

    for files in os.listdir(ifolderpath):
        file_path = os.path.join(ifolderpath, files)
        if os.path.isfile(file_path):
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            for sht in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sht)
                if not sht in sheet_data:
                    sheet_data[sht] = []
                sheet_data[sht].append(df)

    with pd.ExcelWriter(ofilepath, engine='openpyxl') as writer:
        for sht, data in sheet_data.items():
                combined_df = pd.concat(data, ignore_index=True)
                combined_df.to_excel(writer,sheet_name=sht, index=False)
    os.startfile(ofolderpath)
else:
    sys.exit(0)



