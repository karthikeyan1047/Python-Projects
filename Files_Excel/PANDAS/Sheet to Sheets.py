import pandas as pd
import _functions as cfx
import os, shutil, sys


choice = int(cfx.inputbox(title='Method', prompt='Choose the suitable method\n   1. Single File_Single Sheet\n   2. Single File_Multiple Sheets\n   3. Multiple Files_Multiple Sheets'))

###########################################################
### SINGLE FILE WITH SINGLE SHEET - SPLIT DATA BY CATEGORY
###########################################################
if choice == 1:
    ifile_path = cfx.ifile()
    ofolder_path = os.path.join(os.path.dirname(ifile_path), "Sheet_Sheets")
    col_name = str(cfx.inputbox(title='Column Name', prompt='Enter the column Name'))

    if os.path.exists(ofolder_path):
        shutil.rmtree(ofolder_path)

    os.makedirs(ofolder_path)

    ofile_path = os.path.join(ofolder_path, "Sheets.xlsx")

    excel_file = pd.ExcelFile(ifile_path, engine='openpyxl')

    with pd.ExcelWriter(ofile_path, engine='openpyxl') as writer:
        for sht in excel_file.sheet_names:
            df = pd.read_excel(ifile_path, sheet_name=sht)
            for category in df[col_name].unique():
                newDf = df[df[col_name]==category]
                newDf.to_excel(writer, sheet_name=category, index=False)
                
    os.startfile(ofolder_path)

##########################################################################
### SINGLE FILE WITH MULTIPLE SHEETS - SPLIT AND COMBINE DATA BY CATEGORY
##########################################################################
elif choice == 2:
    input_file = cfx.ifile()
    odirectory = os.path.dirname(input_file)
    output_folder = os.path.join(odirectory, "Combined_Sheets")
    col_name = str(cfx.inputbox(title='Column Name', prompt='Enter the column Name'))

    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder)

    output_file = os.path.join(output_folder, "Combined_Categories.xlsx")

    combined_data = {}

    excel_file = pd.ExcelFile(input_file, engine='openpyxl')

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        for category in df[col_name].unique():
            category_data = df[df[col_name] == category]
            if not category in combined_data:
                combined_data[category] = []
            combined_data[category].append(category_data)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for category, data in combined_data.items():
            combined_category = pd.concat(data, ignore_index=True)
            combined_category.to_excel(writer, sheet_name=str(category), index=False)

    os.startfile(output_folder)


#############################################################################
### MULTIPLE FILES WITH MULTIPLE SHEETS - SPLIT AND COMBINE DATA BY CATEGORY
#############################################################################
elif choice == 3:
    input_folder = cfx.ifolder()
    output_folder = os.path.join(input_folder, "Combined_Sheets")
    col_name = str(cfx.inputbox(title='Column Name', prompt='Enter the column Name'))

    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder)

    output_file = os.path.join(output_folder, "Combined_Categories.xlsx")

    combined_data = {}

    for file_name in os.listdir(input_folder):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(input_folder, file_name)
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                for category in df[col_name].unique():
                    category_data = df[df[col_name] == category]
                    if not category in combined_data:
                        combined_data[category] = []
                    combined_data[category].append(category_data)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for category, data in combined_data.items():
            combined_category = pd.concat(data, ignore_index=True)
            combined_category.to_excel(writer, sheet_name=str(category), index=False)

    os.startfile(output_folder)

else:
    sys.exit(0)