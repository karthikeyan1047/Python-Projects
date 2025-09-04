import os, time, sys
from datetime import datetime
import pandas as pd
import _functions as cfx
import xlwings as xw

def get_folder_details(main_folder, excel_output):
    if not os.path.exists(main_folder):
        print(f"The folder '{main_folder}' does not exist.")
        return
    folder_details = []
    for folder in os.listdir(main_folder):
        folder_path = os.path.join(main_folder, folder)
        if os.path.isdir(folder_path):
            created_time = os.path.getctime(folder_path)
            modified_time = os.path.getmtime(folder_path)
            folder_details.append({
                "Folder Name": folder,
                "Created Date": datetime.fromtimestamp(created_time).strftime('%Y-%m-%d %H:%M:%S'),
                "Modified Date": datetime.fromtimestamp(modified_time).strftime('%Y-%m-%d %H:%M:%S')
            })
    df = pd.DataFrame(folder_details)
    df.to_excel(excel_output, index=False)

def update_folder_details(main_folder, excel_file):
    if not os.path.exists(main_folder):
        print(f"The main folder '{main_folder}' does not exist.")
        return
    df = pd.read_excel(excel_file)
    required_columns = {"Folder Name", "Created Date", "Modified Date"}
    if not required_columns.issubset(df.columns):
        return
    for index, row in df.iterrows():
        folder_name = row["Folder Name"]
        created_date = row["Created Date"]
        modified_date = row["Modified Date"]
        try:
            created_timestamp = time.mktime(datetime.strptime(created_date, "%Y-%m-%d %H:%M:%S").timetuple())
            modified_timestamp = time.mktime(datetime.strptime(modified_date, "%Y-%m-%d %H:%M:%S").timetuple())
        except ValueError as e:
            print(f"Skipping folder '{folder_name}': Invalid date value. Error: {e}")
            continue
        folder_path = os.path.join(main_folder, str(folder_name))
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            os.utime(folder_path, (created_timestamp, modified_timestamp))
        else:
            pass


choice = int(cfx.inputbox("Select Method", "1. Get folders details\n\n2. Update folders details\n"))

if choice == 1:
    folder_path = cfx.ifolder("Get Files Details")
    excel_output = cfx.ifile("Output File")
    get_folder_details(folder_path, excel_output)
    xw.Book(excel_output)
elif choice == 2:
    folder_path = cfx.ifolder("Update Files Details")
    input_excel = cfx.ifile("Input File")
    update_folder_details(folder_path, input_excel)
    os.startfile(folder_path)
else:
    sys.exit(0)