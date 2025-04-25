import os
import pandas as pd
import time
from datetime import datetime
import _functions as cfx
import xlwings as xw

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
            print(f"Folder '{folder_name}' does not exist in '{main_folder}'.")

main_folder = cfx.ifolder()
excel_file = cfx.ifile()
update_folder_details(main_folder, excel_file)
os.startfile(main_folder)