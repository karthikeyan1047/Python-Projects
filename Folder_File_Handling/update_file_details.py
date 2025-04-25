import pandas as pd
import os
import time
from datetime import datetime
import _functions as cfx

def update_file_details(main_folder, excel_file):
    if not os.path.exists(main_folder):
        print(f"The main folder '{main_folder}' does not exist.")
        return
    df = pd.read_excel(excel_file)
    required_columns = {"File Name", "Created Date", "Modified Date"}
    if not required_columns.issubset(df.columns):
        print(f"The Excel file must contain these columns: {required_columns}")
        return
    for index, row in df.iterrows():
        file_name = row["File Name"]
        created_date = row["Created Date"]
        modified_date = row["Modified Date"]
        created_timestamp = time.mktime(datetime.strptime(created_date, "%Y-%m-%d %H:%M:%S").timetuple())
        modified_timestamp = time.mktime(datetime.strptime(modified_date, "%Y-%m-%d %H:%M:%S").timetuple())
        file_path = os.path.join(main_folder, file_name)
        if os.path.exists(file_path) and os.path.isfile(file_path):
            os.utime(file_path, (created_timestamp, modified_timestamp))
            print(f"Updated file: {file_name}")
        else:
            print(f"File '{file_name}' does not exist in '{main_folder}'.")

main_folder = cfx.ifolder()
excel_file = cfx.ifile()
update_file_details(main_folder, excel_file)
os.startfile(main_folder)