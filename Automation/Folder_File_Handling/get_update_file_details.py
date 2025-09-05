import os, time, sys
from datetime import datetime
import pandas as pd
import _functions as cfx
import xlwings as xw

def get_file_details(folder_path, excel_output):
    if not os.path.exists(folder_path):
        print(f"The folder '{folder_path}' does not exist.")
        return
    file_details = []
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path):
            created_time = os.path.getctime(file_path)
            modified_time = os.path.getmtime(file_path)
            file_details.append({
                "File Name": file,
                "Created Date": datetime.fromtimestamp(created_time).strftime('%Y-%m-%d %H:%M:%S'),
                "Modified Date": datetime.fromtimestamp(modified_time).strftime('%Y-%m-%d %H:%M:%S')
            })
    df = pd.DataFrame(file_details)
    df.to_excel(excel_output, index=False)

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
        else:
            pass

choice = int(cfx.inputbox("Select Method", "1. Get files details\n\n2. Update files details\n"))

if choice == 1:
    folder_path = cfx.ifolder("Get Files Details")
    excel_output = cfx.ifile("Output File")
    get_file_details(folder_path, excel_output)
    xw.Book(excel_output)
elif choice == 2:
    folder_path = cfx.ifolder("Update Files Details")
    input_excel = cfx.ifile("Input File")
    update_file_details(folder_path, input_excel)
    os.startfile(folder_path)
else:
    sys.exit(0)