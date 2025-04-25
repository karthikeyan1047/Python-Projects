import os
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

main_folder = cfx.ifolder()
excel_output = cfx.ifile()
get_folder_details(main_folder, excel_output)
xw.Book(excel_output)