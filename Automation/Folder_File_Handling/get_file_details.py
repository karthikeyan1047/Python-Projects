import os
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


folder_path = cfx.ifolder()
excel_output = cfx.ifile()
get_file_details(folder_path, excel_output)
xw.Book(excel_output)