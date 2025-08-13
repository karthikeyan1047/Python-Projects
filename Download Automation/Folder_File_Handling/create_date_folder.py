import os
from datetime import datetime, timedelta
import _functions as cfx

def create_date_folders(main_folder, start_date, end_date):
    start_date = datetime.strftime(start_date, "%Y.%m.%d")
    end_date = datetime.strftime(end_date, "%Y.%m.%d")
    start = datetime.strptime(start_date, "%Y.%m.%d")
    end = datetime.strptime(end_date, "%Y.%m.%d")
    current_date = start
    while current_date <= end:
        folder_name = current_date.strftime("%Y.%m.%d")
        folder_path = os.path.join(main_folder, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        current_date += timedelta(days=1)
    
main_folder = cfx.ifolder()
start_date = cfx.get_date("Start Date : ")
end_date = cfx.get_date("End Date : ")

create_date_folders(main_folder, start_date, end_date)

