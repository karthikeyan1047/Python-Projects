import os
from datetime import datetime, timedelta
from openpyxl import Workbook
import _functions as cfx

def create_date_files(main_folder, start_date, end_date, extension):
    start_date = datetime.strftime(start_date, "%Y.%m.%d")
    end_date = datetime.strftime(end_date, "%Y.%m.%d")
    start = datetime.strptime(start_date, "%Y.%m.%d")
    end = datetime.strptime(end_date, "%Y.%m.%d")
    current_date = start
    while current_date <= end:
        file_name = current_date.strftime("%Y.%m.%d") + extension
        file_path = main_folder + "/" + file_name
        workbook  = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        workbook.save(file_path)
        current_date += timedelta(days=1)


main_folder = cfx.ifolder()
start_date = cfx.get_date("Start Date : ")
end_date = cfx.get_date("End Date : ")
create_date_files(main_folder, start_date, end_date, ".xlsx")
