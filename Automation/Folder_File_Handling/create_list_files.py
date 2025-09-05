from openpyxl import Workbook
import _functions as cfx
import os

def create_list_files(main_folder, filenames, extension):
    for fls in filenames:
        file_path = os.path.join(main_folder, f"{fls}{extension}")
        if not os.path.exists(file_path):
            workbook  = Workbook()
            sheet = workbook.active
            sheet.title = "Sheet1"
            workbook.save(file_path)

main_folder = cfx.ifolder()
filenames = ["file1", "file2", "file3", "file4", "file5"]
create_list_files(main_folder, filenames, ".xlsx")
os.startfile(main_folder)