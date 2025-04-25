from openpyxl import Workbook
import _functions as cfx
import os

def create_date_files(main_folder, filenames, extension):
    for fls in filenames:
        file_path = main_folder + "/" + fls + extension
        workbook  = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        workbook.save(file_path)

main_folder = cfx.ifolder()
filenames = ["water", "plant", "fuel", "Petrol"]
create_date_files(main_folder, filenames, ".xlsx")
os.startfile(main_folder)