import xlwings as xw
import os
import shutil
import _functions as cfx
from datetime import datetime

def save_as_xlsx(sheet, new_filename):
    with xw.App(visible=False) as app:
        new_wb = app.books.add()
        new_sheet = new_wb.sheets[0]
        new_sheet.name = sheet.name
        
        data = sheet.used_range.value
        new_sheet.range("A1").value = data
        
        new_wb.save(new_filename)
        new_wb.close()

file_path = cfx.ifile()
folder_path = os.path.join(os.path.dirname(file_path), "Sheets_Files")
if os.path.exists(folder_path):
    shutil.rmtree(folder_path)
os.makedirs(folder_path)

with xw.App(visible=False) as app:
    wb = app.books.open(file_path)
    
    for sheet in wb.sheets:
        new_filename = os.path.join(folder_path, f"{sheet.name}.xlsx")
        save_as_xlsx(sheet, new_filename)
    
    wb.close()

os.startfile(folder_path)
