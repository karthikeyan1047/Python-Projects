import xlwings as xw
import os, shutil
from datetime import datetime
import _functions as cfx

app = xw.App(visible=False)
file_path = cfx.ifile()
wb = xw.Book(file_path)
combined_book = xw.Book()
combined_sheet = combined_book.sheets[0]
combined_sheet.name = "Combined"

is_first_sheet = True
current_row = 1

def is_date(value):
    return isinstance(value, datetime)

for sheet in wb.sheets:
    data = sheet.used_range.value
    if not data:
        continue
    
    if not is_first_sheet:
        data = data[1:]
    
    for row in data:
        for col_num, value in enumerate(row, start=1):
            cell = combined_sheet.cells(current_row, col_num)
            cell.value = value
            if is_date(value):
                cell.number_format = "MM/DD/YYYY"
        current_row += 1
    
    is_first_sheet = False

wb.close()

dest_directory = os.path.dirname(file_path)
dest_folder = os.path.join(dest_directory, "Sheets_Sheet")

if os.path.exists(dest_folder):
    shutil.rmtree(dest_folder)
os.makedirs(dest_folder)

dest_file = os.path.join(dest_folder, "Sheets_Sheet.xlsx")
combined_book.save(dest_file)
combined_book.close()
app.quit()
os.startfile(dest_folder)
