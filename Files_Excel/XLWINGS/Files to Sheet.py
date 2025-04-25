import os
import xlwings as xw
from datetime import datetime
import _functions as cfx

folder_path = cfx.ifolder()
workbook_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

app = xw.App(visible=False)
combined_workbook = xw.Book()
combined_sheet = combined_workbook.sheets[0]
combined_sheet.name = 'Combined'
current_row = 1
is_first_sheet = True

for file in workbook_files:
    file_path = os.path.join(folder_path, file)
    try:
        wb = xw.Book(file_path)
    except:
        continue
    for sheet in wb.sheets:
        for row_num, row in enumerate(sheet.range('A1').expand('table').value, start=1):
            if row_num == 1 and not is_first_sheet:
                continue
            for col_num, value in enumerate(row, start=1):
                if isinstance(value, datetime):
                    combined_sheet.range((current_row, col_num)).value = value
                    combined_sheet.range((current_row, col_num)).number_format = 'MM/DD/YYYY'
                else:
                    combined_sheet.range((current_row, col_num)).value = value
            current_row += 1
        is_first_sheet = False
    wb.close()
combined_workbook.save(os.path.join(folder_path, 'Files_Sheet.xlsx'))
combined_workbook.close()
app.quit()