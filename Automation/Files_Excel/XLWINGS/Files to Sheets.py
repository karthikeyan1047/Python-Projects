import os, sys, shutil
import xlwings as xw
from datetime import datetime
import _functions as cfx

choice = int(cfx.inputbox(title='Method', prompt="Choose the method : \n   1. Combined Same Sheet's Data\n   2. Dont Combine Same Sheet's Data [FileName_SheetName] "))

app = xw.App(visible=False)

if choice == 1:
    folder_path = cfx.ifolder()
    new_wb = xw.Book()
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        try:
            current_wb = xw.Book(file_path)
        except:
            continue
        for sheet in current_wb.sheets:
            if sheet.name in [sht.name for sht in new_wb.sheets]:
                new_sheet = new_wb.sheets[sheet.name]
                if new_sheet.range('A1').value is None:
                    start_row = 1
                    data_start_row = 2
                else:
                    start_row = new_sheet.range('A1').end('down').row + 1
                    data_start_row = 2
            else:
                new_sheet = new_wb.sheets.add(sheet.name)
                start_row = 1
                data_start_row = 1

            for row in sheet.range('A1').expand('table').value[data_start_row-1:]:
                for col_num, value in enumerate(row, start=1):
                    cell = new_sheet.range((start_row, col_num))
                    cell.value = value
                    if isinstance(value, datetime):
                        cell.number_format = 'MM/DD/YYYY'
                start_row += 1

        current_wb.close()

    if "Sheet1" in [sht.name for sht in new_wb.sheets]:
        new_wb.sheets["Sheet1"].delete()

    ofolder_path = os.path.join(folder_path, 'Files_Sheets')

    if os.path.exists(ofolder_path):
        shutil.rmtree(ofolder_path)
    os.makedirs(ofolder_path)

    new_wb.save(os.path.join(ofolder_path, 'Files_Sheets.xlsx'))
    new_wb.close()
    os.startfile(folder_path)

elif choice == 2:
    folder_path = cfx.ifolder()
    new_wb = xw.Book()
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        try:
            current_wb = xw.Book(file_path)
        except:
            continue
        for sheet in current_wb.sheets:
            current_row = 1
            sh_nm = os.path.splitext(file_name)[0] + "_" + sheet.name
            new_sheet = new_wb.sheets.add(sh_nm)
            for row in sheet.range('A1').expand('table').value:
                for col_num, value in enumerate(row, start=1):
                    cell = new_sheet.range((current_row, col_num))
                    cell.value = value
                    if isinstance(value, datetime):
                        cell.number_format = 'MM/DD/YYYY'
                current_row += 1

        current_wb.close()

    if "Sheet1" in [sht.name for sht in new_wb.sheets]:
        new_wb.sheets["Sheet1"].delete()

    ofolder_path = os.path.join(folder_path, 'Files_Sheets')

    if os.path.exists(ofolder_path):
        shutil.rmtree(ofolder_path)
    os.makedirs(ofolder_path)

    new_wb.save(os.path.join(ofolder_path, 'Files_Sheets.xlsx'))
    new_wb.close()
    os.startfile(folder_path)

else:
    app.quit()
    sys.exit(0)

app.quit()