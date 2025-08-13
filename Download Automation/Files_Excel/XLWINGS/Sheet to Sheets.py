import os
import shutil
import sys
import xlwings as xw
import _functions as cfx

def save_and_close(wb, path):
    if "Sheet1" in [sht.name for sht in wb.sheets]:
        wb.sheets["Sheet1"].delete()
    wb.save(path)
    wb.close()

def process_single_file_single_sheet():
    file_path = cfx.ifile()
    filter_column = int(cfx.inputbox(title='Column Number', prompt="Enter the column number containing category to split"))
    
    folder_path = os.path.dirname(file_path)
    odirectory = os.path.join(folder_path, "Sheet to Sheets")
    if os.path.exists(odirectory):
        shutil.rmtree(odirectory)
    os.makedirs(odirectory)
    
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets[0]
    data = sheet.used_range.value
    headers, rows = data[0], data[1:]
    
    categories = {}
    for row in rows:
        category = row[filter_column - 1]
        if category not in categories:
            categories[category] = []
        categories[category].append(row)
    
    dest_wb = xw.Book()
    for category, rows in categories.items():
        category = " ".join(str(category).split(" ")[:2])
        new_sheet = dest_wb.sheets.add(name=category)
        new_sheet.range("A1").value = [headers] + rows
    
    dest_file_path = os.path.join(odirectory, "Categorized_Data.xlsx")
    save_and_close(dest_wb, dest_file_path)
    wb.close()
    app.quit()
    os.startfile(os.path.dirname(odirectory))

def process_single_file_multiple_sheets():
    file_path = cfx.ifile()
    filter_column = int(cfx.inputbox(title='Column Number', prompt="Enter the column number containing category to split"))
    
    folder_path = os.path.dirname(file_path)
    odirectory = os.path.join(folder_path, "Sheet to Sheets")
    if os.path.exists(odirectory):
        shutil.rmtree(odirectory)
    os.makedirs(odirectory)
    
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    categories = {}
    headers = None
    
    for sheet in wb.sheets:
        data = sheet.used_range.value
        if headers is None:
            headers = data[0]
        rows = data[1:]
        for row in rows:
            category = row[filter_column - 1]
            if category not in categories:
                categories[category] = []
            categories[category].append(row)
    
    dest_wb = xw.Book()
    for category, rows in categories.items():
        category = " ".join(str(category).split(" ")[:2])
        new_sheet = dest_wb.sheets.add(name=category)
        new_sheet.range("A1").value = [headers] + rows
    
    dest_file_path = os.path.join(odirectory, "Categorized_Data.xlsx")
    save_and_close(dest_wb, dest_file_path)
    wb.close()
    app.quit()
    os.startfile(os.path.dirname(odirectory))

def process_multiple_files_multiple_sheets():
    folder_path = cfx.ifolder()
    filter_column = int(cfx.inputbox(title='Column Number', prompt="Enter the column number containing category to split"))
    
    odirectory = os.path.join(folder_path, "Sheet to Sheets")
    if os.path.exists(odirectory):
        shutil.rmtree(odirectory)
    os.makedirs(odirectory)
    
    app = xw.App(visible=False)
    categories = {}
    headers = None
    
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    for file in excel_files:
        wb = app.books.open(os.path.join(folder_path, file))
        for sheet in wb.sheets:
            data = sheet.used_range.value
            if headers is None:
                headers = data[0]
            rows = data[1:]
            for row in rows:
                category = row[filter_column - 1]
                if category not in categories:
                    categories[category] = []
                categories[category].append(row)
        wb.close()
    
    dest_wb = xw.Book()
    for category, rows in categories.items():
        category = " ".join(str(category).split(" ")[:2])
        new_sheet = dest_wb.sheets.add(name=category)
        new_sheet.range("A1").value = [headers] + rows
    
    dest_file_path = os.path.join(odirectory, "Categorized_Data.xlsx")
    save_and_close(dest_wb, dest_file_path)
    app.quit()
    os.startfile(os.path.dirname(odirectory))

choice = int(cfx.inputbox(title='Method', prompt='Choose the suitable method\n   1. Single File_Single Sheet\n   2. Single File_Multiple Sheets\n   3. Multiple Files_Multiple Sheets'))
if choice == 1:
    process_single_file_single_sheet()
elif choice == 2:
    process_single_file_multiple_sheets()
elif choice == 3:
    process_multiple_files_multiple_sheets()
else:
    sys.exit(0)
