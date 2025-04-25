import os
import shutil
import sys
import xlwings as xw
import _functions as cfx

app = xw.App(visible=False)

def create_new_workbook():
    wb = xw.Book()
    while len(wb.sheets) > 1:
        wb.sheets[-1].delete()
    return wb

choice = int(cfx.inputbox(title='Method', prompt='Choose the suitable method\n   1. Single File_Single Sheet\n   2. Single File_Multiple Sheets\n   3. Multiple Files_Multiple Sheets'))

if choice == 1:
    file_path = cfx.ifile()
    filter_column = int(cfx.inputbox(title='Column Number', prompt="Enter the column number contains category to split"))

    folder_path = os.path.dirname(file_path)
    odirectory = os.path.join(folder_path, "Sheet to Files")
    if os.path.exists(odirectory):
        shutil.rmtree(odirectory)
    os.makedirs(odirectory)

    wb = xw.Book(file_path)
    sheet = wb.sheets[0]
    data = sheet.used_range.value
    headers = data[0]
    categories = {}

    for row in data[1:]:
        category = row[filter_column - 1]
        if category not in categories:
            categories[category] = []
        categories[category].append(row)

    for category, rows in categories.items():
        new_wb = create_new_workbook()
        new_sheet = new_wb.sheets[0]
        new_sheet.name = str(category)
        new_sheet.range("A1").value = [headers] + rows

        dest_file_path = os.path.join(odirectory, f"{str(category)}.xlsx")
        new_wb.save(dest_file_path)
        new_wb.close()

    os.startfile(odirectory)
    wb.close()
    app.quit()

elif choice == 2:
    file_path = cfx.ifile()
    filter_column = int(cfx.inputbox(title='Column Number', prompt="Enter the column number contains category to split"))

    folder_path = os.path.dirname(file_path)
    odirectory = os.path.join(folder_path, "Sheet to Files")
    if os.path.exists(odirectory):
        shutil.rmtree(odirectory)
    os.makedirs(odirectory)

    wb = xw.Book(file_path)
    categories = {}

    for sheet in wb.sheets:
        data = sheet.used_range.value
        headers = data[0]
        for row in data[1:]:
            category = row[filter_column - 1]
            if category not in categories:
                categories[category] = []
            categories[category].append(row)

    for category, rows in categories.items():
        new_wb = create_new_workbook()
        new_sheet = new_wb.sheets[0]
        new_sheet.name = str(category)
        new_sheet.range("A1").value = [headers] + rows

        dest_file_path = os.path.join(odirectory, f"{str(category)}.xlsx")
        new_wb.save(dest_file_path)
        new_wb.close()

    os.startfile(odirectory)
    wb.close()
    app.quit()

elif choice == 3:
    folder_path = cfx.ifolder()
    filter_column = int(cfx.inputbox(title='Column Number', prompt="Enter the column number contains category to split"))

    odirectory = os.path.join(folder_path, "Sheet to Files")
    if os.path.exists(odirectory):
        shutil.rmtree(odirectory)
    os.makedirs(odirectory)

    categories = {}
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        wb = xw.Book(file_path)
        for sheet in wb.sheets:
            data = sheet.used_range.value
            headers = data[0]
            for row in data[1:]:
                category = row[filter_column - 1]
                if category not in categories:
                    categories[category] = []
                categories[category].append(row)
        wb.close()

    for category, rows in categories.items():
        new_wb = create_new_workbook()
        new_sheet = new_wb.sheets[0]
        new_sheet.name = str(category)
        new_sheet.range("A1").value = [headers] + rows

        dest_file_path = os.path.join(odirectory, f"{str(category)}.xlsx")
        new_wb.save(dest_file_path)
        new_wb.close()

    os.startfile(odirectory)
    app.quit()
    
else:
    sys.exit(0)
