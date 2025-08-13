import win32com.client as win32
import _functions as cfx
import os
import time
import sys

def convert_file(inputfile_path, dest_ext):
    src_ext = os.path.splitext(inputfile_path)[1].lower()
    output_folder = os.path.join(os.path.dirname(inputfile_path), "Converted Files")

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(inputfile_path)

        if dest_ext.lower() == '.csv' and src_ext in ['.xlsx', '.xlsb', '.xls']:
            for sheet in wb.Sheets:
                sheet_name = sheet.Name.replace("/", "_").replace("\\", "_")
                csv_file = os.path.join(output_folder, f"{sheet_name}.csv")
                sheet.Copy()
                temp_wb = excel.ActiveWorkbook
                temp_wb.SaveAs(csv_file, FileFormat=6)
                temp_wb.Close(False)
        else:
            output_file = os.path.normpath(
                os.path.join(
                    output_folder,
                    os.path.basename(inputfile_path).replace(src_ext, dest_ext)
                )
            )

            file_format_map = {
                '.xlsx': 51,
                '.xls': 56,
                '.xlsb': 50,
                '.csv': 6
            }
            fl_format = file_format_map.get(dest_ext.lower(), 51)
            wb.SaveAs(output_file, FileFormat=fl_format, Local=True)

        wb.Close(False)
        os.startfile(output_folder)
        return output_folder
    except Exception as e:
        print(f"Error: {e}")
        return None
    finally:
        excel.DisplayAlerts = True
        excel.Quit()

def convert_folder(folder_path, src_exts, dest_ext):
    for file in os.listdir(folder_path):
        src_ext = os.path.splitext(file)[1]
        if src_ext in src_exts:
            inputfile_path = os.path.join(folder_path, file)
            output_file = convert_file(inputfile_path, dest_ext)
            
        else:
            sys.exit(0)
    os.startfile(folder_path)
    return output_file

choice = int(cfx.inputbox(title='Options', prompt='1. Folder\n2. File'))
file_format = int(cfx.inputbox(title='Conversion', prompt='1. To CSV\n2. To XLSX\n3. To XLSB\n4. To XLS'))
extensions = {
    1: (".csv", 6),
    2: (".xlsx", 51),
    3: (".xlsb", 50),
    4: (".xls", 56)
}
dest_ext, fl_format = extensions.get(file_format, (None, None))

if choice == 1:
    folder_path = cfx.ifolder(title='Select the folder to convert files')
    src_exts = [".xls", ".csv", ".xlsx", ".xlsb"]
    convert_folder(folder_path, src_exts, dest_ext)

elif choice == 2:
    inputfile_path = cfx.ifile(title='Select the file to convert')
    convert_file(inputfile_path, dest_ext)
else:
    sys.exit(0)