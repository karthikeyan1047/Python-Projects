import os
import _functions as cfx

def keep_latest_n_files(main_folder, keep_count):
    files_to_delete = [
        os.path.join(main_folder, file)
        for file in os.listdir(main_folder)
        if os.path.isfile(os.path.join(main_folder, file))
    ]
    files_to_delete.sort(key=os.path.getmtime, reverse=True)     # reverse = False --- to keep n oldest files
    files_to_delete = files_to_delete[keep_count:]
    for file in files_to_delete:
            os.remove(file)

main_folder = cfx.ifolder()
keep_count = int(cfx.inputbox("Folder", "Enter the number folders to keep"))
keep_latest_n_files(main_folder)
os.startfile(main_folder)