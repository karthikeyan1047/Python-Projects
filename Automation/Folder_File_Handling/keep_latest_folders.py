import os
import shutil
import _functions as cfx
import stat

def keep_latest_n_folder(main_folder, keep_count):
    def wirte_access(func, path, exc_info):
        os.chmod(path, stat.S_IWRITE)
        func(path)
    flds_to_delete = [
        os.path.join(main_folder, fld) 
        for fld in os.listdir(main_folder) 
        if os.path.isdir(os.path.join(main_folder, fld))
    ]
    flds_to_delete.sort(key=os.path.getmtime, reverse=True)     # reverse = False --- to keep n oldest folders
    flds_to_delete = flds_to_delete[keep_count:]
    for f in flds_to_delete:
        shutil.rmtree(f, onerror=wirte_access)

main_folder = cfx.ifolder()
keep_count = int(cfx.inputbox("Folder", "Enter the number folders to keep"))
keep_latest_n_folder(main_folder, keep_count)
os.startfile(main_folder)