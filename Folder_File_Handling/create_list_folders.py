import os
import _functions as cfx

def create_date_files(main_folder, foldersnames):
    for flds in foldersnames:
        folder_path = main_folder + "/" + flds
        os.makedirs(folder_path)


main_folder = cfx.ifolder()
foldersnames = ["water", "plant", "fuel", "Petrol"]
create_date_files(main_folder, foldersnames)