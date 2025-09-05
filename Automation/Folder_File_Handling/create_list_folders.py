import os
import _functions as cfx

def create_list_folders(main_folder, foldersnames):
    for flds in foldersnames:
        folder_path = main_folder + "/" + flds
        os.makedirs(folder_path)


main_folder = cfx.ifolder()
foldersnames = ["folder1", "folder2", "folder3", "folder4"]
create_list_folders(main_folder, foldersnames)