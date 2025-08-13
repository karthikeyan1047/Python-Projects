import os
import shutil

def handle_remove_readonly(func, path, exc_info):
    import stat
    os.chmod(path, stat.S_IWRITE)
    func(path)

def copy_file_to_subfolders(main_folder, source_file):
    if not os.path.exists(source_file):
        return
    if not os.path.exists(main_folder):
        return
    for root, dirs, files in os.walk(main_folder):
        for subfolder in dirs:
            subfolder_path = os.path.join(root, subfolder)
            target_file_path = os.path.join(subfolder_path, os.path.basename(source_file))
            try:
                shutil.copy(source_file, target_file_path)
            except Exception as e:
                print(f"Failed to copy to '{subfolder_path}': {e}")

def copy_files_to_last_subfolder(main_folder, source_file, exclude_folder):

    if not os.path.exists(source_file):
        return
    if not os.path.exists(main_folder):
        return

    exclude_folder = os.path.abspath(exclude_folder)
    main_folder = os.path.abspath(main_folder)

    for root, dirs, files in os.walk(main_folder):
        if exclude_folder in os.path.abspath(root):
            continue
        if not dirs:
            target_file_path = os.path.join(root, os.path.basename(source_file))
            try:
                shutil.copy(source_file, target_file_path)
            except Exception as e:
                print(f"Failed to copy to '{root}': {e}")

def delete_file_from_folders(main_folder, file_name):
    if not os.path.exists(main_folder):
        return
    for root, dirs, files in os.walk(main_folder):
        if root == main_folder:       
            continue                # remove these 2 line if main folder's files also needs to be deleted
        if file_name in files:
            file_path = os.path.join(root, file_name)
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"Failed to delete '{file_path}': {e}")

def delete_a_folders(main_folder):
    if not os.path.exists(main_folder):
        return
    for root, dirs, files in os.walk(main_folder):
        for folder in dirs:
            if folder == "__pycache__":
                folder_path = os.path.join(root, folder)
                try:
                    shutil.rmtree(folder_path, onerror=handle_remove_readonly)
                except Exception as e:
                    print(f"Failed to delete folder '{folder_path}': {e}")

if __name__ == "__main__":
    main_folder = r"C:\Users\karthikeyans\Documents\Automations\Python"
    exclude_folder = r"C:\Users\karthikeyans\Documents\Automations\Python\.venv"
    source_file = r"C:\Users\karthikeyans\Documents\Automations\Python\_functions.py"
    file_name = '_functions.py'
    delete_file_from_folders(main_folder, file_name)
    delete_a_folders(main_folder)
    copy_files_to_last_subfolder(main_folder, source_file, exclude_folder)
    # copy_file_to_subfolders(main_folder, source_file)
    # delete_file_from_folders(main_folder, file_name)

