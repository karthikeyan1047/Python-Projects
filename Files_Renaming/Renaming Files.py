import os
import _functions as cfx

folder_path = cfx.ifolder()
os.startfile(folder_path)
choice = int(cfx.inputbox("Renaming", "1. Before Delimiter\n2. After Delimiter\n3. Add Suffix\n4. Add Prefix\n5. Remove first n characters\n6. Remove last n characters\n"))

if choice == 1 :
    delimiter = str(cfx.inputbox(title='Before Delimiter', prompt="Enter the delimiter"))
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            _, file_extension = os.path.splitext(filename)
            pos_del = filename.find(delimiter)
            newfilename = filename[:pos_del] + file_extension
            old_file_path = os.path.join(folder_path, filename)
            new_file_path = os.path.join(folder_path, newfilename)
            os.rename(old_file_path, new_file_path)
            os.startfile(folder_path)

elif choice == 2:
    delimiter = str(cfx.inputbox(title='After Delimiter', prompt="Enter the delimiter"))
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            pos_del = filename.find(delimiter)
            newfilename = filename[pos_del+1:]
            old_file_path = os.path.join(folder_path, filename)
            new_file_path = os.path.join(folder_path, newfilename)
            os.rename(old_file_path, new_file_path)
            os.startfile(folder_path)

elif choice == 3:
    suffix = str(cfx.inputbox("Prefix", "Type the suffix to add in the File Name"))
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            fname, fext= os.path.splitext(filename)
            nfname = fname + "_" + suffix + fext
            old_file_path = os.path.join(folder_path, filename)
            new_file_path = os.path.join(folder_path, nfname)
            os.rename(old_file_path, new_file_path)
            os.startfile(folder_path)
            
elif choice == 4:
    prefix = str(cfx.inputbox("Prefix", "Type the preffix to add in the File Name"))
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            fname, fext= os.path.splitext(filename)
            nfname = prefix + "_" + fname + fext
            old_file_path = os.path.join(folder_path, filename)
            new_file_path = os.path.join(folder_path, nfname)
            os.rename(old_file_path, new_file_path)
            os.startfile(folder_path)

elif choice == 5:
    n = int(cfx.inputbox("First", "Number of characters to remove"))
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            fname, fext= os.path.splitext(filename)
            fname = fname[n:] + fext
            old_file_path = os.path.join(folder_path, filename)
            new_file_path = os.path.join(folder_path, fname)
            os.rename(old_file_path, new_file_path)
            os.startfile(folder_path)

elif choice == 6:
    n = int(cfx.inputbox("Last", "Number of characters to remove"))
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            fname, fext= os.path.splitext(filename)
            fname = fname[:-n]+fext
            old_file_path = os.path.join(folder_path, filename)
            new_file_path = os.path.join(folder_path, fname)
            os.rename(old_file_path, new_file_path)
            os.startfile(folder_path)

