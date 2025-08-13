import os
import _functions as cfx

def rename_files(folder, rename_func):
    for filename in os.listdir(folder):
        full_path = os.path.join(folder, filename)
        if os.path.isfile(full_path):
            new_name = rename_func(filename)
            if new_name and new_name != filename:
                new_path = os.path.join(folder, new_name)
                os.rename(full_path, new_path)

def main():
    folder_path = cfx.ifolder()
    os.startfile(folder_path)

    choice = int(cfx.inputbox("Renaming", 
        "1. Before Delimiter\n2. After Delimiter\n3. Add Suffix\n4. Add Prefix\n5. Remove first n characters\n6. Remove last n characters\n"))

    if choice == 1:
        delimiter = cfx.inputbox("Before Delimiter", "Enter the delimiter")
        n = int(cfx.inputbox("Delimiter Occurrence", "Use before which occurrence of the delimiter? (1 = first, 2 = second, etc.)"))

        def before_nth_delimiter(filename):
            fname, ext = os.path.splitext(filename)
            parts = fname.split(delimiter)
            if len(parts) >= n:
                return delimiter.join(parts[:n]) + ext
            return filename  # return unchanged if not enough delimiters

        rename_files(folder_path, before_nth_delimiter)

    elif choice == 2:
        delimiter = cfx.inputbox("After Delimiter", "Enter the delimiter")
        n = int(cfx.inputbox("Delimiter Occurrence", "Use after which occurrence of the delimiter? (1 = first, 2 = second, etc.)"))

        def after_nth_delimiter(filename):
            fname, ext = os.path.splitext(filename)
            parts = fname.split(delimiter)
            if len(parts) > n:
                return delimiter.join(parts[n:]) + ext
            return filename  # return unchanged if not enough delimiters

        rename_files(folder_path, after_nth_delimiter)

    elif choice == 3:
        suffix = cfx.inputbox("Suffix", "Type the suffix to add in the File Name")
        rename_files(folder_path, lambda f: f"{os.path.splitext(f)[0]}_{suffix}{os.path.splitext(f)[1]}")

    elif choice == 4:
        prefix = cfx.inputbox("Prefix", "Type the prefix to add in the File Name")
        rename_files(folder_path, lambda f: f"{prefix}_{os.path.splitext(f)[0]}{os.path.splitext(f)[1]}")

    elif choice == 5:
        n = int(cfx.inputbox("Remove First", "Number of characters to remove"))
        rename_files(folder_path, lambda f: f"{os.path.splitext(f)[0][n:]}{os.path.splitext(f)[1]}")

    elif choice == 6:
        n = int(cfx.inputbox("Remove Last", "Number of characters to remove"))
        rename_files(folder_path, lambda f: f"{os.path.splitext(f)[0][:-n]}{os.path.splitext(f)[1]}" if n else f)

    os.startfile(folder_path)

if __name__ == "__main__":
    main()

