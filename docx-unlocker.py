import zipfile
import re
import os
import sys

def unlock_docx(file_name):
    already_unlocked_pattern = r'^.*\_UNLOCKED\.docx$' #determine if the file has already been unlocked
    if re.match(already_unlocked_pattern, file_name):
        print(f"{file_name} is already unlocked, skipping.")
        return
    
    # Create a new file name with the "_UNLOCKED" suffix
    new_file_name = file_name[:-5] + "_UNLOCKED.docx"

    # Check if the new file name already exists
    if os.path.exists(new_file_name):
        overwrite = input(f"{new_file_name} already exists. Do you want to overwrite it? (y/n): ")
        if overwrite.lower() != 'y':
            print("Operation canceled.")
            return

    # Open the original .docx file (which is a zip archive)
    with zipfile.ZipFile(file_name, 'r') as docx_zip:
        # Extract the 'settings.xml' file
        with docx_zip.open('word/settings.xml') as settings_file:
            xml_content = settings_file.read().decode('utf-8')

            # Use regex to find and replace 'w:enforcement="1"' with 'w:enforcement="0"'
            new_xml_content, replacements_made = re.subn(r'(<w:documentProtection[^>]*w:enforcement=")1(")', r'\g<1>0\g<2>', xml_content)

            # Check if any replacements were made
            if replacements_made > 0:
                # Save the modified 'settings.xml' file
                with zipfile.ZipFile(new_file_name, 'w') as new_docx_zip:
                    for item in docx_zip.infolist():
                        # Replace the 'settings.xml' file with the modified version
                        if item.filename == 'word/settings.xml':
                            new_docx_zip.writestr(item, new_xml_content)
                        else:
                            new_docx_zip.writestr(item, docx_zip.read(item))
                print(f"{file_name} has been unlocked, and the new file is saved as {new_file_name}")
            else:
                print(f"{file_name} is already unlocked or does not contain the specified documentProtection element.")

def main(input_path):
    if os.path.isdir(input_path):
        confirmation = input(f"{input_path} is a directory. Do you want to unlock all .docx files in this directory? (y/n): ")
        if confirmation.lower() == 'y':
            for filename in os.listdir(input_path):
                if filename.endswith('.docx'):
                    unlock_docx(os.path.join(input_path, filename))
        else:
            print("Operation canceled.")
    elif os.path.isfile(input_path) and input_path.endswith('.docx'):
        unlock_docx(input_path)
    else:
        print(f"{input_path} is not a valid .docx file or directory.")

def print_usage():
    print("Usage:")
    print("  python script_name.py [path_to_docx_file_or_directory]")
    print("\nDescription:")
    print("  This script unlocks .docx files by modifying the 'w:enforcement' attribute.")
    print("  If a directory is provided, the script will unlock all .docx files within that directory.")
    print("\nNOTE: if your filename or path has a space in it, encapsulate the filename/path with double quotes.")
    print("\nExamples:")
    print("  python script_name.py protected_file.docx")
    print("  python script_name.py /path/to/directory")
    print("  python script_name.py \"C:\\Users\\username\\Downloads\\my docx file.docx\"")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print_usage()
    else:
        input_path = sys.argv[1]
        main(input_path)

