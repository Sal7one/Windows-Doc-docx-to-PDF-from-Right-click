import winreg
from pyuac import main_requires_admin
import sys
import os

label = "Convert to PDF"

@main_requires_admin
def create_registry_key(key, sub_key, value):
    try:
        with winreg.OpenKey(key, sub_key, 0, winreg.KEY_SET_VALUE) as key_handle:
            winreg.SetValueEx(key_handle, '', 0, winreg.REG_SZ, value)
        return True
    except FileNotFoundError:
        with winreg.CreateKey(key, sub_key) as key_handle:
            winreg.SetValueEx(key_handle, '', 0, winreg.REG_SZ, value)
        return True
    except Exception as e:
        print(e)
        return False

@main_requires_admin
def delete_registry_key(key, sub_key):
    try:
        with winreg.OpenKey(key, sub_key, 0, winreg.KEY_ALL_ACCESS) as key_handle:
            winreg.DeleteKey(key_handle, '')
        return True
    except FileNotFoundError:
        return False
    except Exception as e:
        print(e)
        return False

def add_context_menu():
    python_executable = sys.executable
    script_path = os.path.abspath('convertToPDF.py')

    command = f'"{python_executable}" "{script_path}" "%1"'

    file_types = ['.doc', '.docx']

    for file_type in file_types:
        key = winreg.HKEY_CLASSES_ROOT
        sub_key = rf'SystemFileAssociations\{file_type}\shell\{label}\command'

        create_registry_key(key, sub_key, command)

    return True

def remove_context_menu():
    key = winreg.HKEY_CLASSES_ROOT
    file_types = ['.doc', '.docx']
    success = True

    for file_type in file_types:
        sub_key = rf'SystemFileAssociations\{file_type}\shell\{label}\command'
        removeSubKey = delete_registry_key(key, sub_key)

        main_key = rf'SystemFileAssociations\{file_type}\shell\{label}'
        removeKey = delete_registry_key(key, main_key)
        success = removeSubKey and removeKey

    return success

if __name__ == "__main__":
    if len(sys.argv) > 1:
        command = sys.argv[1]
        if command == "add":
            result = add_context_menu()
            exit_code = 0 if result else 1
        elif command == "remove":
            result = remove_context_menu()
            exit_code = 0 if result else 1
        else:
            print("Invalid command")
            exit_code = 1
    else:
        print("No command provided")
        exit_code = 1

    sys.exit(exit_code)
