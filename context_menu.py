import winreg
from pyuac import main_requires_admin
import sys
import os
import subprocess

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

def get_python_path():
    if getattr(sys, 'frozen', False):
        command = 'where' if os.name == 'nt' else 'which'
        try:
            python_paths = subprocess.check_output([command, 'pythonw'], universal_newlines=True).strip().split('\n')
            # in windows you can get multiple paths to python interpreter choose the one that doesn't redirect to windows apps
            if python_paths:
                return [element for element in python_paths if "WindowsApps" not in element][0]
        except subprocess.CalledProcessError:
            raise RuntimeError('Python interpreter not found')
    else:
        python_path = sys.executable

    return python_path

def add_context_menu():
    ## because of pyinstaller otherwise use sys.executable
    python_executable = os.path.join(get_python_path()) 
    
    utils_path = os.path.abspath('scripts')
    script_path = os.path.join(utils_path, 'convertToPDF.py')

    command = f'{python_executable} "{script_path}" "%1"' # %1 gives file path, registry rules

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
