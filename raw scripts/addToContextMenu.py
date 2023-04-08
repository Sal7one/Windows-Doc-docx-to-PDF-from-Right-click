import os
import sys
import ctypes
import winreg
from pyuac import main_requires_admin

@main_requires_admin
def create_registry_key(key, sub_key, value):
    try:
        with winreg.OpenKey(key, sub_key, 0, winreg.KEY_SET_VALUE) as key_handle:
            winreg.SetValueEx(key_handle, '', 0, winreg.REG_SZ, value)
    except FileNotFoundError:
        with winreg.CreateKey(key, sub_key) as key_handle:
            winreg.SetValueEx(key_handle, '', 0, winreg.REG_SZ, value)

@main_requires_admin
def main():
    python_executable = sys.executable
    script_path = os.path.abspath('convertToPDF.py')

    command = f'"{python_executable}" "{script_path}" "%1" "%2"'

    file_types = ['.doc', '.docx']

    for file_type in file_types:
        key = winreg.HKEY_CLASSES_ROOT
        sub_key = rf'SystemFileAssociations\{file_type}\shell\Convert to PDF\command'

        create_registry_key(key, sub_key, command)

    ctypes.windll.user32.MessageBoxW(0, "Successfully added 'Convert to PDF' to the context menu for .doc and .docx files.", "Success", 0x40 | 0x1)

if __name__ == '__main__':
    main()
