import winreg
from pyuac import main_requires_admin
import ctypes

@main_requires_admin
def delete_registry_key(key, sub_key):
    try:
        with winreg.OpenKey(key, sub_key, 0, winreg.KEY_ALL_ACCESS) as key_handle:
            winreg.DeleteKey(key_handle, '')
        return True
    except FileNotFoundError:
        return False

@main_requires_admin
def main():
    key = winreg.HKEY_CLASSES_ROOT
    file_types = ['.doc', '.docx']
    success = True

    for file_type in file_types:

        sub_key = rf'SystemFileAssociations\{file_type}\shell\Convert to PDF\command'
        removeSubKey = delete_registry_key(key, sub_key)

        main_key = rf'SystemFileAssociations\{file_type}\shell\Convert to PDF'
        removeKey = delete_registry_key(key, main_key)
        success = removeSubKey and removeKey

    if success:
        ctypes.windll.user32.MessageBoxW(0, "Successfully removed 'Convert to PDF' from the context menu for .doc and .docx files.", "Success", 0x40 | 0x1)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Could not find 'Convert to PDF' in the context menu || Already removed or you didn't install in the first place", "Error",  0x10)

if __name__ == '__main__':
    main()
