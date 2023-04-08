import winreg
import os
from pyuac import main_requires_admin

@main_requires_admin
def removeFromContextMenu():
    # Get the HKEY_CLASSES_ROOT key
    root = winreg.ConnectRegistry(None, winreg.HKEY_CLASSES_ROOT)

    # Define the name of the context menu item to remove
    commandLabel = "Convert file to PDF"
    docxScriptsPath = ".docx\Word.Document.12\ShellNew"

    # Check if the context menu item exists and delete it if it does
    if fileExists():
        try:
            key = winreg.OpenKey(root, docxScriptsPath)
            winreg.DeleteValue(key, commandLabel)
            print("Context menu item removed successfully!")
            winreg.CloseKey(key)

        except OSError as e:
            print(f"Error: {e}")
        deleteFile()
    else:
        print("Nothing to remove")
    winreg.CloseKey(root)
    winreg.CloseKey(root)

fileName = "whoCaresSal7oneHopesThisHelps.txt"

def deleteFile():
    os.remove(fileName)

def fileExists():
    # Check if the file exists
    return os.path.exists(fileName)

if __name__ == "__main__":
    removeFromContextMenu()
