import winreg
import os
from pyuac import main_requires_admin

@main_requires_admin
def removeFromContextMenu():
    # Get the HKEY_CLASSES_ROOT key
    root = winreg.ConnectRegistry(None, winreg.HKEY_CLASSES_ROOT)

    # Define the name of the context menu item to remove
    globalCommandName = "ConvertfiletoPDF"

    # Check if the context menu item exists and delete it if it does
    if fileExists():
        try:
            key = winreg.OpenKey(root, r"Directory\shell\Convert file to PDF")
            winreg.DeleteKey(key, globalCommandName)
            winreg.DeleteKey(root, r"Directory\shell\Convert file to PDF")
            print("Context menu item removed successfully!")
        except OSError as e:
            print(f"Error: {e}")
    else:
        print("Nothing to remove")
    
    # Close the key
    winreg.CloseKey(root)

def fileExists():
    # Check if the file exists
    file_path = "whoCaresSal7oneHopesThisHelps.txt"
    return os.path.exists(file_path)

if __name__ == "__main__":
    removeFromContextMenu()
