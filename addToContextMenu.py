import winreg
import os
import random

def editWindowsRegistry():
    # Get the current working directory and the path to the Python script
    current_dir = os.getcwd()
    script_path = os.path.join(current_dir, "convertToPDF.py")

    # Get the HKEY_CLASSES_ROOT key and create a new key for the context menu
    root = winreg.ConnectRegistry(None, winreg.HKEY_CLASSES_ROOT)
    key = winreg.CreateKey(root, r"Directory\shell\Convert file to PDF")

    # Set the command to run the Python script
    command_key = winreg.CreateKey(key, "ConvertfiletoPDF")
    winreg.SetValueEx(command_key, "", 0, winreg.REG_SZ, f'python "{script_path}" %2 %1')

    # Close the keys
    winreg.CloseKey(command_key)
    winreg.CloseKey(key)
    winreg.CloseKey(root)

def addToContextMenu():
    if os.path.exists("whoCaresSal7oneHopesThisHelps.txt"):
        print("Context menu item already added. To remove it, use the remove script.")
    else:
        editWindowsRegistry()
        setRandomNumberFile()
        print("Context menu item added successfully.")

def setRandomNumberFile():
    # Generate a random 7-digit number and save it to a file
    random_number = str(random.randint(1000000, 9999999))
    with open("whoCaresSal7oneHopesThisHelps.txt", "w") as file:
        file.write(random_number)

if __name__ == "__main__":
    addToContextMenu()
