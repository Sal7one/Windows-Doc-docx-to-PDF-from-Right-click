import winreg
import os
from pyuac import main_requires_admin

@main_requires_admin
def editWindowsRegistry():
    # Get the current working directory and the path to the Python script
    commandLabel = "Convert file to PDF"
    scriptTitle = "convertToPDF.py"
    currentDir = os.getcwd()
    scriptPath = os.path.join(currentDir, scriptTitle)
    docxScriptsPath = ".docx\Word.Document.12\ShellNew"

    root = winreg.ConnectRegistry(None, winreg.HKEY_CLASSES_ROOT)
    key = winreg.OpenKeyEx(root, docxScriptsPath)
    key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, docxScriptsPath,0, winreg.KEY_ALL_ACCESS)

    regShellCommand = f'python "{scriptPath}" %1 %2' # %1 directory path, %2 filepath
    winreg.SetValueEx(key, commandLabel, 0, winreg.REG_SZ, regShellCommand ) 

    # Close the keys
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
    with open("whoCaresSal7oneHopesThisHelps.txt", "w") as file:
        file.write("state")

if __name__ == "__main__":
    addToContextMenu()
