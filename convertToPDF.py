import win32com.client
import sys
import os

def convert_to_pdf():
    # Check if the script was called with the correct number of arguments
    if len(sys.argv) < 3:
        print("Please provide a directory path and a complete file path with extension.")
        sys.exit(1)

    # Create a Word Application object and open the file
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    directory_path = sys.argv[1]
    complete_file_path_with_extension = sys.argv[2]
    
    doc = word.Documents.Open(f"{complete_file_path_with_extension}")

    # Save the file as a PDF in the specified directory
    file_name, extension = os.path.splitext(complete_file_path_with_extension)

    # 17 pdf ref https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
    fileName = os.path.join(directory_path, file_name + ".pdf")
    doc.SaveAs2(fileName, FileFormat=17)

    # Close the file and quit Word
    doc.Close()
    word.Quit()

if __name__ == "__main__":
    convert_to_pdf()
