import win32com.client
import sys
import os
import ctypes

def convert_to_pdf(docx_path):
    # Check if the file path is valid and has the correct extension
    if not os.path.exists(docx_path):
        directRunMessage()
        sys.exit(1)

    # Create a Word Application object and open the file
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(docx_path)

    # Save the file as a PDF in the same directory
    file_name, extension = os.path.splitext(docx_path)
    pdf_path = os.path.join(file_name + ".pdf")
    doc.SaveAs2(pdf_path, FileFormat=17)  # 17 is the constant for PDF format in Word

    # Close the file and quit Word
    doc.Close()
    word.Quit()

def directRunMessage():
    ctypes.windll.user32.MessageBoxW(0, f"This script is supposed to run from a docx | doc file context menu (No action)", "Success", 0x40 | 0x1)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        docx_path = sys.argv[1]
        convert_to_pdf(docx_path)
    else:
        directRunMessage()
