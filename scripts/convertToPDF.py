import sys
import os
import win32com.client

def convert_to_pdf(docx_path):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(docx_path)
    file_name, _ = os.path.splitext(docx_path)
    pdf_path = f"{file_name}.pdf"
    # 17 = PDF file format
    doc.SaveAs2(pdf_path, FileFormat=17)

   
    doc.Close()
    word.Quit()

if __name__ == "__main__":
    docx_path = sys.argv[1]
    convert_to_pdf(docx_path)
