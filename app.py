import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QProgressBar, QFileDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from context_menu_handler import add_context_menu, remove_context_menu
from convertToPDF import convert_to_pdf

class ConvertThread(QThread):
    progress = pyqtSignal(int)

    def __init__(self, docx_path):
        super().__init__()
        self.docx_path = docx_path

    def run(self):
        try:
            convert_to_pdf(self.docx_path, self.progress)
            self.progress.emit(100)
        except Exception as e:
            print(e)
            self.progress.emit(-1)

class GlossyGUI(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Glossy GUI')

        # Set a custom style sheet for a glossy feel
        self.setStyleSheet("""
            QPushButton {
                background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #8a2be2, stop: 1 #551a8b);
                border: 1px solid #4A4A4A;
                border-radius: 5px;
                color: white;
                padding: 15px;
                min-width: 150px;
                font-size: 16px;
            }
 
            QPushButton:hover {
                background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #6C6C6C, stop: 1 #353535);
            }
            QPushButton:pressed {
                background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #5D5D5D, stop: 1 #1E1E1E);
            }
            QProgressBar {
                border: 1px solid #4A4A4A;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #5D5D5D, stop: 1 #252525);
                border-radius: 4px;
            }
            QLabel {
                font-size: 16px;
            }
        """)

        layout = QVBoxLayout()

        add_button = QPushButton('Add to Context Menu')
        add_button.clicked.connect(self.add_to_context_menu)
        layout.addWidget(add_button)

        remove_button = QPushButton('Remove from Context Menu')
        remove_button.clicked.connect(self.remove_from_context_menu)
        layout.addWidget(remove_button)

        self.status_label = QLabel()
        layout.addWidget(self.status_label)

        self.setLayout(layout)


    def add_to_context_menu(self):
        if add_context_menu():
            self.status_label.setText("Successfully added 'Convert to PDF' to the context menu for .doc and .docx files.")
        else:
            self.status_label.setText("An error occurred while adding 'Convert to PDF' to the context menu.")

    def remove_from_context_menu(self):
        if remove_context_menu():
            self.status_label.setText("Successfully removed 'Convert to PDF' from the context menu for .doc and .docx files.")
        else:
            self.status_label.setText("Could not find 'Convert to PDF' in the context menu || Already removed or you didn't install in the first place")

def main():
    app = QApplication(sys.argv)
    app.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app.setStyle('Fusion')

    gui = GlossyGUI()
    gui.show()

    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
