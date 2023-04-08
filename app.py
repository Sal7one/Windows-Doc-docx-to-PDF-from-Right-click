import sys
import os
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QProgressBar, QFileDialog)
from PyQt5.QtCore import Qt, QProcess

label = "Convert to PDF"
python_executable = sys.executable

class GlossyGUI(QWidget):
    def __init__(self):
        super(GlossyGUI,self).__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Context Menu')
        self.setGeometry(500, 500, 400, 180)

        # Set a custom style sheet for a glossy feel
        self.setStyleSheet("""
            QPushButton {
                background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #8a2be2, stop: 1 #551a8b);
                border: 1px solid #4A4A4A;
                border-radius: 5px;
                color: white;
                padding: 15px;
                min-width: 150px;
                font-size: 18px;
            }
 
            QPushButton:hover {
                background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #6C6C6C, stop: 1 #353535);
            }
            QPushButton:pressed {
                background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #5D5D5D, stop: 1 #1E1E1E);
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
            self.process = QProcess(self)
            self.process.finished.connect(self.on_add_to_context_menu_finished)
            script_path = os.path.abspath('context_menu.py')
            self.process.start(python_executable, [script_path, "add"])

    def on_add_to_context_menu_finished(self, exit_code, exit_status):
        if exit_code == 0:
            self.status_label.setText(rf"Successfully added {label} to the context menu for .doc and .docx files.")
        else:
            self.status_label.setText(rf"An error occurred while adding {label} to the context menu.")

    def remove_from_context_menu(self):
            self.process = QProcess(self)
            self.process.finished.connect(self.on_remove_to_context_menu_finished)
            script_path = os.path.abspath('context_menu.py')
            self.process.start(python_executable, [script_path, "remove"])

    def on_remove_to_context_menu_finished(self, exit_code, exit_status):
        if exit_code == 0:
            self.status_label.setText(rf"Successfully removed {label} from the context menu for .doc and .docx files.")
        else:
            self.status_label.setText(rf"Could not find {label} in the context menu || Already removed or you didn't install in the first place")

def main():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    app.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app.setStyle('Fusion')
    gui = GlossyGUI()
    gui.show()

    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
