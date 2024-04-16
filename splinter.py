import sys
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog
import docx2txt
from docx import Document


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = "PySplitter"
        self.left = 10
        self.top = 10
        self.width = 640
        self.height = 480
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.openFileNameDialog()

        self.show()
        self.close()
        sys.exit(0)

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(
            self,
            "PySplitter",
            "",
            "All Files (*);;Microsoft Word Document (*.docx);;Microsoft Word Document 2003 (*.doc);;Apple Pages Document (*.pages);;Rich Text Format File (*.rtf);;Plain Text File (*.txt)",
            options=options,
        )
        if fileName:
            extention = fileName.split(".")
            splitter(fileName, extention[-1])


def splitter(file_name, file_type):
    if file_type == "docx" or file_type == "doc":
        text = docx2txt.process(file_name).replace(".", ".\n")
        doc = Document()
        doc.add_paragraph(text)
        doc.save(file_name)

    else:
        with open(file_name, "r+") as f:
            text = f.read().replace(".", ".\n")
            f.write(text)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
