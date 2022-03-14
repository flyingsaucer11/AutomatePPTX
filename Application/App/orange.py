import sys

from PyQt5.QtWidgets import QApplication, QProgressBar
from PyQt5.QtWidgets import QPushButton,QFileDialog,QTextEdit, QMainWindow

from Application.App import apple


class Window(QMainWindow):
    targetfile=''

    def __init__(self):
        super().__init__()

        self.setWindowTitle('Application')
        self.setGeometry(400,200,600,250)

        self.UiComponents()
        self.show()


    def UiComponents(self):
        self.filebutton=QPushButton('Choose a workbook file', self)
        self.filebutton.setGeometry(100,15,180,30)
        self.filebutton.clicked.connect(self.clickFileOpen)

        self.txtbx = QTextEdit('No File Selected yet..', self)
        self.txtbx.setGeometry(100, 50, 450, 30)

        self.startbutton = QPushButton('Generate Presentations', self)
        self.startbutton.setGeometry(100, 105, 200, 30)
        self.startbutton.clicked.connect(self.transform)

        self.prbar = QProgressBar(self)
        self.prbar.setGeometry(100, 145, 450, 30)


    def clickFileOpen(self):
        self.open()

    def open(self):
        fileName= QFileDialog.getOpenFileName(self,'OpenFile')
        self.txtbx.setText(fileName[0])
        self.targetfile=fileName[0]

    def progress(self):
        for i in range(100+1):
            for j in range(100000):
                k=j*j
            self.prbar.setValue(i)

    def transform(self):
        Apl= apple.Apple(self.targetfile)
        Apl.getAndSetData(self.prbar)

def exec():
    App=QApplication(sys.argv)
    window=Window()
    sys.exit(App.exec())

