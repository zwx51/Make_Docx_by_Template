import sys
import pymssql
from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.uic import loadUi
from WordDialog import WordDialog


class MainDialog(QDialog):
    def __init__(self):
        super(MainDialog, self).__init__()
        loadUi('MainDialog.ui', self)
        self.pushButton.clicked.connect(lambda: self.showDialog())

    def showDialog(self):
        s = WordDialog()
        s.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainDialog()
    w.show()
    sys.exit(app.exec())