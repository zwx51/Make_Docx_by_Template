import sys
import pymssql
import os

from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from PyQt5.uic import loadUi
from PyQt5.QtGui import QRegExpValidator, QIntValidator
from PymysqlUtil import PymysqlUtil
from CreateDocx import createdocx,printdocx


class WordDialog(QDialog):
    def __init__(self):
        super(WordDialog, self).__init__()
        loadUi('WordDialog.ui', self)
        # 点击事件
        self.pushButton_print.clicked.connect(self.printEvent)
        self.pushButton_docx.clicked.connect(self.docxEvent)

        # 绑定pcslabel
        self.lineEdit_pcs.textChanged[str].connect(self.onchanged_pcs)

        # 计算每板数量
        self.lineEdit_box.textChanged[str].connect(self.onchanged_qtypallet)
        self.lineEdit_pcs.textChanged[str].connect(self.onchanged_qtypallet)

        # 计算卡板数量
        self.lineEdit_qty_pallet.textChanged[str].connect(self.onchanged_palletcount)
        self.lineEdit_totalqty.textChanged[str].connect(self.onchanged_palletcount)

        # 设置限制
        self.lineEdit_pcs.setValidator(QIntValidator(0, 2147483647))
        self.lineEdit_pcs.setValidator(QIntValidator(0, 2147483647))
        self.lineEdit_box.setValidator(QIntValidator(0, 2147483647))
        self.lineEdit_qty_pallet.setValidator(QIntValidator(0, 2147483647))

    def onchanged_pcs(self, text):
        self.label_pcs.setText(text)
        try:
            pcs2 = int(self.lineEdit_pcs2.text())
            box = int(self.lineEdit_box.text())
            total = pcs2 * box
            self.lineEdit_qty_pallet.setText(str(total))

        except Exception:
            pass

    def onchanged_qtypallet(self, text):
        try:
            pcs2 = int(self.lineEdit_pcs.text())
            box = int(self.lineEdit_box.text())
            total = pcs2 * box
            self.lineEdit_qty_pallet.setText(str(total))

        except Exception:
            pass

    def onchanged_palletcount(self, text):
        try:
            qty_pallet = int(self.lineEdit_qty_pallet.text())
            totalqty = int(self.lineEdit_totalqty.text())
            if qty_pallet > totalqty:
                self.palletcount.setText(str(int(1)))
            elif totalqty % qty_pallet == 0 :
                self.palletcount.setText(str(int(totalqty/qty_pallet)))
            elif totalqty % qty_pallet != 0:
                self.palletcount.setText(str(int((totalqty / qty_pallet))+1))

        except Exception:
            pass

    def printEvent(self):
        proddate=self.dateEdit_proddate.date().toString(Qt.ISODate)
        print(proddate)
        createdocx()

    def docxEvent(self):
        sproddate = self.dateEdit_proddate.date().toString(Qt.ISODate)
        srodtime = self.lineEdit_rodtime.text()
        smodel = self.lineEdit_model.text()
        sr = self.lineEdit_r.text()
        spdo = self.lineEdit_pdo.text()
        slot = self.lineEdit_lot.text()
        stotalqty = self.lineEdit_totalqty.text()
        spcs = self.lineEdit_pcs.text()
        sbox = self.lineEdit_box.text()
        sqtypallet = self.lineEdit_qty_pallet.text()
        spallet = self.palletcount.text()
        srecdate = self.lineEdit_recdate.text()
        sremark = self.lineEdit_remark.text()
        sdestination = self.lineEdit_destination.text()
        sqtypallethead = sbox+"箱 x "+spcs+" = "
        createdocx(srodtime, sproddate, sdestination,
                   smodel, sr, spdo, slot, stotalqty,
                   spcs, sqtypallethead, sqtypallet, spallet, srecdate, sremark)
        pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = WordDialog()
    w.show()
    sys.exit(app.exec())