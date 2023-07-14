# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

class Combo(QComboBox):
    def __init__(self,title,parent):
        super(Combo, self).__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e) :
        print(e)
        if e.mimeData().hasText():
            e.accept()
        else:
            e.ignore()
    def dropEvent(self, e) :
        self.addItem(e.mimeData().text())



class Example(QWidget):
    def __init__(self):
        super(Example, self).__init__()
        self.initUI()

    def initUI(self):
        lo = QFormLayout()
        lo.addWidget(QLabel("請把左邊的文字拖曳"))
        edit = QLineEdit()
        edit.setDragEnabled(True)
        com = Combo("Button", self)
        lo.addRow(edit,com)
        self.setLayout(lo)
        self.setWindowTitle("簡單的拖曳")

def TDrop():
    import sys
    app = QApplication(sys.argv)
    Form = Example()
    Form.show()
    sys.exit(app.exec_())


def TQPixmap():
    import sys
    app = QApplication(sys.argv)
    Form = QWidget()
    lab1 = QLabel()
    lab1.setPixmap(QPixmap(r"D:/img/IMG_3200.JPG"))
    vbox = QVBoxLayout()
    vbox.addWidget(lab1)
    Form.setLayout(vbox)
    Form.setWindowTitle("QPixmap範例")
    Form.setGeometry(10,50,700,700)
    Form.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    TDrop()