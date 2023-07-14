# -*- coding: utf-8 -*-
import os.path

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

class Example(QWidget):
    def __init__(self,parent=None):
        super(Example, self).__init__(parent)
        testCopyButton = QPushButton("&Copy Text")
        testPasteBtton = QPushButton("paste &Text")
        htmlCopyButton = QPushButton("C&opy HTML")
        htmlPasteBtton = QPushButton("paste &HTML")
        imageCopyButton = QPushButton("Co&py image")
        imagePasteBtton = QPushButton("paste &image")
        self.textLabel = QLabel("Original text")
        self.imageLabel = QLabel()
        self.imageLabel.setPixmap(QPixmap(os.path.join(os.path.dirname(__file__),r"d:/img/python.png")))
        layout = QGridLayout()
        layout.addWidget(testCopyButton,0 ,0)
        layout.addWidget(imageCopyButton, 0, 1)
        layout.addWidget(htmlCopyButton, 0, 2)
        layout.addWidget(testPasteBtton,1 ,0)
        layout.addWidget(imagePasteBtton, 1, 1)
        layout.addWidget(htmlPasteBtton, 1, 2)
        layout.addWidget(self.textLabel, 2, 0, 1, 2)
        layout.addWidget(self.imageLabel, 2, 2)
        self.setLayout(layout)
        self.setWindowTitle("Clipboard 範例")
        testCopyButton.clicked.connect(self.copyText)
        testPasteBtton.clicked.connect(self.pasteText)
        htmlCopyButton.clicked.connect(self.copyHtml)
        htmlPasteBtton.clicked.connect(self.pasteHtml)
        htmlCopyButton.clicked.connect(self.copyHtml)
        htmlPasteBtton.clicked.connect(self.pasteHtml)
        imageCopyButton.clicked.connect(self.copyImage)
        imagePasteBtton.clicked.connect(self.pasteImage)
        self.setWindowTitle("Clipboard 範例")

    def copyText(self):
        clipboard = QApplication.clipboard()
        clipboard.setText("AAAAAAAA")

    def pasteText(self):
        clipboard = QApplication.clipboard()
        self.textLabel.setText(clipboard.text())

    def copyImage(self):
        clipboard = QApplication.clipboard()
        clipboard.setPixmap(QPixmap(os.path.join(
            os.path.dirname(__file__), "./images/python.png")))

    def pasteImage(self):
        clipboard = QApplication.clipboard()
        self.imageLabel.setPixmap(clipboard.pixmap())

    def copyHtml(self):
        mimeData = QMimeData()
        mimeData.setHtml("<b>Bold and <font color=red>Red</font></b>")
        clipboard = QApplication.clipboard()
        clipboard.setMimeData(mimeData)

    def pasteHtml(self):
        clipboard = QApplication.clipboard()
        mimeData = clipboard.mimeData()
        if mimeData.hasHtml():
            self.textLabel.setText(mimeData.html())


'''        

'''




def main():
    import sys
    import os
    app = QApplication(sys.argv)
    Form = Example()
    Form.show()
    sys.exit(app.exec_())



if __name__ == "__main__":
    main()