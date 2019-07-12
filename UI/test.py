import sys
import os
from PyQt5.QtWidgets import QWidget, QFormLayout, QPushButton, QLineEdit, \
        QInputDialog, QApplication

class InputDialogWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QFormLayout()

        self.btn1 = QPushButton('获得列表里的选项')
        self.btn1.clicked.connect(self.getItem)
        self.le1 = QLineEdit()
        layout.addRow(self.btn1, self.le1)

        self.btn2 = QPushButton('获得字符串')
        self.btn2.clicked.connect(self.getText)
        self.le2 = QLineEdit()
        layout.addRow(self.btn2, self.le2)

        self.btn3 = QPushButton('获得整数')
        self.btn3.clicked.connect(self.getInt)
        self.le3 = QLineEdit()
        layout.addRow(self.btn3, self.le3)

        self.setLayout(layout)
        self.setWindowTitle('QInputDialog')
        self.show()

    def getItem(self):
        items = ('C', 'C++', 'Java', 'Python')
        item, ok = QInputDialog.getItem(self, 'Select Input Dialog', \
                '语言列表', items, 0, False)
        if ok and item:
            self.le1.setText(item)

    def getText(self):
        text, ok = QInputDialog.getText(self, 'Text Input Dialog', \
                '输入姓名：',text="321")
        if ok:
            self.le2.setText(str(text))

    def getInt(self):
        num, ok = QInputDialog.getInt(self, 'Integer Input Dialog', \
                '输入数字：')
        if ok:
            self.le3.setText(str(num))
print(os.getcwd())
# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     w = InputDialogWindow()
#     sys.exit(app.exec_())
