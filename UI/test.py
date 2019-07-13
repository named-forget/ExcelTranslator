import sys

from UI import Tab

from PyQt5.QtWidgets import QApplication, QMainWindow
from openpyxl import load_workbook

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = QMainWindow()
    #openfile.Ui_Dialog()
    w = Tab.Ui_MainWindow()
    w.setupUi(window)
    window.show()
    sys.exit(app.exec_())
    print("123.xlsx")



