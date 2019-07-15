import sys
import re
from UI import Tab
import numpy as np
import numpy as np

from PyQt5.QtWidgets import QApplication, QMainWindow
from openpyxl import load_workbook

text = "180,000,000.00\t20,016,000.00\r\n651,000,000.00\t359,352,000.00\r\n71,000,000.00\t71,000,000.00\r\n78,702,324.00\t78,702,324.00"
pattern = text.split("\t\t\t")
rows = pattern[0].split("\r\n")
cell = []
for i in range(len(rows)):
    cell.append(rows[i].split("\t"))

print(cell[0][0])



# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     window = QMainWindow()
#     #openfile.Ui_Dialog()
#     w = Tab.Ui_MainWindow()
#     w.setupUi(window)
#     window.show()
#     print(HexToRgb("00000000"))
#     sys.exit(app.exec_())




