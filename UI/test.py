import sys
import re
from UI import Tab

from PyQt5.QtWidgets import QApplication, QMainWindow
from openpyxl import load_workbook


def BGR_hex2BGR(bgr_hex):
    bgr_int = int(bgr_hex)
    b = bgr_int // 16**4
    g = (bgr_int - b*16**4) // 16**2
    r = bgr_int - b*16**4 - g*16**2
    return (b, g, r)


def HexToRgb(tmp):
    rgb = dict()
    opt = re.findall(r'(.{2})',tmp)
    rgb["r"] = int(opt[0], 16)
    rgb["g"] = int(opt[1], 16)
    rgb["b"] = int(opt[2], 16)
    return rgb

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = QMainWindow()
    #openfile.Ui_Dialog()
    w = Tab.Ui_MainWindow()
    w.setupUi(window)
    window.show()
    print(HexToRgb("00000000"))
    sys.exit(app.exec_())




