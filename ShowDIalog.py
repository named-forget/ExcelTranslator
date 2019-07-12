import sys

from UI import MainWindow

from PyQt5.QtWidgets import QApplication, QMainWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)
    #openfile.Ui_Dialog()
    w = MainWindow.MainWindow()
    sys.exit(app.exec_())

