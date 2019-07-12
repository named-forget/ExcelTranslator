# text-editor.py

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import webbrowser
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QTextEdit, QAction, QApplication, \
        QMessageBox, QFileDialog, QDesktopWidget

class TextEditor(QMainWindow):
    '''
    TextEditor : 一个简单的记事本程序
    '''
    def __init__(self):
        super().__init__()
        self.copiedText = ''
        self.initUI()

    # 初始化窗口界面
    def initUI(self):
        # 设置中心窗口部件为QTextEdit
        self.textEdit = QTextEdit()
        self.setCentralWidget(self.textEdit)
        self.textEdit.setText('')

        # 定义一系列的Action
        # 退出
        exitAction = QAction(QIcon('./images/exit.png'), 'Exit', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('Exit application')
        exitAction.triggered.connect(self.close)

        # 新建
        newAction = QAction(QIcon('./images/new.png'), 'New', self)
        newAction.setShortcut('Ctrl+N')
        newAction.setStatusTip('New application')
        newAction.triggered.connect(self.__init__)

        # 打开
        openAction = QAction(QIcon('./images/open.png'), 'Open', self)
        openAction.setShortcut('Ctrl+O')
        openAction.setStatusTip('Open Application')
        openAction.triggered.connect(self.open)

        # 保存
        saveAction = QAction(QIcon('./images/save.png'), 'Save', self)
        saveAction.setShortcut('Ctrl+S')
        saveAction.setStatusTip('Save Application')
        saveAction.triggered.connect(self.save)

        # 撤销
        undoAction = QAction(QIcon('./images/undo.png'), 'Undo', self)
        undoAction.setShortcut('Ctrl+Z')
        undoAction.setStatusTip('Undo')
        undoAction.triggered.connect(self.textEdit.undo)

        # 重做
        redoAction = QAction(QIcon('./images/redo.png'), 'Redo', self)
        redoAction.setShortcut('Ctrl+Y')
        redoAction.setStatusTip('Redo')
        redoAction.triggered.connect(self.textEdit.redo)

        # 拷贝
        copyAction = QAction(QIcon('./images/copy.png'), 'Copy', self)
        copyAction.setShortcut('Ctrl+C')
        copyAction.setStatusTip('Copy')
        copyAction.triggered.connect(self.copy)

        # 粘贴
        pasteAction = QAction(QIcon('./images/paste.png'), 'Paste', self)
        pasteAction.setShortcut('Ctrl+V')
        pasteAction.setStatusTip('Paste')
        pasteAction.triggered.connect(self.paste)

        # 剪切
        cutAction = QAction(QIcon('./images/cut.png'), 'Cut', self)
        cutAction.setShortcut('Ctrl+X')
        cutAction.setStatusTip('Cut')
        cutAction.triggered.connect(self.cut)

        # 关于
        aboutAction = QAction(QIcon('./images/about.png'), 'About', self)
        aboutAction.setStatusTip('About')
        aboutAction.triggered.connect(self.about)

        # 添加菜单
        # 对于菜单栏，注意menuBar，menu和action三者之间的关系
        # 首先取得QMainWindow自带的menuBar：menubar = self.menuBar()
        # 然后在menuBar里添加Menu：fileMenu = menubar.addMenu('&File')
        # 最后在Menu里添加Action：fileMenu.addAction(newAction)
        menubar = self.menuBar()

        fileMenu = menubar.addMenu('&File')
        fileMenu.addAction(newAction)
        fileMenu.addAction(openAction)
        fileMenu.addAction(saveAction)
        fileMenu.addAction(exitAction)

        editMenu = menubar.addMenu('&Edit')
        editMenu.addAction(undoAction)
        editMenu.addAction(redoAction)
        editMenu.addAction(cutAction)
        editMenu.addAction(copyAction)
        editMenu.addAction(pasteAction)

        helpMenu = menubar.addMenu('&Help')
        helpMenu.addAction(aboutAction)

        # 添加工具栏
        # 对于工具栏，同样注意ToolBar和Action之间的关系
        # 首先在QMainWindow中添加ToolBar：tb1 = self.addToolBar('File')
        # 然后在ToolBar中添加Action：tb1.addAction(newAction)
        tb1 = self.addToolBar('File')
        tb1.addAction(newAction)
        tb1.addAction(openAction)
        tb1.addAction(saveAction)

        tb2 = self.addToolBar('Edit')
        tb2.addAction(undoAction)
        tb2.addAction(redoAction)
        tb2.addAction(cutAction)
        tb2.addAction(copyAction)
        tb2.addAction(pasteAction)

        tb3 = self.addToolBar('Exit')
        tb3.addAction(exitAction)

        # 添加状态栏，以显示每个Action的StatusTip信息
        self.statusBar()

        self.setGeometry(0, 0, 600, 600)
        self.setWindowTitle('Text Editor')
        self.setWindowIcon(QIcon('./images/text.png'))
        self.center()
        self.show()

    # 主窗口居中显示
    def center(self):
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2, (screen.height() - size.height()) / 2)

    # 定义Action对应的触发事件，在触发事件中调用self.statusBar()显示提示信息
    # 重写closeEvent
    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Confirm', \
                'Are you sure to quit without saving ?', \
                QMessageBox.Yes | QMessageBox.No, \
                QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.statusBar().showMessage('Quiting...')
            event.accept()
        else:
            event.ignore()
            self.save()
            event.accept()

    # open
    def open(self):
        self.statusBar().showMessage('Open Text Files')
        fname = QFileDialog.getOpenFileName(self, 'Open file', '/home')
        self.statusBar().showMessage('Open File')
        if fname[0]:
            f = open(fname[0], 'r')
            with f:
                data = f.read()
                self.textEdit.setText(data)

    # save
    def save(self):
        self.statusBar().showMessage('Add extension to file name')
        fname = QFileDialog.getSaveFileName(self, 'Save File')
        if (fname[0]):
            data = self.textEdit.toPlainText()
            f = open(fname[0], 'w')
            f.write(data)
            f.close()

    # copy
    def copy(self):
        cursor = self.textEdit.textCursor()
        textSelected = cursor.selectedText()
        self.copiedText = textSelected

    # paste
    def paste(self):
        self.textEdit.append(self.copiedText)

    # cut
    def cut(self):
        cursor = self.textEdit.textCursor()
        textSelected = cursor.selectedText()
        self.copiedText = textSelected
        self.textEdit.cut()

    # about
    def about(self):
        url = 'https://en.wikipedia.org/wiki/Text_editor'
        self.statusBar().showMessage('Loading url...')
        webbrowser.open(url)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = TextEditor()
    sys.exit(app.exec_())