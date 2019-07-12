# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'openfile.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!
from FileTranslator import template
from PyQt5.QtGui import QIcon
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QTextEdit, QAction, QApplication, \
        QMessageBox, QFileDialog, QDesktopWidget, QWidget

class Ui_Dialog(QWidget):

    def __init__(self):
        super().__init__()
        self.copiedText = ''
        self.setupUi(self)
        self.setWindowIcon(QIcon("Resource/icon/Icon_OpenTitle.ico"))
    def setupUi(self, Dialog):
        Dialog.setObjectName("打开文件")
        Dialog.resize(786, 273)
        Dialog.setWindowModality(QtCore.Qt.ApplicationModal)
        #确定按钮
        self.button_ok = QtWidgets.QPushButton(Dialog)
        self.button_ok.setGeometry(QtCore.QRect(500, 220, 93, 28))
        self.button_ok.setObjectName("button_ok")

        #取消按钮
        self.button_cancel = QtWidgets.QPushButton(Dialog)
        self.button_cancel.setGeometry(QtCore.QRect(620, 220, 93, 28))
        self.button_cancel.setObjectName("button_cancel")
        self.button_cancel.setShortcut("Escape")
        #源文件
        self.label_SourceFile = QtWidgets.QLabel(Dialog)
        self.label_SourceFile.setGeometry(QtCore.QRect(40, 40, 121, 41))
        self.label_SourceFile.setTextFormat(QtCore.Qt.AutoText)
        self.label_SourceFile.setAlignment(QtCore.Qt.AlignCenter)
        self.label_SourceFile.setObjectName("label_SourceFile")

        self.lineEdit_SourceFile = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_SourceFile.setGeometry(QtCore.QRect(160, 50, 431, 21))
        self.lineEdit_SourceFile.setObjectName("lineEdit_SourceFile")

        self.button_SourceFile = QtWidgets.QPushButton(Dialog)
        self.button_SourceFile.setGeometry(QtCore.QRect(620, 50, 93, 28))
        self.button_SourceFile.setObjectName("button_SourceFile")

        #xml文件
        self.label_xmlFIle = QtWidgets.QLabel(Dialog)
        self.label_xmlFIle.setGeometry(QtCore.QRect(50, 90, 121, 41))
        self.label_xmlFIle.setTextFormat(QtCore.Qt.AutoText)
        self.label_xmlFIle.setAlignment(QtCore.Qt.AlignCenter)
        self.label_xmlFIle.setObjectName("label_xmlFIle")

        self.lineEdit_xmlFIle = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_xmlFIle.setGeometry(QtCore.QRect(160, 100, 431, 21))
        self.lineEdit_xmlFIle.setObjectName("lineEdit_xmlFIle")

        self.button_xmlFIle = QtWidgets.QPushButton(Dialog)
        self.button_xmlFIle.setGeometry(QtCore.QRect(620, 100, 93, 28))
        self.button_xmlFIle.setObjectName("button_xmlFIle")

        #输出文件夹
        self.label_OutputFolder = QtWidgets.QLabel(Dialog)
        self.label_OutputFolder.setGeometry(QtCore.QRect(50, 140, 121, 41))
        self.label_OutputFolder.setTextFormat(QtCore.Qt.AutoText)
        self.label_OutputFolder.setAlignment(QtCore.Qt.AlignCenter)
        self.label_OutputFolder.setObjectName("label_OutputFolder")

        self.lineEdit_OutputFolder = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_OutputFolder.setGeometry(QtCore.QRect(160, 150, 431, 21))
        self.lineEdit_OutputFolder.setObjectName("lineEdit_OutputFolder")

        self.button_OutputFolder = QtWidgets.QPushButton(Dialog)
        self.button_OutputFolder.setGeometry(QtCore.QRect(620, 150, 93, 28))
        self.button_OutputFolder.setShortcut("")
        self.button_OutputFolder.setObjectName("button_OutputFolder")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
        Dialog.show()

        #设置图标
        self.button_SourceFile.setIcon(QIcon("Resource/icon/Icon_folder.ico"))
        self.button_xmlFIle.setIcon(QIcon("Resource/icon/Icon_folder.ico"))
        self.button_OutputFolder.setIcon(QIcon("Resource/icon/Icon_folder.ico"))
        #设置事件
        self.button_OutputFolder.clicked.connect(lambda: self.eventOpenFloder("OutputFolder"))
        self.button_xmlFIle.clicked.connect(lambda: self.eventOpenFloder("xmlFIle"))
        self.button_ok.clicked.connect(lambda: self.eventOK())
        self.button_SourceFile.clicked.connect(lambda :self.eventOpenFloder("SourceFile"))
        self.button_cancel.clicked.connect(lambda :self.close())

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "打开文件"))
        self.button_ok.setWhatsThis(_translate("Dialog", "button_ok"))
        self.button_ok.setText(_translate("Dialog", "确定"))
        self.button_cancel.setWhatsThis(_translate("Dialog", "button_ok"))
        self.button_cancel.setText(_translate("Dialog", "取消"))
        self.label_SourceFile.setText(_translate("Dialog", "待转换文件"))
        self.label_xmlFIle.setText(_translate("Dialog", "xml模板"))
        self.label_OutputFolder.setText(_translate("Dialog", "输出目录"))
        self.button_SourceFile.setWhatsThis(_translate("Dialog", "button_ok"))
        self.button_SourceFile.setText(_translate("Dialog", ".."))
        self.button_xmlFIle.setWhatsThis(_translate("Dialog", "button_ok"))
        #self.button_xmlFIle.setText(_translate("Dialog", ".."))
        self.button_OutputFolder.setWhatsThis(_translate("Dialog", "button_ok"))
        #self.button_OutputFolder.setText(_translate("Dialog", ".."))

    def eventOK(self):
        input = self.lineEdit_SourceFile.text()
        output = self.lineEdit_OutputFolder.text()
        xml = self.lineEdit_xmlFIle.text()

        self.button_ok.setDisabled(True)
        template.main(input, output, xml)
        self.button_ok.setDisabled(True)
        self.close()
    # def eventCancel(self):
    #     self.close()

    def eventOpenFloder(self, flag):
        #self.statusBar().showMessage('Open Text Files')
        if flag == "SourceFile":
            foder_Exolorer = QFileDialog.getExistingDirectory(self, "选择文件夹", "")
            if foder_Exolorer != "":
                self.lineEdit_SourceFile.setText(foder_Exolorer)

        elif flag == "xmlFIle":
            file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择文件', filter=("*.xml"))
            if file_Exolorer[0]:
                f = open(file_Exolorer[0], 'r')
                with f:
                    data = f.name
                    self.lineEdit_xmlFIle.setText(data)

        elif flag == "OutputFolder":
            foder_Exolorer = QFileDialog.getExistingDirectory(self, "选择文件夹", "")
            if foder_Exolorer != "":
                self.lineEdit_OutputFolder.setText(foder_Exolorer)

        #self.statusBar().showMessage('Open File')

