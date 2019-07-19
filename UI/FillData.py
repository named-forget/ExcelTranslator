# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Dialog.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import os
import xml.etree.ElementTree as XETree
import shutil
from UI.StyleComboBox import StyledComboBox

class FillData(QDialog):
    submitted = pyqtSignal(str, str, dict)
    def __init__(self, parent):
        super().__init__(parent)
        self.str_templatefile = "templateFile"
        self.str_xmlFIle = "xmlFIle"
        self.str_OutputFolder = "OutputFolder"
        self.configFileDirectory = "./Config"
        self.configFileName = "config.xml"
        self.configFilePath = os.path.join(self.configFileDirectory, self.configFileName)

        self.__initUI()

    def __initUI(self):
        self.setObjectName("Main")
        stylesheet = open("UI/Dialog.qss", "r").read()
        self.setWindowTitle("内容填充")
        self.setWindowModality(Qt.ApplicationModal)
        self.setStyleSheet(stylesheet)

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setSizeConstraint(QLayout.SetDefaultConstraint)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.setLayout(self.verticalLayout)


        self.body = QWidget()
        self.body.setStyleSheet("")
        self.body.setObjectName("body")


        #添加下拉框
        self.label_selectFile = QLabel(text="原始文件：")
        self.label_selectFile.setAlignment(Qt.AlignCenter)
        self.label_selectFile.setObjectName("label_selectFile")

        self.lineEdit_selectFile = QLineEdit()
        self.lineEdit_selectFile.setText("文件路径")
        self.lineEdit_selectFile.setObjectName("lineEdit_selectFile")
        self.lineEdit_selectFile.setReadOnly(True)

        self.button_selectFIle = QPushButton()
        self.button_selectFIle.setText("选择文件")
        self.button_selectFIle.clicked.connect(self.open)

        #添加下拉框
        self.label_OutputFolder = QLabel(text="输出目录：")
        self.label_OutputFolder.setAlignment(Qt.AlignCenter)
        self.label_OutputFolder.setObjectName("label_OutputFolder")

        self.comboBox_outputfolder = StyledComboBox()
        self.comboBox_outputfolder.setEditable(False)
        self.comboBox_outputfolder.setObjectName("comboBox_outputfolder")
        self.comboBox_outputfolder.addItem("新增输出目录...", "New")
        self.comboBox_outputfolder.setMaxVisibleItems(10)

        self.initComboBox(self.str_OutputFolder, self.comboBox_outputfolder)
        self.comboBox_outputfolder.setCurrentIndex(1)
        self.comboBox_outputfolder.currentIndexChanged.connect(lambda  :self.EventCombox_Output(self.comboBox_outputfolder, self.str_OutputFolder))

        # 输出至：
        self.label_xmlFIle = QLabel(text="转换方案：")
        self.label_xmlFIle.setAlignment(Qt.AlignCenter)
        self.label_xmlFIle.setObjectName("label_xmlFIle")

        self.comboBox_xmlFIle = StyledComboBox()
        self.comboBox_xmlFIle.setEditable(False)
        self.comboBox_xmlFIle.setObjectName("comboBox_xmlFIle")
        self.comboBox_xmlFIle.addItem("新增方案...", "New")
        self.comboBox_xmlFIle.setMaxVisibleItems(10)

        self.initComboBox(self.str_xmlFIle, self.comboBox_xmlFIle)
        self.comboBox_xmlFIle.setCurrentIndex(1)
        self.comboBox_xmlFIle.currentIndexChanged.connect(lambda: self.EventCombox_Output(self.comboBox_xmlFIle, self.str_xmlFIle))
        #模板
        self.label_templatefile = QLabel(text="目标文件：")
        self.label_templatefile.setAlignment(Qt.AlignCenter)
        self.label_templatefile.setObjectName("label_templatefile")

        self.comboBox_templatefile = StyledComboBox()
        self.comboBox_templatefile.setEditable(False)
        self.comboBox_templatefile.setObjectName("comboBox_templatefile")
        self.comboBox_templatefile.addItem("新增目标模板文件...", "New")
        self.comboBox_templatefile.setMaxVisibleItems(10)

        self.initComboBox(self.str_templatefile, self.comboBox_templatefile)
        self.comboBox_templatefile.setCurrentIndex(1)
        self.comboBox_templatefile.currentIndexChanged.connect(
            lambda: self.EventCombox_Output(self.comboBox_templatefile, self.str_templatefile))

        self.bodyLayout = QGridLayout(self.body)
        self.bodyLayout.setObjectName("bodyLayout")
        self.bodyLayout.addWidget(self.label_selectFile, 0, 0)
        self.bodyLayout.addWidget(self.lineEdit_selectFile, 0, 1)
        self.bodyLayout.addWidget(self.button_selectFIle, 0, 2)
        self.bodyLayout.addWidget(self.label_OutputFolder, 1, 0)
        self.bodyLayout.addWidget(self.comboBox_outputfolder, 1, 1)
        self.bodyLayout.addWidget(self.label_xmlFIle, 2, 0)
        self.bodyLayout.addWidget(self.comboBox_xmlFIle, 2, 1)
        self.bodyLayout.addWidget(self.label_templatefile, 3, 0)
        self.bodyLayout.addWidget(self.comboBox_templatefile, 3, 1)



        self.verticalLayout.addWidget(self.body)
        self.bottom = QWidget()
        self.bottom.setObjectName("bottom")
        self.verticalLayout.addWidget(self.bottom)

        self.bottom_layount = QVBoxLayout()
        self.bottom.setLayout(self.bottom_layount)
        self.button_submit = QPushButton()
        #self.pushButton.setGeometry(QRect(0, 320, 171, 91))
        self.button_submit.setObjectName("submit")
        self.button_submit.setText("确认")
        self.button_submit.clicked.connect(self.submit)
        self.bottom_layount.addWidget(self.button_submit)



    def initComboBox(self, section, combobox):

        isExistsSection = False
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find(section)
        items = node.getchildren()
        model = combobox.model()
        for i in range(0, len(items)):
            if section == self.str_xmlFIle:
                combobox.addItem(items[i].attrib["Name"], items[i].text)
            elif section == self.str_OutputFolder:
                combobox.addItem(items[i].text, str(i))
            elif section == self.str_templatefile:
                combobox.addItem(items[i].attrib["Name"], items[i].text)

    def EventCombox_Output(self, combobox, flag):
        folder_Exolorer = ""
        print(combobox.currentText(), combobox.currentData())
        if combobox.currentData() == "New":
            if flag != self.str_OutputFolder:
                if flag == self.str_templatefile:
                    file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择文件', filter=("*.xlsx"))
                elif flag == self.str_xmlFIle:
                    file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择xml配置文件', filter=("*.xml"))
                if file_Exolorer[0]:
                    f = open(file_Exolorer[0], 'r')
                    with f:
                        folder_Exolorer = f.name
            elif flag == "OutputFolder":
                folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "")

            if folder_Exolorer != "":
                isExistsItem = False
                existsName = ""
                tree = XETree.parse(self.configFilePath)
                root = tree.getroot()
                node = root.find(flag)
                items = node.getchildren()
                for item in items:
                    if item.text == folder_Exolorer:
                        isExistsItem = True
                        if flag != self.str_OutputFolder:
                            existsName = ": " + item.attrib["Name"]
                if isExistsItem:
                    QMessageBox.warning(self, '警告', '该项已存在' + existsName, QMessageBox.Yes)
                    combobox.setCurrentIndex(1)
                else:
                    if flag == self.str_OutputFolder:
                        element = XETree.Element("Folder")
                        element.text = folder_Exolorer
                    elif flag == self.str_xmlFIle:
                        element = XETree.Element("File")
                        element.set("Name", "方案" + str(len(items) + 1))
                        element.text = self.xmlFolderPath + "/" + os.path.basename(folder_Exolorer)
                    else:
                        element = XETree.Element("File")
                        element.set("Name", os.path.basename(folder_Exolorer)[0:-5])
                        element.text = self.templateFoldetPath + "/" + os.path.basename(folder_Exolorer)
                    node.append(element)
                    if flag == self.str_OutputFolder:
                        combobox.addItem(folder_Exolorer, str(len(items) + 1))
                        combobox.setCurrentIndex(len(items) + 1)
                    elif flag == self.str_xmlFIle:
                        if not os.path.exists(self.xmlFolderPath):
                            os.mkdir(self.xmlFolderPath)
                        shutil.copy(folder_Exolorer, self.xmlFolderPath + "/" + os.path.basename(folder_Exolorer))
                        combobox.addItem("方案" + str(len(items) + 1), os.path.join(self.xmlFolderPath,
                                                                                  os.path.basename(
                                                                                      folder_Exolorer)))
                        combobox.setCurrentIndex(len(items) + 1)
                    elif flag == self.str_templatefile:
                        if not os.path.exists(self.templateFoldetPath):
                            os.mkdir(self.templateFoldetPath)
                        shutil.copy(folder_Exolorer,
                                    self.templateFoldetPath + "/" + os.path.basename(folder_Exolorer))
                        combobox.addItem(os.path.basename(folder_Exolorer)[0:-5],
                                         os.path.join(self.templateFoldetPath,
                                                      os.path.basename(folder_Exolorer)))
                        combobox.setCurrentIndex(len(items) + 1)
                    self.indent(node)
                    tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)
            else:
                combobox.setCurrentIndex(1)
        # print(combobox.currentText())

    def open(self) -> None:
        file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择文件', filter=("*.xlsx"))
        if file_Exolorer[0]:
            self.lineEdit_selectFile.setText(file_Exolorer[0])

    def submit(self):
        filepath = self.lineEdit_selectFile.text()
        outputFolder = self.comboBox_outputfolder.currentText()
        xmlfilepath = self.comboBox_xmlFIle.currentData()
        templatefile = self.comboBox_templatefile.currentData()
        parameters = dict()
        parameters["inputFile"] = filepath
        parameters["XmlFile"] = xmlfilepath
        parameters["outputFile"] = outputFolder
        parameters["templateFile"] = templatefile
        self.hide()
        self.submitted.emit("FillData", "内容填充",parameters)
        self.close()






