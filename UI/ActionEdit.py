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
import sip
from UI.TabWidget import *

class ActionEdit(QDialog):
    submitted = pyqtSignal(str, str, dict)
    newId = dict()
    maxId = 0
    def __init__(self, parent, acitonId, title):
        super().__init__(parent)
        self.acitonCode = acitonId
        self.title = title
        self.configFileDirectory = "./Config"
        self.configFileName = "Actions.xml"
        self.configFilePath = os.path.join(self.configFileDirectory, self.configFileName)
        self.__initUI()

    def __initUI(self):
        self.setObjectName("Main")
        stylesheet = open("UI/SettingManager.qss", "r").read()
        self.setWindowTitle(self.title)
        self.setWindowModality(Qt.ApplicationModal)
        self.setStyleSheet(stylesheet)


        self.body = QWidget()
        self.body.setStyleSheet("")
        self.body.setObjectName("body")

        self.bodyLayout = QGridLayout(self.body)
        self.bodyLayout.setObjectName("bodyLayout")

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setSizeConstraint(QLayout.SetDefaultConstraint)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")

        self.setLayout(self.verticalLayout)

        self.initTable()

        self.verticalLayout.addWidget(self.body)
        self.bottom = QWidget()
        self.bottom.setObjectName("bottom")
        self.verticalLayout.addWidget(self.bottom)

        self.bottom_layount = QVBoxLayout()
        self.bottom.setLayout(self.bottom_layount)
        self.button_submit = QPushButton()

        #self.pushButton.setGeometry(QRect(0, 320, 171, 91))
        self.button_submit.setObjectName("submit")
        self.button_submit.setText("保存")
        self.button_submit.clicked.connect(self.submit)
        self.bottom_layount.addWidget(self.button_submit)

    def initTable(self):
        self.count = 0
        self.mainTable = Sheet()
        self.mainTable.setSelectionMode(QAbstractItemView.SingleSelection)
        # 初始化表头
        self.mainTable.setColumnCount(5)
        Item = QTableWidgetItem()
        Item.setText("变量名")
        self.mainTable.setHorizontalHeaderItem(1, Item)
        Item = QTableWidgetItem()
        Item.setText("参数名")
        self.mainTable.setHorizontalHeaderItem(2, Item)
        Item = QTableWidgetItem()
        Item.setText("变量值")
        self.mainTable.setHorizontalHeaderItem(3, Item)
        Item = QTableWidgetItem()
        Item.setText("编辑")
        self.mainTable.setHorizontalHeaderItem(0, Item)
        Item = QTableWidgetItem()
        Item.setText("是否常量")
        self.mainTable.setHorizontalHeaderItem(4, Item)
        self.mainTable.verticalHeader().hide()
        self.mainTable.setColumnWidth(0, 91)
        self.mainTable.setColumnWidth(1, 200)
        self.mainTable.setColumnWidth(2, 200)
        self.mainTable.setColumnWidth(3, 91)
        self.mainTable.setColumnWidth(3, 91)

        self.initTableData()
        self.bodyLayout.addWidget(self.mainTable)

    def initTableData(self):
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find("Action[@ActionCode='{0}']".format(self.acitonCode))
        if node is None:
            self.button_add = QPushButton()
            self.button_add.setText("新增")
            self.button_add.setObjectName("btn_add")
            self.button_add.clicked.connect(self.add)
            self.mainTable.setCellWidget(self.mainTable.rowCount(), 0, self.button_add)
            return
        items = node.findall("Variable")
        for i in range(0, len(items)):
            item = TableItem(items[i].attrib["VariableName"])
            item.setFlags(Qt.ItemIsSelectable|Qt.ItemIsDragEnabled|Qt.ItemIsUserCheckable|Qt.ItemIsEnabled)
            #item.setToolTip(items[i].text)
            self.mainTable.setItem(i, 1, item)
            item = TableItem(items[i].attrib["ParameterName"])
            item.setFlags(Qt.ItemIsSelectable|Qt.ItemIsDragEnabled|Qt.ItemIsUserCheckable|Qt.ItemIsEnabled)
            self.mainTable.setItem(i, 2, item)
            item = TableItem(items[i].attrib["Value"])
            self.mainTable.setItem(i, 3, item)

            isParameter = items[i].attrib["IsParameter"]
            item = TableItem("")
            if isParameter == "True":
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
            self.mainTable.setItem(i, 4, item)
            item = QPushButton()
            item.setText("编辑")
            self.mainTable.setCellWidget(i, 0, item)
            item.clicked.connect(lambda : self.seteditable(i))

        # 新增按钮
        self.button_add = QPushButton()
        self.button_add.setText("新增")
        self.button_add.setObjectName("btn_add")
        self.button_add.clicked.connect(self.add)
        self.mainTable.setCellWidget(self.mainTable.rowCount(), 0, self.button_add)

    def add(self):
        count = self.mainTable.rowCount() -1

        item = TableItem("")
        self.mainTable.setItem(count, 1, item)
        item = TableItem("")
        self.mainTable.setItem(count, 2, item)
        item = TableItem("")
        self.mainTable.setItem(count, 3, item)
        item = TableItem("")
        item.setCheckState(False)
        self.mainTable.setItem(count, 4, item)

        self.mainTable.setCellWidget(count + 1, 0, self.button_add)

    def seteditable(self, row):
        row = self.mainTable.selectedIndexes()[0].row()
        self.mainTable.item(row, 2).setFlags(Qt.ItemIsSelectable|Qt.ItemIsDragEnabled|Qt.ItemIsUserCheckable|Qt.ItemIsEnabled|Qt.ItemIsEditable)
        self.mainTable.item(row, 1).setFlags(Qt.ItemIsSelectable|Qt.ItemIsDragEnabled|Qt.ItemIsUserCheckable|Qt.ItemIsEnabled|Qt.ItemIsEditable)

    def submit(self):
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find("Action[@ActionCode='{0}']".format(self.acitonCode))
        if node is None:
            node = XETree.Element("Action")
            node.set("ActionCode", self.acitonCode)
        for row in range(self.mainTable.rowCount() -1):
            VarName = self.mainTable.item(row, 1).text()
            parName = self.mainTable.item(row, 2).text()
            value = self.mainTable.item(row, 3).text()
            isPara = "True" if self.mainTable.item(row, 4).checkState() == 2 else "False"
            if VarName == "":
                continue
            ele_Variable = node.find("Variable[@VariableName='{0}']".format(VarName))
            if ele_Variable is None:
                ele_Variable = XETree.Element("Variable")
                ele_Variable.set("VariableName", VarName)
                ele_Variable.set("ParameterName", parName)
                ele_Variable.set("Value", value)
                ele_Variable.set("IsParameter", isPara)
                node.append(ele_Variable)
            else:
                ele_Variable = XETree.Element("Variable")
                ele_Variable.set("VariableName", VarName)
                ele_Variable.set("ParameterName", parName)
                ele_Variable.set("Value", value)
                ele_Variable.set("IsParameter", isPara)

        indent(node)
        tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)

        sip.delete(self.mainTable)
        self.initTable()
        QMessageBox.about(self, "完成", "保存成功")
        self.close()







