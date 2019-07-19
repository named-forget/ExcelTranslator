# text-editor.py

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import xml.etree.ElementTree as XETree
import shutil
from Action.Action import Action

from FileTranslator import template as fileTranslator
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from UI import TabWidget
from UI.FillData import FillData
from UI.StyleComboBox import *
from UI.SettingManager import *
from UI.ActionManager import AcitonManager

class MainWindow(QMainWindow):
    str_WindowName = "ExcelTranslator"
    str_templatefile = "templateFile"
    str_xmlFIle = "xmlFIle"
    str_OutputFolder = "OutputFolder"
    configFileDirectory = "./Config"
    configFileName = "config.xml"
    xmlFolderPath = os.getcwd() + "/Resource/xml"
    templateFoldetPath = os.getcwd() + "/Resource/templateFile"
    configFilePath = os.path.join(configFileDirectory, configFileName)
    Settings = ["配置管理", "模板管理", "方案管理", "任务管理"]
    #屏幕尺寸
    screen_width = 0
    screen_height = 0
    def __init__(self):
        super().__init__()
        self.window_width = 1280
        self.window_height = 760
        self.initConfig()
        self.initUI()
    # 初始化窗口界面
    def initUI(self):
        #设置样式
        self.setStyleSheet(open("UI/MainWindow.qss", 'r').read())
        #获取屏幕尺寸
        screen = QDesktopWidget().screenGeometry()
        self.screen_width = screen.width()
        self.screen_height = screen.height()
        # 设置中心窗口部件为QTextEdit
        self.verticalSplitter = QSplitter(Qt.Vertical)
        self.setCentralWidget(self.verticalSplitter)

        #self.scroll.setGeometry(QtCore.QRect(100,100, 2000, 1000))


        #标签管理
        self.workTab = TabWidget.WorkTab()
        self.setObjectName("workTab")
        self.setGeometry(QtCore.QRect(0, 0, self.screen_width - 40, self.screen_height - 170))


        # self.scorllTextEdit = QScrollArea()

        self.textEdit = QTextEdit()
        self.textEdit.setText('执行记录：')
        self.textEdit.setStyleSheet("background:white;height:20%")
        self.textEdit.setReadOnly(True)

        # self.vboxScorllTextEdit = QVBoxLayout()
        # self.vboxScorllTextEdit.addWidget(self.textEdit)
        # self.scorllTextEdit.setLayout(self.vboxScorllTextEdit)

        self.verticalSplitter.addWidget(self.workTab)
        self.verticalSplitter.addWidget(self.textEdit)
        # 定义一系列的Action
        # 新建
        newAction = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_create.ico'), '内容填充')
        newAction.setShortcut('Ctrl+N')
        newAction.setStatusTip('内容填充')
        newAction.clicked.connect(self.new)

        #下拉框新增
        self.button_saveAs = QtWidgets.QPushButton()
        self.button_saveAs.setText("另存为")
        #openAction.setShortcut('Ctrl+O')
        self.button_saveAs.setStatusTip('另存为')
        self.button_saveAs.clicked.connect(self.open)

        #数值校验
        self.button_checkValue = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_run.ico'), '数值校验')
        self.button_checkValue.setText("数值校验")
        self.button_checkValue.setToolTip('数值校验')
        self.button_checkValue.clicked.connect(self.run)

        #批量运行
        self.button_multiRun = QtWidgets.QPushButton(QIcon("Resource/icon/Icon_multiRun.ico"), '批量运行')
        self.button_multiRun.setText("批量填充")
        self.button_multiRun.setShortcut('Ctrl+M')
        self.button_multiRun.setStatusTip('批量填充')
        self.button_multiRun.clicked.connect(self.multiRun)


        # 保存
        saveAction = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_save.ico'), '保存')
        saveAction.setShortcut('Ctrl+S')
        saveAction.setStatusTip('保存')
        saveAction.clicked.connect(lambda :self.save(self.currentIndex()))

        #保存全部
        saveAllAction =  QtWidgets.QPushButton()
        saveAllAction.setStatusTip('保存全部')
        saveAllAction.setText("保存全部")
        saveAllAction.clicked.connect(self.saveAll)

        #配置管理
        self.comboBox_setting = DropDownMenu()
        self.comboBox_setting.setObjectName("Setting")
        for item in self.Settings:
            self.comboBox_setting.addItem(item, "")
        self.comboBox_setting.currentIndexChanged.connect(self.showManager)




        # 添加菜单
        # 对于菜单栏，注意menuBar，menu和action三者之间的关系
        # 首先取得Qself自带的menuBar：menubar = self.menuBar()
        # 然后在menuBar里添加Menu：fileMenu = menubar.addMenu('&File')
        # 最后在Menu里添加Action：fileMenu.addAction(newAction)
        # 添加工具栏
        # 对于工具栏，同样注意ToolBar和Action之间的关系
        # 首先在Qself中添加ToolBar：tb1 = self.addToolBar('File')
        # 然后在ToolBar中添加Action：tb1.addAction(newAction)
        tb1 = self.addToolBar('Edit')
        tb1.addWidget(newAction)
        tb1.addWidget(self.button_multiRun)
        tb1.addWidget(self.button_checkValue)

        tb2 = self.addToolBar('File')
        tb2.addWidget(saveAction)
        tb2.addWidget(self.button_saveAs)
        tb2.addWidget(saveAllAction)

        tb_setting = self.addToolBar("Setting")
        tb_setting.addWidget(self.comboBox_setting)




        self.verticalSplitter.setStyleSheet("background-color: rgb(222, 222, 222);")


        self.statusBar()



        #恢复控件状态
        # setting = QSettings("./Config/setting.ini", QSettings.IniFormat)
        # index = setting.value(self.str_OutputFolder)
        # if index is not None:
        #     self.comboBox_outputfolder.setCurrentIndex(int(index))
        # index = setting.value(self.str_xmlFIle)
        self.setObjectName("MainWindow")
        self.setGeometry(0, 0, 1280, 760)
        self.setWindowTitle("Excel转换")
        self.setWindowIcon(QIcon('Resource/icon/Icon_windowIcon.ico'))
        #self.center()
        self.show()
        self.showMaximized()


        #fillDataDialog.setGeometry((self.maximumSize().width() - 500)/2,(self.maximumSize().height()-800)/2, 500, 800 )


    # 主窗口居中显示
    def center(self):
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2, (screen.height() - size.height()) / 2)

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:
        super().resizeEvent(event)
        self.window_width = event.size().width()
        self.window_height = event.size().height()


    # 定义Action对应的触发事件，在触发事件中调用self.statusBar()显示提示信息
    # 重写closeEvent
    def closeEvent(self, event):
        saved = True
        for i in range(self.workTab.count()):
            if self.workTab.widget(i).isChanged:
                str += "\n" + self.tabText(i)
                saved = False
        if not saved:
            reply = QMessageBox.question(self, '确认', \
                    '以下任务结果还未保存，确认退出？' + str, \
                    QMessageBox.Yes | QMessageBox.No, \
                    QMessageBox.No)

            if reply == QMessageBox.Yes:
                self.statusBar().showMessage('退出...')
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def showManager(self, index:int):
        if index == 1:
            self.showModeManager()
        elif index == 2:
            self.showTemplateManager()
        elif index == 3:
            self.showActionManager()
    def showModeManager(self):
        modeManager = SettingManager(self, self.str_xmlFIle, self.Settings[0])
        modeManager.raise_()
        modeManager.setGeometry((self.window_width - 900) / 2, (self.window_height - 800) / 2, 905, 800)
        modeManager.setFixedSize(905, 800)
        modeManager.show()

    def showTemplateManager(self):
        templateManager = SettingManager(self, self.str_templatefile, self.Settings[1])
        templateManager.raise_()
        templateManager.setGeometry((self.window_width - 900) / 2, (self.window_height - 800) / 2, 905, 800)
        templateManager.setFixedSize(905, 800)
        templateManager.show()

    def showActionManager(self):
        templateManager = AcitonManager(self, self.str_templatefile, self.Settings[1])
        templateManager.raise_()
        templateManager.setGeometry((self.window_width - 900) / 2, (self.window_height - 800) / 2, 905, 800)
        templateManager.setFixedSize(905, 800)
        templateManager.show()

        return
    # open
    def new(self):
        #新建任务对话框

        fillDataDialog = FillData(self)
        fillDataDialog.raise_()
        fillDataDialog.submitted.connect(self.newAction)

        fillDataDialog.setGeometry((self.window_width - 500)/2,(self.window_height -500)/2, 500, 500 )
        fillDataDialog.setFixedSize(500, 500)
        fillDataDialog.show()

    def multiRun(self):
        return

    def run(self):
        currentTab = self.currentWidget()
        #inputFile = self.button_checkValue.statusTip()
        inputFile = currentTab.objectName()
        if not os.path.isfile(inputFile):
            QMessageBox.warning(self, '警告', '请先点击新建任务', QMessageBox.Yes)
            return
        currentTab.run()


    def newAction(self, actionCode, actionName, kwargs):
        newtab = TabWidget.Tab(actionCode, actionName, kwargs)
        self.workTab.addTab(newtab, QIcon("Resource/icon/Icon_tag.ico"), "actionName")
        newtab.logGenerated.connect(self.appenText)
        newtab.run()

    def appenText(self, str):
        self.textEdit.append(str)
        self.textEdit.moveCursor(QTextCursor.End)

    def open(self):
        dir = self.comboBox_outputfolder.currentText()
        if os.path.isdir(dir):
            os.startfile(dir)
        else:
            QMessageBox.warning(self, '警告', '选择的不是一个目录', QMessageBox.Yes)

    def saveAll(self):
        for i in range(self.count()):
            result = self.save(i)
            if result is not None and result:
                continue
            else:
                break

    def save(self, index):
        currentTab = self.widget(index)
        currentTab.save()



    # cut
    def cut(self):
        cursor = self.textEdit.textCursor()
        textSelected = cursor.selectedText()
        self.copiedText = textSelected
        self.textEdit.cut()

    # about
    def about(self):
        return

    #初始化配置文件
    def initConfig(self):
        if not os.path.exists(self.configFileDirectory):
            os.mkdir(self.configFileDirectory)
        if not os.path.exists(self.configFilePath):
            #open(self.configFilePath, "wb").write(bytes("", encoding="utf-8"))
            root = XETree.Element('Root')  # 创建节点
            tree = XETree.ElementTree(root)  # 创建文档
            root.append(XETree.Element(self.str_OutputFolder))
            root.append(XETree.Element(self.str_xmlFIle))
            root.append(XETree.Element(self.str_templatefile))
            self.indent(root)  # 增加换行符
            tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)


    #给xml增加换行符
    def indent(self,elem, level=0):
        i = "\n" + level * "\t"
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = i + "\t"
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
            for elem in elem:
                self.indent(elem, level + 1)
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
        else:
            if level and (not elem.tail or not elem.tail.strip()):
                elem.tail = i

    def renameMode(self):
        currentText = self.comboBox_xmlFIle.currentText()
        currentIndex = self.comboBox_xmlFIle.currentIndex()
        text, ok = QInputDialog.getText(self, '重命名转换方案', '输入新名称：', text=currentText)
        if ok:
            tree = XETree.parse(self.configFilePath)
            root = tree.getroot()
            node = root.find(self.str_xmlFIle)
            items = node.getchildren()
            for item in items:
                if item.attrib["Name"] == currentText:
                    item.attrib["Name"] = str(text)
                    break
            self.indent(node)
            self.comboBox_xmlFIle.setItemText(currentIndex, str(text))
            tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)

    def renameTemplate(self):
        currentText = self.comboBox_templatefile.currentText()
        currentIndex = self.comboBox_templatefile.currentIndex()
        text, ok = QInputDialog.getText(self, '重命名模板', '输入新名称：', text=currentText)
        if ok:
            tree = XETree.parse(self.configFilePath)
            root = tree.getroot()
            node = root.find(self.str_templatefile)
            items = node.getchildren()
            for item in items:
                if item.attrib["Name"] == currentText:
                    item.attrib["Name"] = str(text)
                    break
            self.indent(node)
            self.comboBox_templatefile.setItemText(currentIndex, str(text))
            tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)

    def viewXml(self):
        file = self.comboBox_xmlFIle.currentData()
        if file is not  None and os.path.isfile(file):
            os.startfile(file)

    def viewTemplateFile(self):
        file = self.comboBox_templatefile.currentData()
        if file is not  None and os.path.isfile(file):
            os.startfile(file)




# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     w = self()
#     sys.exit(app.exec_())