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
from UI import WorkTab
from UI.FillData import FillData

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
    tabarray = [WorkTab.Tab]

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
        #获取屏幕尺寸
        screen = QDesktopWidget().screenGeometry()
        self.screen_width = screen.width()
        self.screen_height = screen.height()
        # 设置中心窗口部件为QTextEdit
        self.verticalSplitter = QSplitter(Qt.Vertical)
        self.setCentralWidget(self.verticalSplitter)

        #self.scroll.setGeometry(QtCore.QRect(100,100, 2000, 1000))


        #标签管理
        self.workTab = WorkTab.WorkTab()
        self.workTab.setObjectName("workTab")
        self.workTab.setGeometry(QtCore.QRect(0, 0, self.screen_width - 40, self.screen_height - 170))
        self.workTab.setParent(self)


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
        newAction = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_create.ico'), '新建任务')
        newAction.setShortcut('Ctrl+N')
        newAction.setStatusTip('新建任务')
        newAction.clicked.connect(self.new)

        #下拉框新增
        self.openAction_OutputFolder = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_folder.ico'), '打开输出目录')
        self.openAction_OutputFolder.setText("打开输出目录")
        #openAction.setShortcut('Ctrl+O')
        self.openAction_OutputFolder.setStatusTip('新增')
        self.openAction_OutputFolder.clicked.connect(self.open)

        #运行
        self.button_run = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_run.ico'), '运行')
        self.button_run.setText("运行")
        self.button_run.setShortcut('F5')
        self.button_run.setToolTip('对选择的文件进行转换')
        self.button_run.clicked.connect(self.run)

        #批量运行
        self.button_multiRun = QtWidgets.QPushButton(QIcon("Resource/icon/Icon_multiRun.ico"), '批量运行')
        self.button_multiRun.setText("批量运行")
        self.button_multiRun.setShortcut('Ctrl+M')
        self.button_multiRun.setStatusTip('批量运行')
        self.button_multiRun.clicked.connect(self.multiRun)

        self.checkBox_directOutput = QCheckBox(text="转换结果直接输出")

        # 保存
        saveAction = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_save.ico'), '保存')
        saveAction.setShortcut('Ctrl+S')
        saveAction.setStatusTip('保存')
        saveAction.clicked.connect(lambda :self.save(self.workTab.currentIndex()))

        #保存全部
        saveAllAction =  QtWidgets.QPushButton()
        saveAllAction.setStatusTip('保存全部')
        saveAllAction.setText("保存全部")
        saveAllAction.clicked.connect(self.saveAll)


        # 撤销
        undoAction = QAction(QIcon('Resource/icon/Icon_undo.ico'), '撤销', self)
        undoAction.setShortcut('Ctrl+Z')
        undoAction.setStatusTip('撤销')
        undoAction.triggered.connect(self.textEdit.undo)

        # 重做
        redoAction = QAction(QIcon('Resource/icon/Icon_redo.ico'), '恢复', self)
        redoAction.setShortcut('Ctrl+Y')
        redoAction.setStatusTip('恢复')
        redoAction.triggered.connect(self.textEdit.redo)

        # 拷贝
        copyAction = QAction(QIcon('Resource/icon/Icon_copy.ico'), '复制', self)
        copyAction.setShortcut('Ctrl+C')
        copyAction.setStatusTip('复制')
        copyAction.triggered.connect(self.workTab.copy)

        # 粘贴
        pasteAction = QAction(QIcon('Resource/icon/Icon_paste.ico'), '粘贴', self)
        pasteAction.setShortcut('Ctrl+V')
        pasteAction.setStatusTip('粘贴')
        pasteAction.triggered.connect(self.workTab.paste)

        # 剪切
        cutAction = QAction(QIcon('Resource/icon/Icon_cut.ico'), '剪切', self)
        cutAction.setShortcut('Ctrl+X')
        cutAction.setStatusTip('剪切')
        cutAction.triggered.connect(self.cut)

        # 关于
        aboutAction = QAction(QIcon('Resource/icon/Icon_about.ico'), '关于', self)
        aboutAction.setStatusTip('关于')
        aboutAction.triggered.connect(self.about)

        #输出至：
        self.label_OutputFolder = QtWidgets.QLabel(text="输出至：")
        # self.label_OutputFolder.setGeometry(QtCore.QRect(50, 140, 121, 41))
        self.label_OutputFolder.setTextFormat(QtCore.Qt.AutoText)
        self.label_OutputFolder.setAlignment(QtCore.Qt.AlignRight)
        self.label_OutputFolder.setObjectName("label_OutputFolder")

        self.comboBox_outputfolder = QtWidgets.QComboBox()
        # self.comboBox_outputfolder.setGeometry(QtCore.QRect(0, 160, 87, 22))
        self.comboBox_outputfolder.setEditable(False)
        self.comboBox_outputfolder.setObjectName("comboBox_outputfolder")
        self.comboBox_outputfolder.addItem("新增输出目录...", "New")
        self.comboBox_outputfolder.setMaxVisibleItems(10)

        self.initConboBox(self.str_OutputFolder, self.comboBox_outputfolder)
        self.comboBox_outputfolder.setCurrentIndex(1)
        self.comboBox_outputfolder.currentIndexChanged.connect(lambda  :self.EventCombox_Output(self.comboBox_outputfolder, self.str_OutputFolder))

        # 输出至：
        self.label_xmlFIle = QtWidgets.QLabel(text="文件转换方案：")
        # self.label_xmlFIle.setGeometry(QtCore.QRect(50, 140, 121, 41))
        self.label_xmlFIle.setTextFormat(QtCore.Qt.AutoText)
        self.label_xmlFIle.setAlignment(QtCore.Qt.AlignRight)
        self.label_xmlFIle.setObjectName("label_xmlFIle")

        self.comboBox_xmlFIle = QtWidgets.QComboBox()
        # self.comboBox_xmlFIle.setGeometry(QtCore.QRect(0, 160, 87, 22))
        self.comboBox_xmlFIle.setEditable(False)
        self.comboBox_xmlFIle.setObjectName("comboBox_xmlFIle")
        self.comboBox_xmlFIle.addItem("新增方案...", "New")
        self.comboBox_xmlFIle.setMaxVisibleItems(10)

        self.initConboBox(self.str_xmlFIle, self.comboBox_xmlFIle)
        self.comboBox_xmlFIle.setCurrentIndex(1)
        self.comboBox_xmlFIle.currentIndexChanged.connect(lambda: self.EventCombox_Output(self.comboBox_xmlFIle, self.str_xmlFIle))

        self.action_renameXml =  QAction(QIcon('Resource/icon/Icon_edit.ico'), '修改当前方案名称')
        self.action_renameXml.setStatusTip('修改当前方案名称')
        self.action_renameXml.triggered.connect(self.renameMode)

        self.button_ViewXml = QtWidgets.QPushButton(text="查看",)
        self.button_ViewXml.setText("查看配置文件")
        self.button_ViewXml.clicked.connect(self.viewXml)


        #模板
        self.label_templatefile = QtWidgets.QLabel(text="目标模板文件：")
        # self.label_templatefile.setGeometry(QtCore.QRect(50, 140, 121, 41))
        self.label_templatefile.setTextFormat(QtCore.Qt.AutoText)
        self.label_templatefile.setAlignment(QtCore.Qt.AlignRight)
        self.label_templatefile.setObjectName("label_templatefile")

        self.comboBox_templatefile = QtWidgets.QComboBox()
        # self.comboBox_templatefile.setGeometry(QtCore.QRect(0, 160, 87, 22))
        self.comboBox_templatefile.setEditable(False)
        self.comboBox_templatefile.setObjectName("comboBox_templatefile")
        self.comboBox_templatefile.addItem("新增目标模板文件...", "New")
        self.comboBox_templatefile.setMaxVisibleItems(10)

        self.initConboBox(self.str_templatefile, self.comboBox_templatefile)
        self.comboBox_templatefile.setCurrentIndex(1)
        self.comboBox_templatefile.currentIndexChanged.connect(
            lambda: self.EventCombox_Output(self.comboBox_templatefile, self.str_templatefile))

        self.action_renametemplate = QAction(QIcon('Resource/icon/Icon_edit.ico'), '修改当前模板名称')
        self.action_renametemplate.setStatusTip('修改当前模板名称')
        self.action_renametemplate.triggered.connect(self.renameTemplate)

        self.button_ViewTemplateFile = QtWidgets.QPushButton(text="查看")
        self.button_ViewTemplateFile.setText("查看目标模板文件")
        self.button_ViewTemplateFile.clicked.connect(self.viewTemplateFile)
        # 添加菜单
        # 对于菜单栏，注意menuBar，menu和action三者之间的关系
        # 首先取得Qself自带的menuBar：menubar = self.menuBar()
        # 然后在menuBar里添加Menu：fileMenu = menubar.addMenu('&File')
        # 最后在Menu里添加Action：fileMenu.addAction(newAction)
        menubar = self.menuBar()

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
        # 首先在Qself中添加ToolBar：tb1 = self.addToolBar('File')
        # 然后在ToolBar中添加Action：tb1.addAction(newAction)
        tb1 = self.addToolBar('File')
        tb1.addWidget(newAction)
        tb1.addWidget(self.button_run)
        tb1.addWidget(saveAction)
        tb1.addWidget(saveAllAction)

        tb2 = self.addToolBar('Edit')
        tb2.addWidget(self.button_multiRun)
        tb2.addWidget(self.checkBox_directOutput)
        # tb2.addAction(undoAction)
        # tb2.addAction(redoAction)
        # tb2.addAction(cutAction)
        # tb2.addAction(copyAction)
        # tb2.addAction(pasteAction)

        tb_Output = self.addToolBar('Output')
        tb_Output.addWidget(self.label_OutputFolder)
        tb_Output.addWidget((self.comboBox_outputfolder))
        tb_Output.addWidget(self.openAction_OutputFolder)

        tb_Config = self.addToolBar('Config')
        tb_Config.addWidget(self.label_xmlFIle)
        tb_Config.addAction(self.action_renameXml)
        tb_Config.addWidget((self.comboBox_xmlFIle))
        tb_Config.addWidget(self.button_ViewXml)

        tb_templatefile = self.addToolBar('templatefile')
        tb_templatefile.addWidget(self.label_templatefile)
        tb_templatefile.addAction(self.action_renametemplate)
        tb_templatefile.addWidget((self.comboBox_templatefile))
        tb_templatefile.addWidget(self.button_ViewTemplateFile)
        # 控件样式调整

        # self.setStyleSheet("background-color: rgb(222, 222, 222);")
        # tb1.setStyleSheet("background:white;")
        # tb2.setStyleSheet("background:white;")
        # tb_Config.setStyleSheet("background:white;")
        # tb_Output.setStyleSheet("background:white;")
        # tb_templatefile.setStyleSheet("background:white;")
        self.verticalSplitter.setStyleSheet("background-color: rgb(222, 222, 222);")
        self.label_OutputFolder.setStyleSheet('text-align: center;'
                                              'margin-top:11px;'
                                              'margin-bottom:5px')
        self.label_xmlFIle.setStyleSheet('text-align: center;'
                                         'margin-top:11px;'
                                         'margin-bottom:5px')
        self.label_templatefile.setStyleSheet('text-align: center;'
                                              'margin-top:11px;'
                                              'margin-bottom:5px;')
        # self.self.openAction_OutputFolder.setContentsMargins(15,0,0,0)
        self.checkBox_directOutput.setStyleSheet("margin-left:15px;padding:5px")
        self.openAction_OutputFolder.setStyleSheet("margin-left:15px;padding:5px")
        self.button_ViewTemplateFile.setStyleSheet("margin-left:15px;padding:8px")
        self.button_ViewXml.setStyleSheet("margin-left:15px;padding:8px")

        self.comboBox_outputfolder.setStyleSheet("width:200px;padding:5px")
        self.comboBox_templatefile.setStyleSheet("width:200px;padding:5px")
        self.comboBox_xmlFIle.setStyleSheet("width:100px;padding:5px")


        self.statusBar()



        #恢复控件状态
        setting = QSettings("./Config/setting.ini", QSettings.IniFormat)
        index = setting.value(self.str_OutputFolder)
        if index is not None:
            self.comboBox_outputfolder.setCurrentIndex(int(index))
        index = setting.value(self.str_xmlFIle)
        if index is not None:
            self.comboBox_xmlFIle.setCurrentIndex(int(index))
        index = setting.value(self.str_templatefile)
        if index is not None:
            self.comboBox_templatefile.setCurrentIndex(int(index))
        index = setting.value("isChekDirecOutput")
        if index is not None:
            self.checkBox_directOutput.setCheckState(int(index))

        self.setObjectName("self")
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
        setting = QSettings("./Config/setting.ini", QSettings.IniFormat)
        setting.setValue(self.str_OutputFolder, self.comboBox_outputfolder.currentIndex())
        setting.setValue(self.str_xmlFIle, self.comboBox_xmlFIle.currentIndex())
        setting.setValue(self.str_templatefile, self.comboBox_templatefile.currentIndex())
        setting.setValue("isChekDirecOutput", self.checkBox_directOutput.checkState())
        str = ""
        saved = True
        for i in range(self.workTab.count()):
            if self.workTab.widget(i).isChanged:
                str += "\n" + self.workTab.tabText(i)
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

    # open
    def new(self):
        #新建任务对话框
        fillDataDialog = FillData(self)
        fillDataDialog.raise_()
        fillDataDialog.submitted.connect(self.newAction)

        fillDataDialog.setGeometry((self.window_width - 500)/2,(self.window_height -500)/2, 500, 500 )
        fillDataDialog.setFixedSize(500, 500)
        fillDataDialog.show()

    def run(self):
        currentTab = self.workTab.currentWidget()
        #inputFile = self.button_run.statusTip()
        inputFile = currentTab.objectName()
        if not os.path.isfile(inputFile):
            QMessageBox.warning(self, '警告', '请先点击新建任务', QMessageBox.Yes)
            return
        currentTab.run()


    def multiRun(self):
        return
        # foder_Exolorer = QFileDialog.getExistingDirectory(self, "选择文件夹", "")
        # if foder_Exolorer != "":
        #     outputFolder = self.comboBox_outputfolder.currentText()
        #     xmlfilepath = self.comboBox_xmlFIle.currentData()
        #     templatefile = self.comboBox_templatefile.currentData()
        #     reply = QMessageBox.question(self, '确认', \
        #                                  '将会转换' + foder_Exolorer + "下的所有文件，是否继续", \
        #                                  QMessageBox.Yes | QMessageBox.No, \
        #                                  QMessageBox.No)
        #
        #     if reply == QMessageBox.Yes:
        #         self.textEdit.append("开始批量转换：{0} 方案：{1} 目标模板文件：{2} 输出至：{3}".format(foder_Exolorer,
        #                                                                          self.comboBox_xmlFIle.currentText(),
        #                                                                          self.comboBox_templatefile.currentText(),
        #                                                                          outputFolder))
        #         for parent, dirnames, filenames in os.walk(foder_Exolorer, followlinks=True):
        #             for filename in filenames:
        #                 if filename[-5:] == ".xlsx":
        #                     inputFile = os.path.join(parent, filename)
        #                     outputFile = os.path.join(outputFolder, filename)
        #                     try:
        #                         self.addTab(inputFile)
        #                         tab = self.workTab.findChild(WorkTab.Tab, inputFile)
        #                         tab.run()
        #                     except ValueError as e:
        #                         self.textEdit.append(e)
        #
        #     else:
        #         return
    def newAction(self, actionCode, actionName, kwargs):
        newtab = WorkTab.Tab(actionCode, actionName, kwargs)
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
        for i in range(self.workTab.count()):
            result = self.save(i)
            if result is not None and result:
                continue
            else:
                break

    def save(self, index):
        currentTab = self.workTab.widget(index)
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

    def EventCombox_Output(self, combobox, flag):
        folder_Exolorer = ""
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
                                                                                  os.path.basename(folder_Exolorer)))
                        combobox.setCurrentIndex(len(items) + 1)
                    elif flag == self.str_templatefile:
                        if not os.path.exists(self.templateFoldetPath):
                            os.mkdir(self.templateFoldetPath)
                        shutil.copy(folder_Exolorer, self.templateFoldetPath + "/" + os.path.basename(folder_Exolorer))
                        combobox.addItem(os.path.basename(folder_Exolorer)[0:-5], os.path.join(self.templateFoldetPath,
                                                                                    os.path.basename(folder_Exolorer)))
                        combobox.setCurrentIndex(len(items) + 1)
                    self.indent(node)
                    tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)
            else:
                combobox.setCurrentIndex(1)
        #print(combobox.currentText())

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

    #读取配置文件，初始化下拉框
    def initConboBox(self, section, combobox):
        isExistsSection = False
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find(section)
        items = node.getchildren()
        for i in range(0, len(items)):
            if section == self.str_xmlFIle:
                combobox.addItem(items[i].attrib["Name"], items[i].text)
            elif section == self.str_OutputFolder:
                combobox.addItem(items[i].text, str(i))
            elif section == self.str_templatefile:
                combobox.addItem(items[i].attrib["Name"], items[i].text)

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