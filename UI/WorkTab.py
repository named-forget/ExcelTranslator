#每个tab包含两个view用于显示源文件与目标文件
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import openpyxl
import re

#列字母转数字
def letterToint(s):
    letterdict = {}
    for i in range(26):
        letterdict[chr(ord('A') + i)] = i + 1
    output = 0
    for i in range(len(s)):
        output = output * 26 + letterdict[s[i]]
    return output - 1
#数字转列字母
def intToletter(i):
    if type(i) is not int:
        return i
    str = ''
    i += 1
    while (not (i // 26 == 0 and i % 26 == 0)):

        temp = 25

        if (i % 26 == 0):
            str += chr(temp + 65)
        else:
            str += chr(i % 26 - 1 + 65)

        i //= 26
        # print(str)
    # 倒序输出拼写的字符串
    return str[::-1]


class WorkTab(QTabWidget):
    def __init__(self):
        super().__init__()
        self.__init()

    def __init(self):
        self.setTabsClosable(True)
        self.setMovable(True)
        self.tabCloseRequested.connect(lambda index: self.removeTab(index)
                                       )
        #self.tabBar().tabBarClicked.connect(lambda index: print(("this is %d" % index)))
        self.tabBar().setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabBar().customContextMenuRequested.connect(self.__showMenu)
        self.contextMenu = QMenu(self)


        self.action_CloseAll = QAction("关闭所有标签")
        self.action_CloseLeft = QAction("关闭左侧标签")
        self.action_CloseRight = QAction("关闭右侧标签")
        self.action_CloseAllButThis = QAction("关闭除当前标签以外的所有标签")

        self.contextMenu.addAction(self.action_CloseAll)
        self.contextMenu.addAction(self.action_CloseLeft)
        self.contextMenu.addAction(self.action_CloseRight)
        self.contextMenu.addAction(self.action_CloseAllButThis)

        self.action_CloseAll.triggered.connect(self.clear)
        self.action_CloseLeft.triggered.connect(self.__removeLeftTab)
        self.action_CloseRight.triggered.connect(self.__removeRightTab)
        self.action_CloseAllButThis.triggered.connect(self.__removeAllButThis)
        self.tabBar().tabBarClicked.connect(lambda  index:print(index))



    def __showMenu(self):
        self.contextMenu.exec_(QCursor.pos())
        # str = "QTabBar::tab{background-color:rbg(255,255,255,0);}" + \
        #       "QTabBar::tab:selected{color:red;background-color:rbg(255,200,255);} "

    def __removeLeftTab(self):
        index = self.currentIndex()
        for i in range(0, index):
             self.removeTab(0)

    def __removeRightTab(self):
        index = self.currentIndex()
        for i in range(index + 1, self.count()):
            self.removeTab(index + 1)
    def __removeAllButThis(self):
        self.__removeLeftTab();
        self.__removeRightTab()

    def removeTab(self, index: int) -> None:
        tab = self.widget(index)
        if tab is not None:
            super().removeTab(index)
            tab.setParent(None)
            tab = None

    def clear(self) -> None:
        for i in range(0, self.count()):
            self.removeTab(i)




class Tab(QWidget):
    sourceWidget = ""
    destWidget = ""
    sourceSheet = ""
    destSheet = ""
    filepath = ""
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.sourceWidget = QWidget(self)
        self.destWidget = QWidget(self)
        self.sourceSheet = Sheet()
        self.destSheet = Sheet()


        self.vbox = QVBoxLayout()
        self.vbox.addWidget(self.sourceWidget)
        self.vbox.addWidget(self.destWidget)
        self.vbox.setDirection(0)

        self.sbox = QVBoxLayout()
        self.sbox.addWidget(self.sourceSheet)
        self.sourceWidget.setLayout(self.sbox)

        self.dbox = QVBoxLayout()
        self.dbox.addWidget(self.destSheet)
        self.destWidget.setLayout(self.dbox)

        self.sourceWidget.setStyleSheet("background-color: rgb(85, 170, 255);")
        self.destWidget.setStyleSheet("background-color: rgb(255, 135, 137);")
        self.setLayout(self.vbox)
        # self.sourceWidget.setGeometry(QtCore.QRect(0, 0, 1000, 500))
        # self.destWidget.setGeometry(QtCore.QRect(1020, 0, 1000, 500))
        # self.sourceSheet.setGeometry(QtCore.QRect(1, 0, 1000, 500))
        # self.destSheet.setGeometry(QtCore.QRect(1, 0, 1000, 500))


    def fillLeft(self, filepath, sheetIndex):
        self.sourceSheet.fillSheetByExcelSheetIndex(filepath, sheetIndex)

    def fillRight(self, filepath, sheetIndex):
        self.destSheet.fillSheetByExcelSheetIndex(filepath, sheetIndex)



class Sheet(QTableWidget):
    rowsMark = ""
    def __init__(self):
        super().__init__()

    def fillSheetByExcelSheetIndex(self, filepath, sheetIndex):
        workbok = openpyxl.load_workbook(filepath)
        sheetNames = workbok.sheetnames
        self.__fileSheetByWorkBook(workbok[sheetNames[sheetIndex]])

    def fillSheetByExcelSheetName(self, filepath, sheetName):
        workbok = openpyxl.load_workbook(filepath)
        self.__fileSheetByWorkBook(workbok[sheetName])
        sheet = workbok[sheetName]

    def __fileSheetByWorkBook(self, sheet):
        row_count = sheet.max_row
        col_count = sheet.max_column
        for i in range(0, col_count):
            item = TableItem(value=intToletter(i))
            item.setText(intToletter(i))
            self.setHorizontalHeaderItem(i, item)

        for i in range(0, row_count):
            item = TableItem(value=str(i))
            self.setVerticalHeaderItem(i, item)

        #开始初始化表格
        self.setColumnCount(col_count)
        self.setRowCount(row_count)
        for row in sheet.iter_rows(min_row=0):
            for cell in row:
                if cell.value is not None and cell.value != "":
                    item = TableItem(value=str(cell.value))
                    self.setItem(cell.row -1, cell.col_idx -1, item)


class TableItem(QTableWidgetItem):
    dataType = "S"
    errorColor = QBrush(QColor(255, 114, 116))
    noramlColor = QBrush(QColor(255, 255, 255))
    def __int__(self):
        super().__init__()

    def __init__(self, value):
        super().__init__()
        self.setText(value)
    # def __init__(self, dataType, value):
    #     super().__init__()
    #     self.setFormatValue(value)
    def setFormatValue(self,value):
        value =  self.updateValue(value, self.dataType)
        if not self.checkData(value):
            self.setBackground(self.errorColor)
        else:
            self.setBackground(self.noramlColor)
        super().setText(str)

    def updateValue(self,value):  # 根据类型进行字符串处理，D日期，P百分数，S字符串不做处理，F或不填按数字处理
        try:
            if value is None:
                return ''
            value = str(value)
            value = value.replace(' ', '')
            if self.dataType == 'D':
                # dateFormat = '[^\d/\d/\d]'
                # '2016年3月1日至2016年3月31日'
                value = re.sub('.*至', '', value)
                value = re.sub('\D$', '', value)
                value = re.sub(r'\D', r'/', value)
            elif self.dataType == 'P':
                if type(value) == float:
                    value = str(value * 100) + '%'
                else:
                    Persentage = '[^\d%|\d.\d%]'
                    value = re.sub(Persentage, '', value)
            elif self.dataType == 'S':
                return value
            else:
                floatFormat = '[^\d|\d.\d]'
                value = re.sub(floatFormat, '', value)
                if value == '':
                    value = 0
                value = float(value)
            return value
        except Exception as e:
            print(e)
            print("ErrorValue: " + value + "Type: " + self.ataType)

    def checkData(self,value):  # 值校验，是否符合对应类型
        
        value = str(value)
        date = '^\d{4}\D\d{1,2}\D\d{1,2}\D?$'
        num = '^[-]?\d*[.]?\d*$'
        persentige = '^[-]?\d*[.]?\d*[%]$'
        if value is None or value == '':
            return False
        flag = 0
        if self.dataType == 'D':
            flag = re.search(date, value)
        elif self.dataType == 'P':
            flag = re.search(persentige, value)
        elif self.dataType == 'S':
            return True
        else:
            flag = re.search(num, value)
        if str(flag) == 'None':
            return False
        else:
            return True

    def getDataType(self):
        return self.dataType
    def setDataType(self, dataType):
        self.dataType = dataType



