#每个tab包含两个view用于显示源文件与目标文件
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import openpyxl
import re
import xml.etree.cElementTree as XETree


class WorkTab(QTabWidget):
    signal_changed = pyqtSignal(int, bool)

    def __init__(self):
        super().__init__()
        self.__init()

    def __init(self):
        self.setTabsClosable(True)
        self.setMovable(True)
        self.tabCloseRequested.connect(lambda index: self.removeTab(index)
                                       )
        #设置最大宽度
        self.setStyleSheet("QTabBar::tab{max-width:150px}")
        #self.tabBar().tabBarClicked.connect(lambda index: print(("this is %d" % index))
        # str = "QTabBar::tab{background-color:rbg(255,255,255,0);}" + \
        #       "QTabBar::tab:selected{color:red;background-color:rbg(255,200,255);} "
        self.tabBar().setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabBar().customContextMenuRequested.connect(self.__showMenu)
        self.contextMenu = QMenu(self)
        self.contextMenu.setStyleSheet("background-color:rbg(0,0,255);"
                                       "selection-background-color: rgb(85, 170, 255);")

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

        #连接信号
        self.signal_changed.connect(self.setTabIconByStatus)

    def __showMenu(self):
        self.contextMenu.exec_(QCursor.pos())

    def __removeLeftTab(self):
        index = self.currentIndex()
        for i in range(0, index):
            result = self.removeTab(0)

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
            str = ""
            saved = True
            if tab.isChanged:
                str += "\n" + self.tabText(index)
                saved = False
            if not saved:
                reply = QMessageBox.question(self, '确认', \
                                             '以下任务结果还未保存，确认退出？' + str, \
                                             QMessageBox.Yes | QMessageBox.No, \
                                             QMessageBox.No)

                if reply == QMessageBox.Yes:
                    super().removeTab(index)
                    tab.setParent(None)
                    tab = None
                    return True
            else:
                super().removeTab(index)
                tab.setParent(None)
                tab = None

    def clear(self) -> None:
        self.__removeAllButThis()
        self.removeTab(self.currentIndex())

    def setTabIconByStatus(self, index, status: bool):
        print(index, status)
        if not status:
            self.setTabIcon(index, QIcon("Resource/icon/Icon_tag.ico"))
        else:
            self.setTabIcon(index, QIcon("Resource/icon/Icon_UnSave.png"))
    def setTabStatus(self, index:int, status):
        self.widget(index).isChanged = status
        print(index,status,"this is sender")
        self.signal_changed.emit(index, status)

    def changeEvent(self, event: QEvent) -> None:
        print(event)




class Tab(QWidget):
    sourceWidget = ""
    destWidget = ""
    sourceSheet = ""
    destSheet = ""
    filepath = ""
    destFilepath = ""
    essentialColor = QColor(255, 230, 153)
    map = dict()
    isChanged = None
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setStyleSheet("background:white;")
        self.sourceWidget = QWidget()
        self.destWidget = QWidget()
        self.sourceSheet = Sheet()
        self.destSheet = Sheet()


        self.gridLayout = QGridLayout()

        self.horizontalSplitter = QSplitter(Qt.Horizontal, self)
        self.horizontalSplitter.addWidget(self.sourceWidget)
        self.horizontalSplitter.addWidget(self.destWidget)
        self.horizontalSplitter.setStyleSheet("height:100%")

        self.gridLayout .addWidget(self.horizontalSplitter)
        self.setLayout(self.gridLayout)

        #self.vbox.setDirection(0)

        self.sbox = QVBoxLayout()
        self.sbox.addWidget(self.sourceSheet)
        self.sourceWidget.setLayout(self.sbox)

        self.dbox = QVBoxLayout()
        self.dbox.addWidget(self.destSheet)
        self.destWidget.setLayout(self.dbox)

        # self.sourceWidget.setStyleSheet("background-color: rgb(85, 170, 255);")
        # self.destWidget.setStyleSheet("background-color: rgb(255, 135, 137);")



    def fillLeft(self, filepath, sheetIndex):
        self.filepath = filepath
        self.sourceSheet.fillSheetByExcelSheetIndex(filepath, sheetIndex)

    def fillRight(self, filepath, sheetIndex):
        self.destSheet.fillSheetByExcelSheetIndex(filepath, sheetIndex)

    def beginTranslate(self, filePath, xmlFilePath, templateFilePath):
        self.fillRight(templateFilePath, 0)
        cfgItems = XETree.parse(xmlFilePath).getroot()
        for i in range(len(cfgItems)):
            cfgItem = cfgItems[i]
            istable = cfgItem.attrib['istable'].lower()
            if istable == 'true':
                self.__extractForTable(cfgItem)
            else:
                self.__extractForKeyValue(cfgItem)

    def __extractForKeyValue(self, cfgItem):
        sNode = cfgItem.find('source')
        dNode = cfgItem.find('dest')
        dCols = dNode.attrib['cols']
        anchors = sNode.attrib['anchor'].strip().split(';')
        length = len(anchors)
        beginCol = 0
        beginRow = 0
        sCols = sNode.attrib['cols']
        isFind = 0;
        datatype = dNode.attrib['datatype']
        tempcell = None
        for i in range(length):
            tempcell = self.__findStr(self.sourceSheet, anchors[i], beginRow, beginCol)
            if tempcell is not None:
                beginRow = tempcell.row() + 1
                beginCol = tempcell.column() + 1
                if i == length - 1:
                    isFind = 1
        if isFind == 1:
            sourceItem = self.sourceSheet.item(tempcell.row(), letterToint(sCols))
            self.map[parseChar(dCols)] = "{0},{1}".format(sCols, tempcell.row() + 1)
            item = self.destSheet.item(int(re.sub(r"\D", "", dCols)) - 1, letterToint(re.sub(r"[^A-Z]", "", dCols)))
            item.dataType = datatype
            item.setFormatValue(sourceItem.text())
            item.setBackground(self.essentialColor)
            sourceItem.setBackground(QColor(85, 170, 255))
        if isFind ==0:
            self.map[parseChar(dCols)] = ""
            item = self.destSheet.item(int(re.sub(r"\D", "", dCols)) - 1, letterToint(re.sub(r"[^A-Z]", "", dCols)))
            item.dataType = datatype
            item.setFormatValue("")
            item.setBackground(self.essentialColor)

    def __extractForTable(self, cfgItem):
        endrow = 0
        beginrow = 0
        tempcell = self.__findStr(self.sourceSheet, cfgItem[0].attrib["anchor"], 0, 0)

        self.sourceSheet.rowsMark.append(tempcell.row() - 1)
        destBeginRow = int(cfgItem[1].attrib["beginrow"])
        sList = cfgItem[0].attrib["cols"].split(',')
        dList = cfgItem[1].attrib["cols"].split(',')
        datatype = cfgItem[1].attrib["datatype"].split(',')

        for col in range(len(dList)):
            for row in range(int(cfgItem[1].attrib["limited"])):
                item = self.destSheet.item(destBeginRow + row -1, letterToint(dList[col]))
                if item is not None:
                    item.setBackground(self.essentialColor)
                    item.dataType = datatype[col]
        #如果差找不到，首行设置为NA
        if tempcell is None:
            for col in range(len(dList)):
                self.map[col + str(destBeginRow)] = "NA"
                item = self.destSheet.item(destBeginRow - 1, letterToint(dList[col]))
                #item.dataType = datatype[col]
                item.setText("NA")
            return
        # 判断在源文件中的起始行
        beginrow = tempcell.row() + int(cfgItem[0].attrib["skiprows"]) + 1

        # 判断结束行
        # 如果范围已经确定，直接确定结束行
        if cfgItem[0].attrib["range"] != "":
            endrow = beginrow + int(cfgItem[0].attrib["range"]) - 1
        # 范围不确定，根据下一行字符确定结束行
        elif cfgItem[0].attrib["anchorend"] != "":
            endrow = self.__findStr(self.sourceSheet, cfgItem[0].attrib["anchorend"], beginrow, 0).row
        else:
            # 找不到字符，则直到最后一个不为空行的为止
            limited = 1
            while self.sourceSheet.item(beginrow + limited, 0).text() != "" and self.sourceSheet.item(beginrow + limited, 0).text() is not None:
                limited += 1
            endrow = beginrow + limited - 1

        #print(cfgItem[0].attrib["anchor"])
        # 粘贴数据
        for row in range(beginrow, endrow + 1):
            for col in range(0, len(sList)):
                sourceItem = self.sourceSheet.item(row, letterToint(sList[col]))
                self.map[dList[col] + str(destBeginRow + row - beginrow)] = sList[col] + str(row + 1)
                item = self.destSheet.item(destBeginRow + row - beginrow - 1, letterToint(dList[col]))
                #item.dataType = datatype[col]
                item.setFormatValue(sourceItem.text())
                sourceItem.setBackground(QColor(85, 170, 255))


    def __findStr(self, sheet, key, startRow, startCol):
        for col in range(startCol, sheet.columnCount()):
            for row in range(startRow, sheet.rowCount()):
                cell = sheet.item(row, col)
                if cell is None:
                    print(col, row)
                if cell.text() is not None and cell.text() != "":
                    if key in str(cell.text()).replace(' ', ''):
                        return cell

    def save(self, filepath: str, sheetIndex: str):
        if self.isChanged is not None:
            self.destFilepath = filepath
            workbok = openpyxl.load_workbook(filepath)
            sheetNames = workbok.sheetnames
            sheet = workbok[sheetNames[sheetIndex]]
            for col in range(0, self.destSheet.columnCount()):
                for row in range(0, self.destSheet.rowCount()):
                    cell = self.destSheet.item(row, col)
                    if cell is None:
                        continue
                    if cell.text() is not None and cell.text() != "":
                        try:
                            value = cell.text()
                            if cell.dataType == "F":
                                value = float(value)
                            sheet[intToletter(col) + str(row + 1)].value = value
                        except Exception as e:
                            print(intToletter(col) + str(row))
            workbok.save(filepath)






class Sheet(QTableWidget):
    rowsMark = [int]
    def __init__(self):
        super().__init__()
        self.rowsMark.append(0)
        self.setColumnCount(10)
        self.setRowCount(30)
        for i in range(0, 10):
            item = TableItem(value=intToletter(i))
            self.setHorizontalHeaderItem(i, item)
        for i in range(0, 30):
            item = TableItem(value=str(i+1))
            self.setVerticalHeaderItem(i, item)

        self.horizontalScrollBar().setStyleSheet("height:25px;background-color: rgb(222, 222, 222);")
        self.verticalScrollBar().setStyleSheet("width:25px;background-color: rgb(222, 222, 222);")


    def fillSheetByExcelSheetIndex(self, filepath, sheetIndex):
        try:
            workbok = openpyxl.load_workbook(filepath)
            sheetNames = workbok.sheetnames
            self.__fileSheetByWorkBook(workbok[sheetNames[sheetIndex]])
        except Exception as e:
            QMessageBox.warning(self, "错误", str(e), QMessageBox.Ok)

    def fillSheetByExcelSheetName(self, filepath, sheetName):
        try:
            workbok = openpyxl.load_workbook(filepath)
            self.__fileSheetByWorkBook(workbok[sheetName])
        except Exception as e:
            QMessageBox.warning(self, "错误", str(e), QMessageBox.Ok)


    def __fileSheetByWorkBook(self, sheet):
        row_count = sheet.max_row
        col_count = sheet.max_column
        if self.rowCount() < row_count:
            self.setRowCount(row_count)
        if self.columnCount() < sheet.max_column:
            self.setColumnCount(col_count)
        for i in range(0, col_count):
            if self.horizontalHeaderItem(i) is None:
                item = TableItem(value=intToletter(i))
                self.setHorizontalHeaderItem(i, item)
        for i in range(0, row_count):
            if self.verticalHeaderItem(i) is None:
                item = TableItem(value=str(i))
                self.setVerticalHeaderItem(i, item)

        #开始初始化表格
        for row in sheet.iter_rows(min_row=0):
            for cell in row:
                item = TableItem("")
                self.setItem(cell.row - 1, cell.col_idx - 1, item)
                # color = cell.fill.bgColor.rgb
                # if color is not None and color != "00000000" and type(color) != openpyxl.styles.colors.RGB:
                #     color = HexToRgb(color)
                #     item.setBackground(QColor(color["r"], color["g"], color["b"]))
                if cell.value is not None and cell.value != "":
                    item.setText(str(cell.value))

        #合并单元格
        merged_cell = str(sheet.merged_cells).split(" ")
        for cell in merged_cell:
            startPoint = cell.split(":")[0]
            endPoint = cell.split(":")[1]
            startRow = int(re.sub(r"\D", "", startPoint)) - 1
            startCol = letterToint(re.sub(r"[^A-Z]", "", startPoint))
            endRow = int(re.sub(r"\D", "", endPoint)) - 1
            endCol = letterToint(re.sub(r"[^A-Z]", "", endPoint))
            self.setSpan(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1)



class TableItem(QTableWidgetItem):
    dataType = "S"
    IsSaved = True
    errorColor = QBrush(QColor(255, 114, 116))
    noramlColor = QBrush(QColor(255, 255, 255))
    essentialColor = QColor(255, 230, 153)
    def __int__(self):
        super().__init__()

    def __init__(self, value):
        super().__init__()
        self.setText(value)
    # def __init__(self, dataType, value):
    #     super().__init__()
    #     self.setFormatValue(value)
    def setFormatValue(self,value):
        value = str(value)
        try:
            value = self.updateValue(value)
            if not self.checkData(value):
                self.setBackground(self.errorColor)
            else:
                self.setBackground(self.essentialColor)
            super().setText(value)

        except Exception as e:
            print(e)

    def updateValue(self,value):  # 根据类型进行字符串处理，D日期，P百分数，S字符串不做处理，F或不填按数字处理
        value = str(value)
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
                    value = "0"
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

def parseChar(string):
    letter = re.sub(r"[^A-Z]", "", string)
    num = re.sub(r"\D", "", string)
    return "{0},{1}".format(letter, num)
#十六进制颜色编码转rgb
def HexToRgb(tmp):
    rgb = dict()
    opt = re.findall(r'(.{2})',tmp)
    rgb["r"] = 255 - int(opt[0], 16)
    rgb["g"] = 255 - int(opt[1], 16)
    rgb["b"] = 255 - int(opt[2], 16)
    return rgb
