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
from Source.TabWidget import *
import pyodbc
import copy

class PoolStaticDataInsert(QDialog):
    submitted = pyqtSignal(str, str, dict)
    def __init__(self, parent):
        super().__init__(parent)
        self.title = '静态池PDF文件导入'
        self.str_AssetType = "AssetType"
        self.configFileDirectory = "./Config"
        self.filetype = 'PoolStaticInsertSql'
        self.FileName = "Config.xml"
        self.dbConnectionStr = 'DRIVER={SQL Server};SERVER=172.16.6.143\MSSQL;DATABASE=PortfolioManagement;UID=sa;PWD=PasswordGS2017'
        self.sql = 'exec PortfolioManagement.DVImport.usp_getPoolStaticColumnMapping '
        self.configFilePath = os.path.join(self.configFileDirectory, self.FileName)
        self.Mapping = self.execSQLCmdFetchAll()
        self.MyComboKey = []  #字段名下拉框
        self.MyComboDesc = []  #字段解释
        self.MyComboValue =[]     #字段名解释字段下拉框
        self.InitCombo()
        self.__initUI()

    def InitCombo(self):
        for i in range(len(self.Mapping)):
            self.MyComboKey.append(self.Mapping[i][0])
            self.MyComboDesc.append(self.Mapping[i][1])
            self.MyComboValue.append(self.Mapping[i][0]+' ('+self.Mapping[i][1]+')')

    def __initUI(self):
        self.setObjectName("Main")
        stylesheet = open("Source/SettingManager.qss", "r").read()
        self.setWindowTitle(self.title)
        self.setWindowModality(Qt.ApplicationModal)
        self.setStyleSheet(stylesheet)

        self.body = QWidget()
        self.body.setStyleSheet("")
        self.body.setObjectName("body")

        self.bodyLayout = QGridLayout(self.body)
        self.bodyLayout.setObjectName("bodyLayout")
        #选择入库的PDF文件
        self.label_selectFile = QLabel(text="待入库文件：")
        self.label_selectFile.setAlignment(Qt.AlignCenter)
        self.label_selectFile.setObjectName("label_selectFile")

        self.lineEdit_selectFile = QLineEdit()
        self.lineEdit_selectFile.setObjectName("lineEdit_selectFile")
        self.lineEdit_selectFile.setReadOnly(True)

        self.button_selectFIle = QPushButton()
        self.button_selectFIle.setText("..")
        self.button_selectFIle.clicked.connect(self.open)
        # 选择入库的PDF文件开始读取页数
        self.label_BeginPage = QLabel(text="开始页数：")
        self.label_BeginPage.setAlignment(Qt.AlignCenter)
        self.label_BeginPage.setObjectName("label_BeginPage")

        self.lineEdit_BeginPage = QLineEdit()
        self.lineEdit_BeginPage.setObjectName("lineEdit_BeginPage")
        #self.lineEdit_selectFile.setReadOnly(True)
        # 选择入库的PDF文件截止读取页数
        self.label_EndPage = QLabel(text="截止页数：")
        self.label_EndPage.setAlignment(Qt.AlignCenter)
        self.label_EndPage.setObjectName("label_EndPage")

        self.lineEdit_EndPage = QLineEdit()
        self.lineEdit_EndPage.setObjectName("lineEdit_EndPage")
        #self.lineEdit_selectFile.setReadOnly(True)

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setSizeConstraint(QLayout.SetDefaultConstraint)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")

        self.setLayout(self.verticalLayout)
        self.bodyLayout.addWidget(self.label_selectFile, 0, 0)
        self.bodyLayout.addWidget(self.lineEdit_selectFile, 0, 1)
        self.bodyLayout.addWidget(self.button_selectFIle, 0, 2)
        self.bodyLayout.addWidget(self.label_BeginPage, 1, 0)
        self.bodyLayout.addWidget(self.lineEdit_BeginPage, 1, 1)
        self.bodyLayout.addWidget(self.label_EndPage, 2, 0)
        self.bodyLayout.addWidget(self.lineEdit_EndPage, 2, 1)

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
        self.button_submit.setText("运行")
        self.button_submit.clicked.connect(self.submit)
        self.bottom_layount.addWidget(self.button_submit)
    def open(self) -> None:
        file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择文件', filter='*.pdf')
        self.lineEdit_selectFile.setText(file_Exolorer[0])
        return

    def execSQLCmdFetchAll(self):
        # print(sql)
        cnxn = pyodbc.connect(self.dbConnectionStr)
        try:
            cursor = cnxn.cursor()
            rows = cursor.execute(self.sql).fetchall()
            return rows
        except Exception as ex:
            raise ex
        finally:
            cnxn.close()
    def initTable(self):
        self.mainTable = Sheet()
        self.mainTable.setSelectionMode(QAbstractItemView.SingleSelection)
        # 初始化表头
        self.mainTable.setColumnCount(3)
        Item = QTableWidgetItem()
        Item.setText("Id")
        self.mainTable.setHorizontalHeaderItem(0, Item)
        Item = QTableWidgetItem()
        Item.setText("数据库字段")
        self.mainTable.setHorizontalHeaderItem(1, Item)
        # Item = QTableWidgetItem()
        # Item.setText("Excle字段名")
        # self.mainTable.setHorizontalHeaderItem(2, Item)

        Item = QTableWidgetItem()
        Item.setText("操作")
        self.mainTable.setHorizontalHeaderItem(2, Item)

        self.mainTable.verticalHeader().hide()
        self.mainTable.setColumnWidth(0, 100)
        self.mainTable.setColumnWidth(1, 650)
        self.mainTable.setColumnWidth(2, 100)
        # self.mainTable.setColumnWidth(3, 100)
        self.initTableData()
        self.bodyLayout.addWidget(self.mainTable, 3, 0, 1, 3)

    def initTableData(self):
        ComboValue = QComboBox()
        ComboValue.addItems(self.MyComboValue)
        item = TableItem(str(1))
        desc = TableItem('')
        self.mainTable.setItem(0, 0, item)
        self.mainTable.setCellWidget(0, 1, ComboValue)
        # self.mainTable.setItem(0, 2, desc)


        item = QPushButton()
        item.setText("删除")
        self.mainTable.setCellWidget(0, 2, item)
        item.clicked.connect(self.delete)

        # 新增按钮
        self.button_add = QPushButton()
        self.button_add.setText("新增")
        self.button_add.setObjectName("btn_add")
        self.button_add.clicked.connect(self.add)
        self.mainTable.setCellWidget(self.mainTable.rowCount(), 0, self.button_add)

    # def IndexChange(self, index, row):
    #     desc = self.MyComboValue[index]
    #     self.mainTable.setItem(row, 2, desc)

    def delete(self):
        row = self.mainTable.selectedIndexes()[0].row()
        self.mainTable.removeRow(row)

    def add(self):
        count = self.mainTable.rowCount() -1

        Key = QComboBox()
        Key.addItems(self.MyComboKey)
        Value = QComboBox()
        Value.addItems(self.MyComboValue)
        self.mainTable.setCellWidget(count, 1, Value)
        # item = TableItem("")
        # self.mainTable.setItem(count, 2, item)

        id = TableItem(str(count + 1))

        self.mainTable.setItem(count, 0, id)

        item = QPushButton()
        item.setText("删除")
        self.mainTable.setCellWidget(count, 2, item)
        item.clicked.connect(self.delete)
        self.mainTable.setCellWidget(count + 1, 0, self.button_add)

    def submit(self):
        tree = XETree.parse(self.configFilePath)
        root = tree.getroot()
        node = root.find(self.filetype)
        node.clear()
        insertSql = 'insert into PortfolioManagement.DvImport.StaticPoolData(FileNames'
        columns = ''
        ExcelColumns = ''
        for i in range(self.mainTable.rowCount() - 1):
            # name = self.mainTable.item(i, 2).text()
            index = self.mainTable.cellWidget(i, 1).currentIndex()
            if index != "":
                insertSql = insertSql + ',' + self.MyComboKey[index]
                columns = columns + self.MyComboDesc[index] + ','
                # ExcelColumns = ExcelColumns + name + ','
        insertSql = insertSql +') Values '
        element = XETree.Element("Sql")
        element.set("Text", insertSql)
        node.append(element)
        indent(node)
        tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)

        FilePath = self.lineEdit_selectFile.text()
        BeginPage  = self.lineEdit_BeginPage.text()
        EndPage = self.lineEdit_EndPage.text()
        parameters = dict()
        parameters["sql"] = str(insertSql)
        parameters["filePath"] = str(FilePath)
        parameters["beginPage"] = str(BeginPage)
        parameters["endPage"] = str(EndPage)
        parameters["columns"] = columns.rstrip(',')
        # parameters["ExcelColumns"] = ExcelColumns.rstrip(',')
        parameters["outputFile"] = ''
        print(columns)
        sip.delete(self.mainTable)
        self.hide()
        self.submitted.emit("PoolStaticDataInsert", "静态池PDF导入", parameters)
        self.close()






