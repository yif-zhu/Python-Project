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

class SettingManager(QDialog):
    submitted = pyqtSignal(str, str, dict)
    def __init__(self, parent, filetype, title):
        super().__init__(parent)
        self.filetype = filetype
        self.title = title
        self.str_templatefile = "templateFile"
        self.str_xmlFIle = "xmlFIle"
        self.str_OutputFolder = "OutputFolder"
        self.configFileDirectory = "./Config"
        self.configFileName = "config.xml"
        self.xmlFolderPath = "Resource/xml"
        self.templateFoldetPath = "Resource/templateFile"

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
        self.mainTable = Sheet()
        self.mainTable.setSelectionMode(QAbstractItemView.SingleSelection)
        # 初始化表头
        self.mainTable.setColumnCount(4)
        Item = QTableWidgetItem()
        Item.setText("Id")
        self.mainTable.setHorizontalHeaderItem(0, Item)
        Item = QTableWidgetItem()
        Item.setText("名称")
        self.mainTable.setHorizontalHeaderItem(1, Item)
        Item = QTableWidgetItem()
        Item.setText("操作")
        self.mainTable.setHorizontalHeaderItem(2, Item)

        Item = QTableWidgetItem()
        Item.setText("删除")
        self.mainTable.setHorizontalHeaderItem(3, Item)

        self.mainTable.verticalHeader().hide()
        self.mainTable.setColumnWidth(0, 91)
        self.mainTable.setColumnWidth(1, 600)
        self.mainTable.setColumnWidth(2, 91)
        self.mainTable.setColumnWidth(3, 91)
        self.initTableData()
        self.bodyLayout.addWidget(self.mainTable)
    def openFile(self):
        row = self.mainTable.selectedIndexes()[0].row()
        file = self.mainTable.item(row, 1).toolTip()
        if not os.path.lexists(file):
            QMessageBox.critical(self, "错误", "文件不存在", QMessageBox.Ok)
            return
        os.startfile(os.getcwd()+"/"+ file)

    def initTableData(self):
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find(self.filetype)
        items = node.getchildren()
        for i in range(0, len(items)):
            item = TableItem(items[i].attrib["Name"])
            item.setToolTip(items[i].text)
            self.mainTable.setItem(i, 1, item)
            item = TableItem(str(i))
            self.mainTable.setItem(i, 0, item)
            item = QPushButton()
            item.setText("查看")
            self.mainTable.setCellWidget(i, 2, item)
            item.clicked.connect(lambda : self.openFile())

            item = QPushButton()
            item.setText("删除")
            self.mainTable.setCellWidget(i, 3, item)
            item.clicked.connect(self.delete)


        # 新增按钮
        self.button_add = QPushButton()
        self.button_add.setText("新增")
        self.button_add.setObjectName("btn_add")
        self.button_add.clicked.connect(self.add)
        self.mainTable.setCellWidget(self.mainTable.rowCount(), 0, self.button_add)
    def delete(self):
        row = self.mainTable.selectedIndexes()[0].row()
        self.mainTable.removeRow(row)


    def add(self):
        count = self.mainTable.rowCount() -1

        if self.filetype == self.str_xmlFIle:
            filter = "*.xml"
        elif self.filetype == self.str_templatefile:
            filter = "*.xlsx"
        button_file = QPushButton()
        button_file.setText("选择文件")
        item = TableItem("")
        self.mainTable.setItem(count, 1, item)

        id = TableItem(str(count + 1))

        self.mainTable.setItem(count, 0, id)
        self.mainTable.setCellWidget(count, 2, button_file)
        button_file.clicked.connect(lambda : self.open(filter))

        item = QPushButton()
        item.setText("删除")
        self.mainTable.setCellWidget(count, 3, item)
        item.clicked.connect(self.delete)


        self.mainTable.setCellWidget(count + 1, 0, self.button_add)


    def open(self, filter) -> None:
        file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择文件', filter=filter)
        if file_Exolorer[0]:
            for i in range(self.mainTable.rowCount() -1):
                filepath = self.mainTable.item(i, 1).toolTip()
                if os.path.basename(filepath) == os.path.basename(file_Exolorer[0]):
                    QMessageBox.warning(self, "警告", "已添加同名文件,id 为" + str(i), QMessageBox.Ok)
                    return
            destPath = ""
            if filter == "*.xlsx":
                destPath = self.templateFoldetPath + "/" + os.path.basename(file_Exolorer[0])
            elif filter == "*.xml":
                destPath = self.xmlFolderPath + "/" + os.path.basename(file_Exolorer[0])
            byte = open(file_Exolorer[0], "rb").read()
            open(destPath, "wb").write(byte)
            row = self.mainTable.selectedIndexes()[0].row()
            item = self.mainTable.item(row, 1)
            item.setText(os.path.basename(file_Exolorer[0]))
            item.setToolTip(destPath)

    def submit(self):
        tree = XETree.parse(self.configFilePath)
        root = tree.getroot()
        node = root.find(self.filetype)
        node.clear()
        for i in range(self.mainTable.rowCount() - 1):
            filepath = self.mainTable.item(i, 1).toolTip()
            name = self.mainTable.item(i, 1).text()
            if filepath != "":
                element = XETree.Element("File")
                element.set("Name", name)
                element.text = filepath
                node.append(element)
        indent(node)
        tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)

        sip.delete(self.mainTable)
        self.initTable()
        message = QMessageBox(self)
        message.setWindowTitle("完成")
        message.setWindowIcon(QIcon("Resource/icon/Icon_table.ico"))
        message.setText("保存成功")
        message.show()






