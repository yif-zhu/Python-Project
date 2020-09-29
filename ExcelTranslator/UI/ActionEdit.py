# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Dialog.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import shutil
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
        self.mainTable.setColumnCount(6)
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

        Item = QTableWidgetItem()
        Item.setText("删除")
        self.mainTable.setHorizontalHeaderItem(5, Item)
        self.mainTable.verticalHeader().hide()
        self.mainTable.setColumnWidth(0, 91)
        self.mainTable.setColumnWidth(1, 150)
        self.mainTable.setColumnWidth(2, 150)
        self.mainTable.setColumnWidth(3, 308)
        self.mainTable.setColumnWidth(4, 91)
        self.mainTable.setColumnWidth(5, 91)
        self.initTableData()
        self.bodyLayout.addWidget(self.mainTable)

    def initTableData(self):
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find("Action[@ActionCode='{0}']".format(self.acitonCode))
        #如果不存在该节点，即新增

        if node is None:
            self.addPythonFileRow("", "")
            self.addXmlFileRow("", "", "False")
            self.button_add = QPushButton()
            self.button_add.setText("新增")
            self.button_add.setObjectName("btn_add")
            self.button_add.clicked.connect(self.add)
            self.mainTable.setCellWidget(self.mainTable.rowCount(), 0, self.button_add)
            return
        item = node.find("Variable[@VariableName='PythonFile']")
        if item is not None:
            self.addPythonFileRow(item.attrib["ParameterName"], item.attrib["Value"])
        else:
            self.addPythonFileRow("", "")
        item = node.find("Variable[@VariableName='XmlFile']")
        if item is not None:
            self.addXmlFileRow(item.attrib["ParameterName"], item.attrib["Value"], item.attrib["IsParameter"])
        else:
            self.addXmlFileRow("", "", "False")
        item = node.find("Variable[@VariableName='MethodName']")
        if item is not None:
            self.addMethodName(item.attrib["ParameterName"], item.attrib["Value"], item.attrib["IsParameter"])
        else:
            self.addMethodName("", "", "False")

        items = node.findall("Variable")
        for i in range(0, len(items)):
            count = self.mainTable.rowCount()
            varName = items[i].attrib["VariableName"]
            if varName == "XmlFile" or varName == "PythonFile" or varName == "MethodName":
                continue
            item = TableItem(varName)
            item.setFlags(Qt.ItemIsSelectable|Qt.ItemIsDragEnabled|Qt.ItemIsUserCheckable|Qt.ItemIsEnabled)
            #item.setToolTip(items[i].text)
            self.mainTable.setItem(count, 1, item)
            item = TableItem(items[i].attrib["ParameterName"])
            item.setFlags(Qt.ItemIsSelectable|Qt.ItemIsDragEnabled|Qt.ItemIsUserCheckable|Qt.ItemIsEnabled)
            self.mainTable.setItem(count, 2, item)
            item = TableItem(items[i].attrib["Value"])
            self.mainTable.setItem(count, 3, item)

            isParameter = items[i].attrib["IsParameter"]
            item = TableItem("")
            if isParameter == "True":
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
            self.mainTable.setItem(count, 4, item)
            item = QPushButton()
            item.setText("编辑")
            self.mainTable.setCellWidget(count, 0, item)
            item.clicked.connect(self.seteditable)

            item = QPushButton()
            item.setText("删除")
            self.mainTable.setCellWidget(count, 5, item)
            item.clicked.connect(self.delete)

        # 新增按钮
        self.button_add = QPushButton()
        self.button_add.setText("新增")
        self.button_add.setObjectName("btn_add")
        self.button_add.clicked.connect(self.add)
        self.mainTable.setCellWidget(self.mainTable.rowCount(), 0, self.button_add)

    def addPythonFileRow(self, paraName, value):
        item = TableItem("PythonFile")
        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        # item.setToolTip(items[i].text)
        self.mainTable.setItem(0, 1, item)
        item = TableItem(paraName)
        self.mainTable.setItem(0, 2, item)
        item = TableItem(value)
        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        self.mainTable.setItem(0, 3, item)

        isParameter = "False"
        item = TableItem("")
        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsEnabled)
        if isParameter == "True":
            item.setCheckState(Qt.Checked)
        else:
            item.setCheckState(Qt.Unchecked)
        self.mainTable.setItem(0, 4, item)
        item = QPushButton()
        item.setText("选择文件")
        self.mainTable.setCellWidget(0, 0, item)
        item.clicked.connect(self.openPyFile)

    def addXmlFileRow(self,paraName, value,isParameter):
        item = TableItem("XmlFile")
        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        # item.setToolTip(items[i].text)
        self.mainTable.setItem(1, 1, item)
        item = TableItem(paraName)
        self.mainTable.setItem(1, 2, item)
        item = TableItem(value)
        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        self.mainTable.setItem(1, 3, item)

        item = TableItem("")
        if isParameter == "True":
            item.setCheckState(Qt.Checked)
        else:
            item.setCheckState(Qt.Unchecked)
        self.mainTable.setItem(1, 4, item)
        item = QPushButton()
        item.setText("选择文件")
        self.mainTable.setCellWidget(1, 0, item)
        item.clicked.connect(self.openXmlFile)

    def addMethodName(self, paraName, value, isParameter):
        item = TableItem("MethodName")
        self.mainTable.setItem(2, 1, item)
        item = TableItem(paraName)
        self.mainTable.setItem(2, 2, item)
        item = TableItem(value)
        self.mainTable.setItem(2, 3, item)
        item = TableItem("")
        if isParameter == "True":
            item.setCheckState(Qt.Checked)
        else:
            item.setCheckState(Qt.Unchecked)
        self.mainTable.setItem(2, 4, item)

    def openPyFile(self) -> None:
        root = XETree.parse(self.configFilePath)
        file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择文件', filter="*.py")
        filepath = ""
        if file_Exolorer[0]:
            filepath = file_Exolorer[0]
            fileName = os.path.basename(filepath)
            destPath = "FileTranslator" + "/" + fileName
            if root.find("Variable[@Value='{0}']".format(destPath)) is not None:
                QMessageBox.warning(self, "警告", "已添加同名文件", QMessageBox.Ok)
                return
            byte = open(filepath, "rb").read()
            open(destPath, "wb").write(byte)
            self.mainTable.item(0, 3).setText(fileName)

    def openXmlFile(self) -> None:
        root = XETree.parse(self.configFilePath)
        file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择文件', filter="*.xml")
        filepath = ""
        if file_Exolorer[0]:
            filepath = file_Exolorer[0]
            fileName = os.path.basename(filepath)
            destPath = "FileTranslator" + "/" + fileName
            if root.find("Variable[@Value='{0}']".format(destPath)) is not None:
                QMessageBox.warning(self, "警告", "已添加同名文件", QMessageBox.Ok)
                return
            byte = open(filepath, "rb").read()
            open(destPath, "wb").write(byte)
            item = self.mainTable.item(1, 3)
            item.setText(destPath)

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

        item = QPushButton()
        item.setText("删除")
        self.mainTable.setCellWidget(count, 5, item)
        item.clicked.connect(self.delete)

        self.mainTable.setCellWidget(count + 1, 0, self.button_add)

    def seteditable(self):
        row = self.mainTable.selectedIndexes()[0].row()
        self.mainTable.item(row, 2).setFlags(Qt.ItemIsSelectable|Qt.ItemIsDragEnabled|Qt.ItemIsUserCheckable|Qt.ItemIsEnabled|Qt.ItemIsEditable)
        self.mainTable.item(row, 1).setFlags(Qt.ItemIsSelectable|Qt.ItemIsDragEnabled|Qt.ItemIsUserCheckable|Qt.ItemIsEnabled|Qt.ItemIsEditable)
    def delete(self):
        row = self.mainTable.selectedIndexes()[0].row()
        self.mainTable.removeRow(row)

    def submit(self):
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find("Action[@ActionCode='{0}']".format(self.acitonCode))
        if node is None:
            node = XETree.Element("Action")
            node.set("ActionCode", self.acitonCode)
            tree.getroot().append(node)
        for item in node.getchildren():
            node.remove(item)

        for row in range(self.mainTable.rowCount() -1):
            VarName = self.mainTable.item(row, 1).text()
            parName = self.mainTable.item(row, 2).text()
            value = self.mainTable.item(row, 3).text()
            isPara = "True" if self.mainTable.item(row, 4).checkState() == 2 else "False"
            if VarName == "":
                continue

            ele_Variable = XETree.Element("Variable")
            ele_Variable.set("VariableName", VarName)
            ele_Variable.set("ParameterName", parName)
            ele_Variable.set("Value", value)
            ele_Variable.set("IsParameter", isPara)
            node.append(ele_Variable)
        indent(node)
        tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)

        sip.delete(self.mainTable)
        self.initTable()
        message = QMessageBox()
        message.setWindowTitle("完成")
        message.setWindowIcon(QIcon("Resource/icon/Icon_table.ico"))
        message.setText("保存成功")
        message.exec()
        self.close()







