
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sip
import os
import xml.etree.ElementTree as XETree
from UI.ActionEdit import ActionEdit
from UI.TabWidget import *
from Action.Action import Action

class AcitonManager(QDialog):
    actionShowed = pyqtSignal(str, str, dict)
    newId = dict()
    maxId = 0
    def __init__(self, parent, filetype, title):
        super().__init__(parent)
        self.title = title
        self.configFileDirectory = "./Config"
        self.configFileName = "Actions.xml"
        self.configFilePath = os.path.join(self.configFileDirectory, self.configFileName)
        self.removedAction = []
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
        self.button_submit.setStyleSheet("margin-left: 350px;")
        self.bottom_layount.addWidget(self.button_submit)

    def initTable(self):
        self.count = 0
        self.mainTable = Sheet()
        self.mainTable.setSelectionMode(QAbstractItemView.SingleSelection)
        # 初始化表头
        self.mainTable.setColumnCount(7)
        Item = QTableWidgetItem()
        Item.setText("Id")
        self.mainTable.setHorizontalHeaderItem(0, Item)
        Item = QTableWidgetItem()
        Item.setText("任务名称")
        self.mainTable.setHorizontalHeaderItem(1, Item)
        Item = QTableWidgetItem()
        Item.setText("任务编码")
        self.mainTable.setHorizontalHeaderItem(2, Item)
        Item = QTableWidgetItem()
        Item.setText("操作")
        self.mainTable.setHorizontalHeaderItem(3, Item)
        Item = QTableWidgetItem()
        Item.setText("运行并显示")
        self.mainTable.setHorizontalHeaderItem(4, Item)

        Item = QTableWidgetItem()
        Item.setText("仅运行")
        self.mainTable.setHorizontalHeaderItem(5, Item)

        Item = QTableWidgetItem()
        Item.setText("删除")
        self.mainTable.setHorizontalHeaderItem(6, Item)

        self.mainTable.verticalHeader().hide()
        self.mainTable.setColumnWidth(0, 91)
        self.mainTable.setColumnWidth(1, 150)
        self.mainTable.setColumnWidth(2, 150)
        self.mainTable.setColumnWidth(3, 91)
        self.mainTable.setColumnWidth(4, 91)
        self.mainTable.setColumnWidth(5, 91)
        self.mainTable.setColumnWidth(6, 91)
        self.initTableData()
        self.bodyLayout.addWidget(self.mainTable)

    def initTableData(self):
        tree = XETree.parse(self.configFilePath)
        items = tree.getroot().findall("Action")
        for i in range(0, len(items)):
            id = int(items[i].attrib["Id"])
            if id > self.maxId:
                self.maxId = id
            item = TableItem(items[i].attrib["Id"])
            #item.setToolTip(items[i].text)
            self.mainTable.setItem(i, 0, item)
            item = TableItem(items[i].attrib["AcitonName"])
            self.mainTable.setItem(i, 1, item)
            item = TableItem(items[i].attrib["ActionCode"])
            self.mainTable.setItem(i, 2, item)
            item = QPushButton()
            item.setText("编辑参数")
            self.mainTable.setCellWidget(i, 3, item)
            item.clicked.connect(self.edit)

            item = QPushButton()
            item.setText("运行并查看")
            self.mainTable.setCellWidget(i, 4, item)
            item.clicked.connect(self.runAndShow)

            item = QPushButton()
            item.setText("仅运行")
            self.mainTable.setCellWidget(i, 5, item)
            item.clicked.connect(self.runOnly)

            item = QPushButton()
            item.setText("删除")
            self.mainTable.setCellWidget(i, 6, item)
            item.clicked.connect(self.delete)



        # 新增按钮
        self.button_add = QPushButton()
        self.button_add.setText("新增")
        self.button_add.setObjectName("btn_add")
        self.button_add.clicked.connect(self.add)
        self.mainTable.setCellWidget(self.mainTable.rowCount(), 0, self.button_add)

    def add(self):
        count = self.mainTable.rowCount() -1

        button_edit = QPushButton()
        button_edit.setText("编辑参数")
        item = TableItem("")
        self.mainTable.setItem(count, 1, item)
        item = TableItem("")
        self.mainTable.setItem(count, 2, item)

        self.maxId = self.maxId + 1
        id = TableItem(str(self.maxId))
        self.mainTable.setItem(count, 0, id)
        self.mainTable.setCellWidget(count, 3, button_edit)

        item = QPushButton()
        item.setText("删除")
        self.mainTable.setCellWidget(count, 6, item)
        item.clicked.connect(self.delete)

        button_edit.clicked.connect(self.edit)
        self.mainTable.setCellWidget(count + 1, 0, self.button_add)
        self.newId[str(count)] = XETree.Element("Variable")

    def edit(self):
        row = self.mainTable.selectedIndexes()[0].row()
        actioncode = self.mainTable.item(row, 2).text()
        if actioncode == "":
            QMessageBox.warning(self, "警告", "ActionCode不能为空", QMessageBox.Ok)
            return
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find("Action[@ActionCode='{0}']".format(actioncode))
        if node is None:
            QMessageBox.warning(self, "警告", "请先点击保存再进行编辑", QMessageBox.Ok)
            return
        actionEdit = ActionEdit(self, actioncode, "编辑参数")
        actionEdit.setFixedSize(905, 800)
        actionEdit.show()

    def runAndShow(self):
        self.hide()
        row = self.mainTable.selectedIndexes()[0].row()
        actioncode = self.mainTable.item(row, 2).text()
        actionName = self.mainTable.item(row, 1).text()
        self.actionShowed.emit(actioncode, actionName, {})
        self.show()

    def runOnly(self):
        row = self.mainTable.selectedIndexes()[0].row()
        actioncode = self.mainTable.item(row, 2).text()
        action = Action(actioncode, {})
        action.runAciton()
    def delete(self):
        row = self.mainTable.selectedIndexes()[0].row()
        actioncode = self.mainTable.item(row, 2).text()
        self.mainTable.removeRow(row)
        self.removedAction.append(actioncode)
        #self.newId.pop(str(row))

    def submit(self):
        tree = XETree.parse(self.configFilePath)
        root = tree.getroot()
        for code in self.removedAction:
            node = root.find("Action[@ActionCode='{0}']".format(code))
            if node is not None:
                root.remove(node)

        for (k, v) in self.newId.items():
            row = int(k)
            Id = self.mainTable.item(row, 0).text()
            actionName = self.mainTable.item(row, 1).text()
            actionCode = self.mainTable.item(row, 2).text()
            if actionCode == "":
                continue
            ele_Action = XETree.Element("Action")
            ele_Action.set("Id", Id)
            ele_Action.set("AcitonName", actionName)
            ele_Action.set("ActionCode", actionCode)
            #ele_Action.append(v)
            root.append(ele_Action)
        indent(root)
        tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)
        self.newId.clear()
        sip.delete(self.mainTable)
        self.initTable()
        message = QMessageBox()
        message.setWindowTitle("完成")
        message.setWindowIcon(QIcon("Resource/icon/Icon_table.ico"))
        message.setText("保存成功")
        message.exec()






