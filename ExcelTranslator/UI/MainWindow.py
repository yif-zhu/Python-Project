# text-editor.py

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import xml.etree.ElementTree as XETree
import shutil
import traceback
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
from UI.AssetTypeManager import AssetTypeManager

class MainWindow(QMainWindow):
    str_WindowName = "ExcelTranslator"
    str_templatefile = "templateFile"
    str_xmlFIle = "xmlFIle"
    str_OutputFolder = "OutputFolder"
    str_AssetType = "AssetType"
    configFileDirectory = "./Config"
    configFileName = "config.xml"
    configFilePath = os.path.join(configFileDirectory, configFileName)
    BaseCheck = ["基础资产校验", "基础资产校验_V1", "基础资产校验_V2"]
    Settings = ["配置管理", "方案管理", "模板管理", "资产类型管理", "任务管理"]
    clipboardFile = "Config/clipBoard.et"
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
        #self.textEdit.setStyleSheet("background:white;height:20%")
        self.textEdit.setReadOnly(True)


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

        # 基础校验
        # self.button_checkValue = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_run.ico'), '基础资产校验')
        # self.button_checkValue.setText("基础资产校验")
        # self.button_checkValue.setToolTip('基础资产校验')
        # self.button_checkValue.clicked.connect(self.valueCheck)

        self.comboBox_baseCheck = DropDownMenu()
        self.comboBox_baseCheck.setObjectName("BaseCheck");
        for i in range(len(self.BaseCheck)):
            self.comboBox_baseCheck.addItem(self.BaseCheck[i], "")
            self.comboBox_baseCheck.setItemIcon(i, QIcon('Resource/icon/Icon_run.ico'))
        self.comboBox_baseCheck.currentIndexChanged.connect(self.showBaseCheck)

        # 不良校验
        self.button_NoncheckValue = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_run.ico'), '不良资产校验')
        self.button_NoncheckValue.setText("不良资产校验")
        self.button_NoncheckValue.setToolTip('不良资产校验')
        self.button_NoncheckValue.clicked.connect(self.NonvalueCheck)

        # 池分布校验
        self.button_PoolcheckValue = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_run.ico'), '池分布校验')
        self.button_PoolcheckValue.setText("池分布校验")
        self.button_PoolcheckValue.setToolTip('池分布校验')
        self.button_PoolcheckValue.clicked.connect(self.PoolvalueCheck)

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
        saveAction.clicked.connect(lambda :self.save(self.workTab.currentIndex()))

        #保存全部
        saveAllAction =  QtWidgets.QPushButton()
        saveAllAction.setStatusTip('保存全部')
        saveAllAction.setText("保存全部")
        saveAllAction.clicked.connect(self.saveAll)

        #配置管理
        self.comboBox_setting = DropDownMenu()
        self.comboBox_setting.setObjectName("Setting")
        for i in range(len(self.Settings)):
            self.comboBox_setting.addItem(self.Settings[i], "")
            self.comboBox_setting.setItemIcon(i, QIcon("Resource/icon/Icon_setting.ico"))
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
        tb1.addWidget(self.button_multiRun)   #注释内容填充和批量填充
        # tb1.addWidget(self.comboBox_baseCheck)
        # tb1.addWidget(self.button_NoncheckValue)
        # tb1.addWidget(self.button_PoolcheckValue)
        tb1.addWidget(saveAction)
        # tb1.addWidget(self.button_saveAs)
        #tb1.addWidget(saveAllAction)
        tb1.addWidget(self.comboBox_setting)




        #self.verticalSplitter.setStyleSheet("background-color: rgb(222, 222, 222);")


        self.statusBar()


        #遮罩
        self.maskwidget = QWidget(self)
        self.maskwidget.setObjectName("Mask")
        self.maskwidget.setGeometry(0, 0, 1920, 1080)
        self.maskwidget.raise_()

        self.masklabel = QLabel(self.maskwidget)
        self.maskwidget.setGeometry(0, 0, 1920, 1080)
        self.masklabel.setText("loading...")

        #self.masklabel.move(self.maskwidget.rect().center())
        # self.loadingGif = QMovie('Resource/icon/loading-image.gif')
        # self.masklabel.setMovie(self.loadingGif)
        # self.loadingGif.start()
        self.maskwidget.hide()

        #show
        self.setObjectName("MainWindow")
        self.setGeometry(0, 0, 1280, 760)
        self.setWindowTitle("Excel转换")
        self.setWindowIcon(QIcon('Resource/icon/Icon_table.ico'))
        #self.center()
        self.show()
        self.showMaximized()

        #self.showMask()
        #fillDataDialog.setGeometry((self.maximumSize().width() - 500)/2,(self.maximumSize().height()-800)/2, 500, 800 )


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

    def showBaseCheck(self, index:int):
        if index == 1:
            self.valueCheck()
        elif index == 2:
            self.valueCheck2()
    def showManager(self, index:int):
        if index == 1:
            self.showModeManager()
        elif index == 2:
            self.showTemplateManager()
        elif index == 3:
            self.showAssetTypeManager()
        elif index == 4:
            self.showActionManager()
    def showModeManager(self):
        modeManager = SettingManager(self, self.str_xmlFIle, self.Settings[1])
        modeManager.raise_()
        center(modeManager, 897, 800)
        modeManager.setFixedSize(897, 800)
        modeManager.show()

    def showTemplateManager(self):
        templateManager = SettingManager(self, self.str_templatefile, self.Settings[2])
        templateManager.raise_()
        center(templateManager, 897, 800)
        templateManager.setFixedSize(897, 800)
        templateManager.show()

    def showAssetTypeManager(self):
        assetTypeManager = AssetTypeManager(self, self.str_AssetType, self.Settings[3])
        assetTypeManager.raise_()
        center(assetTypeManager, 897, 800)
        assetTypeManager.setFixedSize(897, 800)
        assetTypeManager.show()

    def showActionManager(self):
        actionManager = AcitonManager(self, self.str_templatefile, self.Settings[4])
        actionManager.raise_()
        actionManager.actionShowed.connect(self.newAction)
        center(actionManager, 779, 800)
        actionManager.setFixedSize(779, 800)
        actionManager.show()

        return
    # open
    def new(self):
        #新建任务对话框

        fillDataDialog = FillData(self, 0)
        fillDataDialog.raise_()
        fillDataDialog.submitted.connect(self.newAction)

        center(fillDataDialog, 500, 500)
        fillDataDialog.setFixedSize(500, 500)
        fillDataDialog.show()

    def multiRun(self):
        fillDataDialog = FillData(self, 1)
        fillDataDialog.raise_()
        fillDataDialog.submitted.connect(self.newAction)
        center(fillDataDialog, 500, 500)
        fillDataDialog.setFixedSize(500, 500)
        fillDataDialog.show()

    def valueCheck(self):
        folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择要校验的文件夹", "")
        if folder_Exolorer == "":
            return
        para = {"sourcefolder":folder_Exolorer}
        self.newAction("ValueCheck1", "基础资产校验_V1", para)

    def valueCheck2(self):
        folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择要校验的文件夹", "")
        if folder_Exolorer == "":
            return
        para = {"sourcefolder":folder_Exolorer}
        self.newAction("ValueCheck2", "基础资产校验_V2", para)

    def NonvalueCheck(self):
        folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择要校验的文件夹", "")
        if folder_Exolorer == "":
            return
        para = {"sourcefolder":folder_Exolorer}
        self.newAction("NonValueCheck" ,"不良资产校验", para)

    def PoolvalueCheck(self):
        folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择要校验的文件夹", "")
        if folder_Exolorer == "":
            return
        para = {"sourcefolder":folder_Exolorer}
        self.newAction("PoolValueCheck" ,"池分布校验", para)

    def newAction(self, actionCode, actionName, kwargs):
        self.showMask()
        self.repaint()
        newtab = TabWidget.Tab(actionCode, actionName, kwargs)
        self.workTab.addTab(newtab, QIcon("Resource/icon/Icon_tag.ico"), actionName)
        newtab.logGenerated.connect(self.appenText)
        try:
            newtab.run()
        except Exception as e:
            QMessageBox.critical(self, "错误", traceback.format_exc(), QMessageBox.Ok)
        self.maskwidget.hide()



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
        currentTab = self.workTab.widget(index)
        if currentTab is None:
            return
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
            root.append(XETree.Element(self.str_AssetType))
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

    def showMask(self):
        self.maskwidget.raise_()
        #self.maskwidget.move(self.rect().center())
        self.maskwidget.show()


# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     w = self()
#     sys.exit(app.exec_())