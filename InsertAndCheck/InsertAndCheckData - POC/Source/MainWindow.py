from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow
import os
import xml.etree.ElementTree as XETree
from Source.StyleComboBox import *
from Source.BaseInfoCheck import *
from Source.FirstFileCheck import *
from Source.SecondFileCheck import *
from Source.ThirdFileCheck import *
from Source.CompareCNABS import *
from Source.ActionManager import AcitonManager
from Source.AssetTypeManager import AssetTypeManager
from Source.SettingManager import *
from Source import TabWidget
import shutil
import traceback

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
    # 屏幕尺寸
    screen_width = 0
    screen_height = 0
    def __init__(self):
        super().__init__()
        self.width = 400
        self.height = 200
        self.initConfig()
        self.initUI()

    # 初始化配置文件
    def initConfig(self):
        if not os.path.exists(self.configFileDirectory):
            os.mkdir(self.configFileDirectory)
        if not os.path.exists(self.configFilePath):
            # open(self.configFilePath, "wb").write(bytes("", encoding="utf-8"))
            root = XETree.Element('Root')  # 创建节点
            tree = XETree.ElementTree(root)  # 创建文档
            root.append(XETree.Element(self.str_OutputFolder))
            root.append(XETree.Element(self.str_xmlFIle))
            root.append(XETree.Element(self.str_templatefile))
            root.append(XETree.Element(self.str_AssetType))
            self.indent(root)  # 增加换行符
            tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)

    # 初始化窗口界面
    def initUI(self):
        # 设置样式
        self.setStyleSheet(open("Source/MainWindow.qss", 'r').read())
        # 获取屏幕尺寸
        screen = QDesktopWidget().screenGeometry()
        self.screen_width = screen.width()
        self.screen_height = screen.height()
        # 设置中心窗口部件为QTextEdit
        self.verticalSplitter = QSplitter(Qt.Vertical)
        self.setCentralWidget(self.verticalSplitter)

        # self.scroll.setGeometry(QtCore.QRect(100,100, 2000, 1000))


        # self.scorllTextEdit = QScrollArea()

        self.textEdit = QTextEdit()
        self.textEdit.setText('执行记录：')
        # self.textEdit.setStyleSheet("background:white;height:20%")
        self.textEdit.setReadOnly(True)

        #self.verticalSplitter.addWidget(self.workTab)
        self.verticalSplitter.addWidget(self.textEdit)
        # 定义一系列的Action
        #产品基础信息校验
        # BaseInfoCheck = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_run.ico'), '基础信息校验')
        # BaseInfoCheck.setShortcut('Ctrl+B')
        # BaseInfoCheck.setStatusTip('基础信息校验')
        # BaseInfoCheck.clicked.connect(self.BaseInfoCheckValue)

        # 第一步校验
        FirstCheck = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_run.ico'), '第一步校验')
        FirstCheck.setShortcut('Ctrl+N')
        FirstCheck.setStatusTip('第一步校验')
        FirstCheck.clicked.connect(self.FirstCheckValue)

        # 第二步校验入库
        self.SecondCheck = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_run.ico'), '第二步校验')
        self.SecondCheck.setText("第二步校验")
        self.SecondCheck.setToolTip('第二步校验')
        self.SecondCheck.clicked.connect(self.SecondCheckValue)

        # 第三步校验
        self.ThirdCheck = QtWidgets.QPushButton(QIcon('Resource/icon/Icon_run.ico'), '第三步校验')
        self.ThirdCheck.setText("第三步校验")
        self.ThirdCheck.setToolTip('第三步校验')
        self.ThirdCheck.clicked.connect(self.ThirdCheckValue)

        # 数据对比XXXXX
        # self.DataComPare = QtWidgets.QPushButton(QIcon("Resource/icon/Icon_multiRun.ico"), '数据对比')
        # self.DataComPare.setText("数据对比")
        # self.DataComPare.setShortcut('Ctrl+M')
        # self.DataComPare.setStatusTip('数据对比')
        # self.DataComPare.clicked.connect(self.DataComPareAbs)

        # 配置管理
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
        # tb1.addWidget(BaseInfoCheck)
        tb1.addWidget(FirstCheck)
        tb1.addWidget(self.SecondCheck)
        tb1.addWidget(self.ThirdCheck)
        #tb1.addWidget(self.DataComPare)
        # tb1.addWidget(self.button_saveAs)
        # tb1.addWidget(saveAllAction)
        tb1.addWidget(self.comboBox_setting)

        # self.verticalSplitter.setStyleSheet("background-color: rgb(222, 222, 222);")

        self.statusBar()

        # 遮罩
        self.maskwidget = QWidget(self)
        self.maskwidget.setObjectName("Mask")
        self.maskwidget.setGeometry(0, 0, 1400, 800)
        #self.center(self.maskwidget)
        self.maskwidget.raise_()

        self.masklabel = QLabel(self.maskwidget)
        #self.maskwidget.setGeometry(0, 0, 1400, 800)
        #self.center(self.masklabel)
        self.masklabel.setText("loading...")

        # self.masklabel.move(self.maskwidget.rect().center())
        # self.loadingGif = QMovie('Resource/icon/loading-image.gif')
        # self.masklabel.setMovie(self.loadingGif)
        # self.loadingGif.start()
        self.maskwidget.hide()

        # show
        self.setObjectName("MainWindow")
        self.setGeometry(0, 0, 1400, 800)
        self.center(self)
        self.setWindowTitle("入库工具1.7_POC")
        self.setWindowIcon(QIcon('Resource/icon/Icon_table.ico'))
        # self.center()
        self.show()
        #self.showMaximized()

        # self.showMask()
        # fillDataDialog.setGeometry((self.maximumSize().width() - 500)/2,(self.maximumSize().height()-800)/2, 500, 800 )

    #给xml增加换行符
    def indent(self, elem, level=0):
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

    # def BaseInfoCheckValue(self):   #基础信息校验校验
    #     # folder_BaseInfo = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "")
    #     # parameters = dict()
    #     # parameters["BaseInfoPath"] = folder_BaseInfo
    #     # self.newAction(self, 'BaseInfoCheck', '基础信息校验', parameters)
    #     baseInfoCheck = BaseInfoCheck(self)
    #     baseInfoCheck.raise_()
    #     baseInfoCheck.submitted.connect(self.newAction)
    #
    #     # center(fileCheck, 500, 500).203
    #     baseInfoCheck.setFixedSize(500, 400)
    #     baseInfoCheck.show()


    #第一步校验方法
    def FirstCheckValue(self):
        fileCheck = FirstFileCheck(self)
        fileCheck.raise_()
        fileCheck.submitted.connect(self.newAction)

        #center(fileCheck, 500, 500).203
        fileCheck.setFixedSize(500, 500)
        fileCheck.show()

    def SecondCheckValue(self):
        SecondCheck = SecondFileCheck(self)
        SecondCheck.raise_()
        SecondCheck.submitted.connect(self.newAction)

        # center(fileCheck, 500, 500)
        SecondCheck.setFixedSize(500, 500)
        SecondCheck.show()

    def ThirdCheckValue(self):    #三种情况，一种根据时间，一种根据资产类型，一个根据TrustId
        ThirdCheck = ThirdFileCheck(self)
        ThirdCheck.raise_()
        ThirdCheck.submitted.connect(self.newAction)

        # center(fileCheck, 500, 500)
        ThirdCheck.setFixedSize(500, 500)
        ThirdCheck.show()

    # def DataComPareAbs(self):
    #     CompareAbs = CompareCNABS(self)
    #     CompareAbs.raise_()
    #     CompareAbs.submitted.connect(self.newAction)
    #
    #     # center(fileCheck, 500, 500)
    #     CompareAbs.setFixedSize(500, 500)
    #     CompareAbs.show()

    def showManager(self, index: int):
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

    # 窗口居中显示
    def center(self, object):
        screen = QDesktopWidget().screenGeometry()
        # +120 稍微向上平移
        size = object.geometry()
        object.move((screen.width() - size.width()) / 2,
                  (screen.height() - size.height()) / 2)
        #self.move((screen.width() - width) / 2, (screen.height() - height) / 2 - 120)

    def newAction(self, actionCode, actionName, kwargs):
        #self.showMask()
        self.appenText('{0}，正在校验中，请稍等。。。。。'.format(actionName))
        self.repaint()
        newtab = TabWidget.Tab(actionCode, actionName, kwargs)
        #self.workTab.addTab(newtab, QIcon("Resource/icon/Icon_tag.ico"), actionName)
        newtab.logGenerated.connect(self.appenText)
        try:
            newtab.run()
        except Exception as e:
            QMessageBox.critical(self, "错误", traceback.format_exc(), QMessageBox.Ok)
        self.maskwidget.hide()

    def appenText(self, str):
        if '开始执行任务：' in str:
            self.textEdit.clear()
        if '正在校验中，请稍等：' in str:
            self.textEdit.clear()
        self.textEdit.append(str)
        self.textEdit.moveCursor(QTextCursor.End)

    def showMask(self):
        self.maskwidget.raise_()
        #self.maskwidget.move(self.rect().center())
        self.maskwidget.show()

    def open(self):
        folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "")
       #self.lineEdit_selectFile.setText(folder_Exolorer)
        return folder_Exolorer