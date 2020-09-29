from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import os
import xml.etree.ElementTree as XETree
from Source.StyleComboBox import *

class CompareCNABS(QDialog):
    submitted = pyqtSignal(str, str, dict)
    def __init__(self, parent):
        super().__init__(parent)
        self.str_xmlFIle = "xmlFIle"
        self.str_assetType = "AssetType"
        self.str_OutputFolder = "OutputFolder"
        self.configFileDirectory = "./Config"
        self.configFileName = "config.xml"
        self.configFilePath = os.path.join(self.configFileDirectory, self.configFileName)
        self.Init()

    def Init(self):
        self.setObjectName("Main")
        stylesheet = open("Source/CompareCNABS.qss", "r").read()
        self.setWindowTitle("校验范围选择")
        self.setWindowModality(Qt.ApplicationModal)
        self.setStyleSheet(stylesheet)

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setSizeConstraint(QLayout.SetDefaultConstraint)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.setLayout(self.verticalLayout)

        self.body = QWidget()
        self.body.setStyleSheet("")
        self.body.setObjectName("body")

        #根据时间选择产品
        self.audio_Time = QRadioButton(text="时间")
        #self.audio_Time.layoutDirection(Qt.)
        self.audio_Time.setObjectName("audio_Time")

        self.audio_Time.click()

        self.timeEdit = QDateTimeEdit()
        self.timeEdit.setObjectName("timeEdit")
        self.timeEdit.setDateTime(QDateTime.currentDateTime())

        #根据资产类型选择产品
        self.audio_assetType = QRadioButton(text="资产类型")
        self.audio_assetType.setObjectName("audio_assetType")

        self.comboBox_assetType = StyledComboBox()
        self.comboBox_assetType.setEditable(False)
        self.comboBox_assetType.setObjectName("comboBox_assetType")
        self.comboBox_assetType.setMaxVisibleItems(10)

        self.initComboBox(self.str_assetType, self.comboBox_assetType)
        self.comboBox_assetType.setCurrentIndex(0)
        self.comboBox_assetType.currentIndexChanged.connect(
            lambda: self.EventCombox_Output(self.comboBox_assetType, self.str_assetType))

        #根据产品TrustId校验
        self.audio_TrustID = QRadioButton(text="产品TrustID")
        self.audio_TrustID.setObjectName("audio_TrustID")

        self.text_TrustId = QLineEdit()
        self.text_TrustId.setObjectName("text_TrustId")

        self.button_Group2 = QButtonGroup()    #用于将radio分组，避免冲突
        self.button_Group2.addButton(self.audio_Time)
        self.button_Group2.addButton(self.audio_TrustID)
        self.button_Group2.addButton(self.audio_assetType)

        # 选择产品的类别
        self.label_type = QLabel(text="资产类别：")
        self.label_type.setAlignment(Qt.AlignCenter)
        self.label_type.setObjectName("label_type")

        self.audio_normal = QRadioButton(text="正常资产")
        self.audio_normal.setObjectName("audio_normal")
        self.audio_normal.click()

        self.audio_NonPool = QRadioButton(text="不良资产")
        self.audio_NonPool.setObjectName("audio_NonPool")

        self.button_Group1 = QButtonGroup()
        self.button_Group1.addButton(self.audio_normal)
        self.button_Group1.addButton(self.audio_NonPool)

        # 输出至
        self.label_OutputFolder = QLabel(text="结果输出目录：")
        self.label_OutputFolder.setAlignment(Qt.AlignCenter)
        self.label_OutputFolder.setObjectName("label_OutputFolder")

        self.comboBox_outputfolder = StyledComboBox()
        self.comboBox_outputfolder.setEditable(False)
        self.comboBox_outputfolder.setObjectName("comboBox_outputfolder")
        self.comboBox_outputfolder.addItem("新增输出目录...", "New")
        self.comboBox_outputfolder.setMaxVisibleItems(10)

        self.initComboBox(self.str_OutputFolder, self.comboBox_outputfolder)
        self.comboBox_outputfolder.setCurrentIndex(0)
        self.comboBox_outputfolder.activated.connect(
            lambda: self.EventCombox_Output(self.comboBox_outputfolder, self.str_OutputFolder))



        self.bodyLayout = QGridLayout(self.body)
        self.bodyLayout.setObjectName("bodyLayout")
        self.bodyLayout.addWidget(self.audio_Time, 0, 0)
        self.bodyLayout.addWidget(self.timeEdit, 0, 1)
        self.bodyLayout.addWidget(self.audio_assetType, 2, 0)
        self.bodyLayout.addWidget(self.comboBox_assetType, 2, 1)
        self.bodyLayout.addWidget(self.audio_TrustID, 1, 0)
        self.bodyLayout.addWidget(self.text_TrustId, 1, 1)
        self.bodyLayout.addWidget(self.label_type, 3, 0)
        self.bodyLayout.addWidget(self.audio_normal, 3, 1)
        self.bodyLayout.addWidget(self.audio_NonPool, 3, 2)
        self.bodyLayout.addWidget(self.label_OutputFolder, 4, 0)
        self.bodyLayout.addWidget(self.comboBox_outputfolder, 4, 1)

        self.verticalLayout.addWidget(self.body)
        self.bottom = QWidget()
        self.bottom.setObjectName("bottom")
        self.verticalLayout.addWidget(self.bottom)

        self.bottom_layount = QVBoxLayout()
        self.bottom.setLayout(self.bottom_layount)
        self.button_submit = QPushButton()
        # self.pushButton.setGeometry(QRect(0, 320, 171, 91))
        self.button_submit.setObjectName("submit")
        self.button_submit.setText("确认")
        self.button_submit.clicked.connect(self.submit)
        self.bottom_layount.addWidget(self.button_submit)

        # 恢复控件状态
        setting = QSettings("./Config/settingCompare.ini", QSettings.IniFormat)
        index = setting.value(self.str_OutputFolder)
        if index is not None:
            self.comboBox_outputfolder.setCurrentIndex(int(index))
        index = setting.value(self.str_assetType)
        if index is not None:
            self.comboBox_assetType.setCurrentIndex(int(index))

    def initComboBox(self, section, combobox):
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find(section)
        items = node.getchildren()
        model = combobox.model()
        for i in range(0, len(items)):
            if section == self.str_xmlFIle:
                combobox.addItem(items[i].attrib["Name"], items[i].text)
            elif section == self.str_OutputFolder:
                combobox.addItem(items[i].text, str(i))
            elif section == self.str_assetType:
                combobox.addItem(items[i].attrib["Name"], items[i].text)

    def EventCombox_Output(self, combobox, flag):
        folder_Exolorer = ""
        if combobox.currentData() == "New":
            folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择文件夹", "")

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
                if isExistsItem:
                    QMessageBox.warning(self, '警告', '该项已存在' + existsName, QMessageBox.Yes)
                    combobox.setCurrentIndex(1)
                else:
                    element = XETree.Element("Folder")
                    element.text = folder_Exolorer
                    node.append(element)

                    combobox.addItem(folder_Exolorer, str(len(items) + 1))
                    combobox.setCurrentIndex(len(items))

                    #indent(node)
                    tree.write(self.configFilePath, encoding='utf-8', xml_declaration=True)
            else:
                combobox.setCurrentIndex(1)
        # print(combobox.currentText())

    def open(self) -> None:
        folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "")
        self.lineEdit_selectFile.setText(folder_Exolorer)
        return

    def closeEvent(self, a0: QCloseEvent) -> None:
        setting = QSettings("./Config/settingCompare.ini", QSettings.IniFormat)
        setting.setValue(self.str_OutputFolder, self.comboBox_outputfolder.currentIndex())
        setting.setValue(self.str_assetType, self.comboBox_assetType.currentIndex())
        a0.accept()

    def isNumber(self, TrustId):
        flag = isinstance(TrustId, int)
        return flag

    def submit(self):
        audio_Time = self.audio_Time.isChecked()
        audio_TrustId = self.audio_TrustID.isChecked()
        audio_assetType = self.audio_assetType.isChecked()
        audio_normal = self.audio_normal.isChecked()
        audio_NonPool = self.audio_NonPool.isChecked()
        type = 1
        pool_Type = 1
        if audio_Time:  # 用什么方式读取
            type = 1
        elif audio_TrustId:
            type = 3
        elif audio_assetType:
            type = 2

        if audio_normal:  # 读取什么类型的资产
            pool_Type = 1
        elif audio_NonPool:
            pool_Type = 2
        TrustId = self.text_TrustId.text()
        if type == 3 and not TrustId.isdigit():
            self.text_TrustId.setText('')
            self.text_TrustId.setPlaceholderText("产品TrustId类型错误，应为数字！")
            return
        ImportTime = self.timeEdit.text()
        outputFolder = self.comboBox_outputfolder.currentText()
        assetType = self.comboBox_assetType.currentData()
        parameters = dict()
        parameters["type"] = str(type)
        parameters["poolType"] = str(pool_Type)
        parameters["ImportTime"] = ImportTime
        parameters["TrustId"] = TrustId
        parameters["outputFile"] = outputFolder
        parameters["AssetType"] = assetType
        self.hide()
        self.submitted.emit("CompareCNABS", "数据对比", parameters)
        self.close()