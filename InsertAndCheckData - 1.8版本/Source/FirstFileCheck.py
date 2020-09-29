from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import os
import xml.etree.ElementTree as XETree
from Source.StyleComboBox import *

class FirstFileCheck(QDialog):
    submitted = pyqtSignal(str, str, dict)
    def __init__(self, parent):
        super().__init__(parent)
        self.str_xmlFIle = "xmlFIle"
        self.str_template = "templateFile"
        self.str_OutputFolder = "OutputFolder"
        self.configFileDirectory = "./Config"
        self.configFileName = "config.xml"
        self.configFilePath = os.path.join(self.configFileDirectory, self.configFileName)
        self.Init()

    def Init(self):
        self.setObjectName("Main")
        stylesheet = open("Source/FirstFileCheck.qss", "r").read()
        self.setWindowTitle("第一步校验")
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

        # 添加下拉框
        self.label_selectFile = QLabel(text="待校验文件夹：")
        self.label_selectFile.setAlignment(Qt.AlignCenter)
        self.label_selectFile.setObjectName("label_selectFile")

        self.lineEdit_selectFile = QLineEdit()
        self.lineEdit_selectFile.setObjectName("lineEdit_selectFile")
        self.lineEdit_selectFile.setReadOnly(True)

        self.button_selectFIle = QPushButton()
        self.button_selectFIle.setText("..")
        self.button_selectFIle.clicked.connect(self.open)

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

        # 资产配置文件：
        self.label_pyFIle = QLabel(text="资产配置文件：")
        self.label_pyFIle.setAlignment(Qt.AlignCenter)
        self.label_pyFIle.setObjectName("label_pyFIle")

        self.comboBox_pyFIle = StyledComboBox()
        self.comboBox_pyFIle.setEditable(False)
        self.comboBox_pyFIle.setObjectName("comboBox_pyFIle")
        self.comboBox_pyFIle.setMaxVisibleItems(10)

        self.initComboBox(self.str_template, self.comboBox_pyFIle)
        self.comboBox_pyFIle.setCurrentIndex(0)
        self.comboBox_pyFIle.currentIndexChanged.connect(
            lambda: self.EventCombox_Output(self.comboBox_pyFIle, self.str_template))

        # 方案：
        self.label_xmlFIle = QLabel(text="转换识别方案：")
        self.label_xmlFIle.setAlignment(Qt.AlignCenter)
        self.label_xmlFIle.setObjectName("label_xmlFIle")

        self.comboBox_xmlFIle = StyledComboBox()
        self.comboBox_xmlFIle.setEditable(False)
        self.comboBox_xmlFIle.setObjectName("comboBox_xmlFile")
        self.comboBox_xmlFIle.setMaxVisibleItems(10)

        self.initComboBox(self.str_xmlFIle, self.comboBox_xmlFIle)
        self.comboBox_xmlFIle.setCurrentIndex(0)
        self.comboBox_xmlFIle.currentIndexChanged.connect(
            lambda: self.EventCombox_Output(self.comboBox_xmlFIle, self.str_xmlFIle))

        # self.button_distinguishXml = QPushButton()
        # self.button_distinguishXml.setText("自动选择")
        # self.button_distinguishXml.clicked.connect(self.distinguishDoc)
        #

        self.bodyLayout = QGridLayout(self.body)
        self.bodyLayout.setObjectName("bodyLayout")
        self.bodyLayout.addWidget(self.label_selectFile, 0, 0)
        self.bodyLayout.addWidget(self.lineEdit_selectFile, 0, 1)
        self.bodyLayout.addWidget(self.button_selectFIle, 0, 2)
        self.bodyLayout.addWidget(self.label_pyFIle, 1, 0)
        self.bodyLayout.addWidget(self.comboBox_pyFIle, 1, 1)
        self.bodyLayout.addWidget(self.label_xmlFIle, 2, 0)
        self.bodyLayout.addWidget(self.comboBox_xmlFIle, 2, 1)
        self.bodyLayout.addWidget(self.label_OutputFolder, 3, 0)
        self.bodyLayout.addWidget(self.comboBox_outputfolder, 3, 1)

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
        setting = QSettings("./Config/settingFirst.ini", QSettings.IniFormat)
        index = setting.value(self.str_OutputFolder)
        if index is not None:
            self.comboBox_outputfolder.setCurrentIndex(int(index))
        index = setting.value(self.str_xmlFIle)
        if index is not None:
            self.comboBox_xmlFIle.setCurrentIndex(int(index))
        index = setting.value(self.str_template)
        if index is not None:
            self.comboBox_pyFIle.setCurrentIndex(int(index))

    def initComboBox(self, section, combobox):
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find(section)
        items = node.getchildren()
        model = combobox.model()
        for i in range(0, len(items)):
            if section == self.str_xmlFIle:
                if '第一步' in items[i].attrib["Name"]:
                    combobox.addItem(items[i].attrib["Name"], items[i].text)
            elif section == self.str_OutputFolder:
                combobox.addItem(items[i].text, str(i))
            elif section == self.str_template:
                if '第一步' in items[i].attrib["Name"]:
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
        setting = QSettings("./Config/settingFirst.ini", QSettings.IniFormat)
        setting.setValue(self.str_OutputFolder, self.comboBox_outputfolder.currentIndex())
        setting.setValue(self.str_xmlFIle, self.comboBox_xmlFIle.currentIndex())
        setting.setValue(self.str_template, self.comboBox_pyFIle.currentIndex())
        a0.accept()

    def submit(self):
        filepath = self.lineEdit_selectFile.text()
        pyPath = self.comboBox_pyFIle.currentData()
        outputFolder = self.comboBox_outputfolder.currentText()
        xmlfilepath = self.comboBox_xmlFIle.currentData()
        parameters = dict()
        parameters["inputFile"] = filepath
        parameters["pyPath"] = pyPath
        parameters["XmlFile"] = xmlfilepath
        parameters["outputFile"] = outputFolder
        self.hide()
        self.submitted.emit("FirstCheck", "第一步校验", parameters)
        self.close()