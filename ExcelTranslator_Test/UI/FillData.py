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
import shutil
from FileTranslator import matchXml
from UI.StyleComboBox import StyledComboBox

class FillData(QDialog):
    submitted = pyqtSignal(str, str, dict)
    #mode = 0，运行单个，=1 运行多个
    def __init__(self, parent, mode):
        super().__init__(parent)
        self.mode = mode
        self.str_templatefile = "templateFile"
        self.str_xmlFIle = "xmlFIle"
        self.str_OutputFolder = "OutputFolder"
        self.configFileDirectory = "./Config"
        self.configFileName = "config.xml"
        self.configFilePath = os.path.join(self.configFileDirectory, self.configFileName)

        self.__initUI()

    def __initUI(self):
        self.setObjectName("Main")
        stylesheet = open("UI/FillData.qss", "r").read()
        self.setWindowTitle("内容填充")
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


        #添加下拉框
        self.label_selectFile = QLabel(text="原始文件：")
        self.label_selectFile.setAlignment(Qt.AlignCenter)
        self.label_selectFile.setObjectName("label_selectFile")

        self.lineEdit_selectFile = QLineEdit()
        self.lineEdit_selectFile.setObjectName("lineEdit_selectFile")
        self.lineEdit_selectFile.setReadOnly(True)

        self.button_selectFIle = QPushButton()
        self.button_selectFIle.setText("选择文件")
        self.button_selectFIle.clicked.connect(self.open)

        #输出至
        self.label_OutputFolder = QLabel(text="输出目录：")
        self.label_OutputFolder.setAlignment(Qt.AlignCenter)
        self.label_OutputFolder.setObjectName("label_OutputFolder")

        self.comboBox_outputfolder = StyledComboBox()
        self.comboBox_outputfolder.setEditable(False)
        self.comboBox_outputfolder.setObjectName("comboBox_outputfolder")
        self.comboBox_outputfolder.addItem("新增输出目录...", "New")
        self.comboBox_outputfolder.setMaxVisibleItems(10)

        self.initComboBox(self.str_OutputFolder, self.comboBox_outputfolder)
        self.comboBox_outputfolder.setCurrentIndex(0)
        self.comboBox_outputfolder.currentIndexChanged.connect(lambda  :self.EventCombox_Output(self.comboBox_outputfolder, self.str_OutputFolder))

        #方案：
        self.label_xmlFIle = QLabel(text="转换方案：")
        self.label_xmlFIle.setAlignment(Qt.AlignCenter)
        self.label_xmlFIle.setObjectName("label_xmlFIle")

        self.comboBox_xmlFIle = StyledComboBox()
        self.comboBox_xmlFIle.setEditable(False)
        self.comboBox_xmlFIle.setObjectName("comboBox_xmlFile")
        self.comboBox_xmlFIle.setMaxVisibleItems(10)

        self.initComboBox(self.str_xmlFIle, self.comboBox_xmlFIle)
        self.comboBox_xmlFIle.setCurrentIndex(0)
        self.comboBox_xmlFIle.currentIndexChanged.connect(lambda: self.EventCombox_Output(self.comboBox_xmlFIle, self.str_xmlFIle))

        self.button_distinguishXml = QPushButton()
        self.button_distinguishXml.setText("自动选择")
        self.button_distinguishXml.clicked.connect(self.distinguishDoc)
        #模板
        self.label_templatefile = QLabel(text="目标文件：")
        self.label_templatefile.setAlignment(Qt.AlignCenter)
        self.label_templatefile.setObjectName("label_templatefile")

        self.comboBox_templatefile = StyledComboBox()
        self.comboBox_templatefile.setEditable(False)
        self.comboBox_templatefile.setObjectName("comboBox_templatefile")
        self.comboBox_templatefile.setMaxVisibleItems(10)

        self.initComboBox(self.str_templatefile, self.comboBox_templatefile)
        self.comboBox_templatefile.setCurrentIndex(0)
        self.comboBox_templatefile.currentIndexChanged.connect(
            lambda: self.EventCombox_Output(self.comboBox_templatefile, self.str_templatefile))

        self.bodyLayout = QGridLayout(self.body)
        self.bodyLayout.setObjectName("bodyLayout")
        self.bodyLayout.addWidget(self.label_selectFile, 0, 0)
        self.bodyLayout.addWidget(self.lineEdit_selectFile, 0, 1)
        self.bodyLayout.addWidget(self.button_selectFIle, 0, 2)
        self.bodyLayout.addWidget(self.label_xmlFIle, 1, 0)
        self.bodyLayout.addWidget(self.comboBox_xmlFIle, 1, 1)
        self.bodyLayout.addWidget(self.button_distinguishXml, 1, 2)
        self.bodyLayout.addWidget(self.label_OutputFolder, 2, 0)
        self.bodyLayout.addWidget(self.comboBox_outputfolder, 2, 1)
        self.bodyLayout.addWidget(self.label_templatefile, 3, 0)
        self.bodyLayout.addWidget(self.comboBox_templatefile, 3, 1)


        self.verticalLayout.addWidget(self.body)
        self.bottom = QWidget()
        self.bottom.setObjectName("bottom")
        self.verticalLayout.addWidget(self.bottom)

        self.bottom_layount = QVBoxLayout()
        self.bottom.setLayout(self.bottom_layount)
        self.button_submit = QPushButton()
        #self.pushButton.setGeometry(QRect(0, 320, 171, 91))
        self.button_submit.setObjectName("submit")
        self.button_submit.setText("确认")
        self.button_submit.clicked.connect(self.submit)
        self.bottom_layount.addWidget(self.button_submit)

        # 恢复控件状态
        setting = QSettings("./Config/setting.ini", QSettings.IniFormat)
        index = setting.value(self.str_OutputFolder)
        if index is not None:
            self.comboBox_outputfolder.setCurrentIndex(int(index))
        index = setting.value(self.str_xmlFIle)
        if index is not None:
            self.comboBox_xmlFIle.setCurrentIndex(int(index))
        index = setting.value(self.str_templatefile)
        if index is not None:
            self.comboBox_templatefile.setCurrentIndex(int(index))

    def initComboBox(self, section, combobox):

        isExistsSection = False
        tree = XETree.parse(self.configFilePath)
        node = tree.getroot().find(section)
        items = node.getchildren()
        model = combobox.model()
        for i in range(0, len(items)):
            if section == self.str_xmlFIle:
                combobox.addItem(items[i].attrib["Name"], items[i].text)
            elif section == self.str_OutputFolder:
                combobox.addItem(items[i].text, str(i))
            elif section == self.str_templatefile:
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

    def distinguishDoc(self):
        filepath = self.lineEdit_selectFile.text()
        if filepath == "":
            QMessageBox.critical(self, "警告", "选择的文件为空", QMessageBox.Ok)
            return
        self.button_distinguishXml.setDisabled(True)
        self.button_distinguishXml.setText("正在识别...")
        self.repaint()
        xmlName = ""
        try:
            xmlName = matchXml.main(filepath)
        except Exception as e:
            self.button_distinguishXml.setDisabled(False)
            self.button_distinguishXml.setText("自动识别")
            QMessageBox.critical(self, "错误", str(e), QMessageBox.Ok)
        if xmlName is None or xmlName == "":
            self.button_distinguishXml.setDisabled(False)
            self.button_distinguishXml.setText("自动识别")
            message = QMessageBox(self)
            message.setWindowTitle("完成")
            message.setWindowIcon(QIcon("Resource/icon/Icon_table.ico"))
            message.setText("没有匹配的方案")
            message.show()
            return
        index = self.comboBox_xmlFIle.findText(xmlName, Qt.MatchContains)
        self.comboBox_xmlFIle.setCurrentIndex(index)
        self.button_distinguishXml.setDisabled(False)
        self.button_distinguishXml.setText("自动识别")


    def open(self) -> None:
        if self.mode == 1:
            folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "")
            self.lineEdit_selectFile.setText(folder_Exolorer)
            return
        file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择文件', filter=("*.xlsx"))
        if file_Exolorer[0]:
            self.lineEdit_selectFile.setText(file_Exolorer[0])
    def closeEvent(self, a0: QCloseEvent) -> None:
        setting = QSettings("./Config/setting.ini", QSettings.IniFormat)
        setting.setValue(self.str_OutputFolder, self.comboBox_outputfolder.currentIndex())
        setting.setValue(self.str_xmlFIle, self.comboBox_xmlFIle.currentIndex())
        setting.setValue(self.str_templatefile, self.comboBox_templatefile.currentIndex())
        a0.accept()

    def submit(self):
        filepath = self.lineEdit_selectFile.text()
        outputFolder = self.comboBox_outputfolder.currentText()
        xmlfilepath = self.comboBox_xmlFIle.currentData()
        templatefile = self.comboBox_templatefile.currentData()
        parameters = dict()
        parameters["inputFile"] = filepath
        parameters["XmlFile"] = xmlfilepath
        parameters["outputFile"] = outputFolder
        parameters["templateFile"] = templatefile
        self.hide()
        self.submitted.emit("FillData", "内容填充",parameters)
        self.close()






