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

class PoolStaticDataInsert_Excel(QDialog):
    submitted = pyqtSignal(str, str, dict)
    def __init__(self, parent):
        super().__init__(parent)
        self.title = '静态池Excel文件导入'
        self.str_AssetType = "AssetType"
        self.configFileDirectory = "./Config"
        self.filetype = 'PoolStaticInsertSql'
        self.FileName = "Config.xml"
        self.dbConnectionStr = 'DRIVER={SQL Server};SERVER=172.16.6.143\MSSQL;DATABASE=PortfolioManagement;UID=sa;PWD=PasswordGS2017'
        self.sql = 'exec PortfolioManagement.DVImport.usp_getPoolStaticColumnMapping '
        self.configFilePath = os.path.join(self.configFileDirectory, self.FileName)
        self.__initUI()

    def __initUI(self):
        self.setObjectName("Main")
        stylesheet = open("Source/FirstFileCheck.qss", "r").read()
        self.setWindowTitle(self.title)
        self.setWindowModality(Qt.ApplicationModal)
        self.setStyleSheet(stylesheet)

        self.body = QWidget()
        self.body.setStyleSheet("")
        self.body.setObjectName("body")

        self.bodyLayout = QGridLayout(self.body)
        self.bodyLayout.setObjectName("bodyLayout")
        #选择入库的文件夹
        self.audio_Dir = QRadioButton(text="文件夹")
        self.audio_Dir.setObjectName("audio_Dir")
        self.audio_Dir.click()

        self.lineEdit_selectDir = QLineEdit()
        self.lineEdit_selectDir.setObjectName("lineEdit_selectDir")
        self.lineEdit_selectDir.setReadOnly(True)

        self.button_selectDir = QPushButton()
        self.button_selectDir.setText("..")
        self.button_selectDir.clicked.connect(self.openDir)
        # 选择入库的文件
        self.audio_File = QRadioButton(text="文件")
        self.audio_File.setObjectName("audio_File")
        #self.audio_File.click()

        self.lineEdit_selectFile = QLineEdit()
        self.lineEdit_selectFile.setObjectName("lineEdit_selectFile")
        self.lineEdit_selectFile.setReadOnly(True)

        self.button_selectFIle = QPushButton()
        self.button_selectFIle.setText("..")
        self.button_selectFIle.clicked.connect(self.openFile)

        self.button_Group1 = QButtonGroup()
        self.button_Group1.addButton(self.audio_Dir)
        self.button_Group1.addButton(self.audio_File)

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setSizeConstraint(QLayout.SetDefaultConstraint)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")

        self.setLayout(self.verticalLayout)
        self.bodyLayout.addWidget(self.audio_Dir, 0, 0)
        self.bodyLayout.addWidget(self.lineEdit_selectDir, 0, 1)
        self.bodyLayout.addWidget(self.button_selectDir, 0, 2)
        self.bodyLayout.addWidget(self.audio_File, 1, 0)
        self.bodyLayout.addWidget(self.lineEdit_selectFile, 1, 1)
        self.bodyLayout.addWidget(self.button_selectFIle, 1, 2)

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
    def openFile(self) -> None:
        file_Exolorer = QFileDialog.getOpenFileName(self, caption='选择文件', filter='*.xlsx ;; *.xls')
        self.lineEdit_selectFile.setText(file_Exolorer[0])
        return
    def openDir(self) -> None:
        folder_Exolorer = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "")
        self.lineEdit_selectDir.setText(folder_Exolorer)
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

    def submit(self):
        audioDir = self.audio_Dir.isChecked()
        #audioFile = self.audio_File.isChecked()
        DirPath = self.lineEdit_selectDir.text()
        FilePath = self.lineEdit_selectFile.text()
        parameters = dict()
        parameters["audioType"] = '1' if audioDir else '2'
        parameters["DirPath"] = str(DirPath)
        parameters["filePath"] = str(FilePath)
        parameters["outputFile"] = ''    #没用到，但是这边需要用来凑数的
        self.hide()
        self.submitted.emit("PoolStaticDataInsert_Excel", "静态池Excel导入", parameters)
        self.close()






