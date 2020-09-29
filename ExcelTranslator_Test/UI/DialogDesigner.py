# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'DialogDesigner.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Main(object):
    def setupUi(self, Main):
        Main.setObjectName("Main")
        Main.resize(801, 707)
        self.layoutWidget = QtWidgets.QWidget(Main)
        self.layoutWidget.setGeometry(QtCore.QRect(-10, 0, 811, 711))
        self.layoutWidget.setObjectName("layoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.layoutWidget)
        self.verticalLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.title = QtWidgets.QWidget(self.layoutWidget)
        self.title.setObjectName("title")
        self.verticalLayout.addWidget(self.title)
        self.body = QtWidgets.QWidget(self.layoutWidget)
        self.body.setStyleSheet("QLabel{background-color: #F3F5FA; border:1px solid #E6E6E6;\n"
"max-width:137px;min-width:137px;min-hegiht:30px;text-align: right}")
        self.body.setObjectName("body")
        self.gridLayoutWidget = QtWidgets.QWidget(self.body)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(100, 150, 651, 301))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.label = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 2, 1, 1, 1)
        self.comboBox = QtWidgets.QComboBox(self.gridLayoutWidget)
        self.comboBox.setObjectName("comboBox")
        self.gridLayout.addWidget(self.comboBox, 2, 2, 1, 1)
        self.verticalLayout.addWidget(self.body)
        self.bottom = QtWidgets.QWidget(self.layoutWidget)
        self.bottom.setObjectName("bottom")
        self.verticalLayout.addWidget(self.bottom)

        self.retranslateUi(Main)
        QtCore.QMetaObject.connectSlotsByName(Main)

    def retranslateUi(self, Main):
        _translate = QtCore.QCoreApplication.translate
        Main.setWindowTitle(_translate("Main", "Form"))
        self.label.setText(_translate("Main", "TextLabel"))

