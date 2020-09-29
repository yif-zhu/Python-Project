import sys

from UI import MainWindow
import xml.etree.ElementTree as XETree
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5 import QtCore, QtWidgets, QtGui
if __name__ == '__main__':
    app = QApplication(sys.argv)
    #openfile.Ui_Dialog()
    w = MainWindow.MainWindow()
    sys.exit(app.exec_())

