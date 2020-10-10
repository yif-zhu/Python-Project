import sys

from Source import MainWindow
import xml.etree.ElementTree as XETree
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import pyodbc
import datetime
import pdfplumber
import calendar
import pandas as pd
import numpy as np

from PyQt5.QtWidgets import QApplication, QMainWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)
    #openfile.Ui_Dialog()
    w = MainWindow.MainWindow()
    sys.exit(app.exec_())

