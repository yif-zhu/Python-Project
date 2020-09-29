from PyQt5.QtWidgets import QComboBox, QStyledItemDelegate
from PyQt5 import QtCore, QtGui
class StyledComboBox(QComboBox):
    def __init__(self):
        super().__init__()
        #设置该属性后可以通过QComboBox QAbstractItemView::item设置样式
        self.setItemDelegate(QStyledItemDelegate(self))

class DropDownMenu(StyledComboBox):
    selected = QtCore.pyqtSignal()
    def __init__(self):
        super().__init__()
        self.currentIndexChanged.connect(lambda : self.resetIndex())


    def resetIndex(self) -> None:
        self.setCurrentIndex(0)
        self.view().setRowHidden(0, True)



