from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QDialog
from PyQt5.QtGui import QIcon
from Utils.ConfigRead import *
from QtGui.gui import Automation
from QtGui.gui import Interface


class Choice(QDialog):

    def __init__(self, parent=None):
        super(Choice, self).__init__(parent)

        self.CreatUi()

    def CreatUi(self):
        # 设置窗口名称
        self.setWindowTitle('自动化测试脚本')
        self.setWindowIcon(QIcon(RESOURSE_PATH + 'lable.png'))
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.setFixedSize(350, 220)
        self.setAcceptDrops(True)
        self.setWindowTitle('脚本选择')
        self.listView = QtWidgets.QListView(self)
        self.toolButton = QtWidgets.QToolButton(self)
        self.toolButton.setText('Web测试')
        self.toolButton.setObjectName('web')
        self.toolButton_2 = QtWidgets.QToolButton(self)
        self.toolButton_2.setText('接口测试')
        self.toolButton_2.setObjectName('interface')

        self.Position()
        self.ButtonBind()

    def ButtonBind(self):
        self.toolButton.clicked.connect(self.web_ui)
        self.toolButton_2.clicked.connect(self.interface_ui)

    def web_ui(self):
        """初始化Automation窗口"""
        # 关闭主窗口
        automation = Automation.Automation()
        styleSheet = automation.readQssFile(QSS_PATH + 'Automator.qss')
        automation.setStyleSheet(styleSheet)
        # 设置子窗口未关闭时无法操作父窗口
        automation.setWindowModality(QtCore.Qt.ApplicationModal)
        automation.show()

    def interface_ui(self):
        widget = Interface.Interface()
        styleSheet = widget.readQssFile(QSS_PATH + 'Interface.qss')
        widget.setStyleSheet(styleSheet)
        # 设置子窗口未关闭时无法操作父窗口
        widget.setWindowModality(QtCore.Qt.ApplicationModal)
        widget.show()

    def Position(self):
        self.listView.setGeometry(QtCore.QRect(0, 0, 350, 220))
        self.toolButton.setGeometry(QtCore.QRect(30, 80, 120, 40))
        self.toolButton_2.setGeometry(QtCore.QRect(200, 80, 120, 40))

    def readQssFile(self, filePath):
        with open(filePath, 'r', encoding='utf-8') as fileObj:
            styleSheet = fileObj.read()
        return styleSheet

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    widget = Choice()
    styleSheet = widget.readQssFile(QSS_PATH+'Choice.qss')
    widget.setStyleSheet(styleSheet)
    widget.show()
    sys.exit(app.exec_())
