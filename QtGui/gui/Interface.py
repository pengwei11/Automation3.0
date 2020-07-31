from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import (QFileDialog, QMessageBox, QDialog)
from PyQt5.QtCore import pyqtSignal, QThread, QTimer
from PyQt5.QtGui import QIcon
from Utils.ParseExcelXlrd import ParseExcelXlrd
from Utils.interface_utils.interface_param import ParamRead
from Utils.WriteFile import YamlWrite
from Utils.ConfigRead import *
from test.interface_test.TestApi import TestApi
from Utils.ParseYaml import ParseYaml
import threading
import traceback
import os, re
import time, shutil


class Interface(QDialog):

    def __init__(self, parent=None):
        super(Interface, self).__init__(parent)
        if os.path.exists(r'E:\Automation3.0\logs\apilog\api-logger.log'):
            os.remove(r'E:\Automation3.0\logs\apilog\api-logger.log')
        self.timer1 = QTimer(self)
        # 记录接口地址
        self.api_dict = {}
        # 记录上一次的日志
        self.old_lines = []
        self.num = 1
        self.newfile = ''
        # 编写用例说明导出数量计算
        self.explainum = 1
        # 记录测试用例模板导出次数
        self.filenum = 1
        # 记录token模板导出此时
        self.tokenfilenum = 1
        self.CreatUi()
        if ParseYaml().ReadAPI_Paramter('API_PATH') is not None:
            self.lineEdit.setText(ParseYaml().ReadAPI_Paramter('API_PATH'))

    def CreatUi(self):
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.setFixedSize(572, 666)
        # 设置窗口名称
        self.setWindowTitle('接口自动化测试脚本')
        self.setWindowIcon(QIcon(RESOURSE_PATH + 'api-resource\\' + 'api-title.jpg'))
        self.tabWidget = QtWidgets.QTabWidget(self)
        self.tabWidget.setObjectName("tabWidget")
        self.tabWidget.setDocumentMode(True)
        # 设置活动页
        self.tabWidget.setCurrentIndex(0)
        self.tab = QtWidgets.QWidget(self)
        self.tab.setObjectName("tab")
        self.tableWidget = QtWidgets.QTableWidget(self.tab)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setRowCount(5)
        # 设置表格标签自适应宽度
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        # 隐藏行表头
        self.tableWidget.verticalHeader().setVisible(False)
        # 设置表格不可编辑
        self.tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        # 设置表格不可选择
        self.tableWidget.setSelectionMode(QtWidgets.QTableWidget.NoSelection )
        # 隐藏滚动条
        self.tableWidget.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        # 设置表头样式
        self.tableWidget.horizontalHeader().setStyleSheet("QHeaderView::section{background:rgb(102,102,102);color:white;}")
        # 设置表头塌陷
        self.tableWidget.horizontalHeader().setHighlightSections(False)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item.setText('Key')
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        item.setText('Value')
        self.tabWidget.addTab(self.tab, "Headers")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.tableWidget_2 = QtWidgets.QTableWidget(self.tab_2)
        self.tableWidget_2.setObjectName("tableWidget_2")
        self.tableWidget_2.setColumnCount(2)
        self.tableWidget_2.setRowCount(5)
        # 设置表格标签自适应宽度
        self.tableWidget_2.horizontalHeader().setStretchLastSection(True)
        # 隐藏行表头
        self.tableWidget_2.verticalHeader().setVisible(False)
        # 设置表格不可编辑
        self.tableWidget_2.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        # 设置表格不可选择
        self.tableWidget_2.setSelectionMode(QtWidgets.QTableWidget.NoSelection)
        # 设置隐藏滚动条
        self.tableWidget_2.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.tableWidget_2.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        # 设置表头样式
        self.tableWidget_2.horizontalHeader().setStyleSheet("QHeaderView::section{background:rgb(102,102,102);color:white;}")
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setVerticalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setVerticalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setHorizontalHeaderItem(0, item)
        item.setText('Key')
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setHorizontalHeaderItem(1, item)
        item.setText('Value')
        self.tabWidget.addTab(self.tab_2, "Body")
        self.listView_2 = QtWidgets.QListView(self)
        self.listView_2.setObjectName("listView_2")
        self.label = QtWidgets.QLabel(self)
        self.label.setObjectName("label")
        self.label.setText('IP地址')
        self.lineEdit = QtWidgets.QLineEdit(self)
        self.lineEdit.setObjectName("lineEdit")
        self.label_2 = QtWidgets.QLabel(self)
        self.label_2.setObjectName("label_2")
        self.label_2.setText('token导入')
        self.label_3 = QtWidgets.QLabel(self)
        self.label_3.setObjectName("label_3")
        self.label_3.setText('请选择excel.xlsx格式表格导入')
        self.label_4 = QtWidgets.QLabel(self)
        self.label_4.setObjectName("label_4")
        self.label_4.setText('测试用例')
        self.toolButton = QtWidgets.QToolButton(self)
        self.toolButton.setObjectName("toolButton")
        self.toolButton.setText('选择文件...')
        self.label_5 = QtWidgets.QLabel(self)
        self.label_5.setObjectName("label_5")
        self.label_5.setText('请选择excel.xlsx格式表格导入')
        self.toolButton_3 = QtWidgets.QToolButton(self)
        self.toolButton_3.setObjectName("toolButton_3")
        self.toolButton_3.setText('选择文件...')
        self.toolButton_2 = QtWidgets.QToolButton(self)
        self.toolButton_2.setObjectName("toolButton_2")
        self.toolButton_2.setText('开始')
        self.toolButton_4 = QtWidgets.QToolButton(self)
        self.toolButton_4.setObjectName("toolButton_4")
        self.toolButton_4.setText('导出模板')
        self.toolButton_5 = QtWidgets.QToolButton(self)
        self.toolButton_5.setObjectName("toolButton_5")
        self.toolButton_5.setText('编写说明')
        self.listWidget_2 = QtWidgets.QListWidget(self)
        self.listWidget_2.setObjectName("listWidget_2")
        self.listWidget_2.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.listWidget_2.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        # self.setCentralWidget(self.centralWidget)

        # self.Position_new()
        self.Position_old()

        self.ButtonBind()

    def Position_old(self):
        self.setFixedSize(880, 666)
        self.treeWidget = QtWidgets.QTreeWidget(self)
        self.treeWidget.setObjectName("TreeWidget")
        # 设置头部标题隐藏
        self.treeWidget.setHeaderHidden(True)
        # 设置垂直滚动条隐藏
        self.treeWidget.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.treeWidget.headerItem().setText(0, '')
        # 给树形窗口设置点击信号
        self.treeWidget.itemDoubleClicked.connect(self.treeClick)
        bgimg2 = RESOURSE_PATH + 'api-resource\\' + 'bgimg1.jpg'
        bgimg2 = bgimg2.replace('\\', '/')
        openpng = RESOURSE_PATH + 'open.png'
        openpng = openpng.replace('\\', '/')
        shrink = RESOURSE_PATH + 'shrink.png'
        shrink = shrink.replace('\\', '/')
        self.treeWidget.setStyleSheet("QTreeWidget{background-image: url(%s);color:#FFFFFF;border:1px solid white;}"
                                      "QTreeView::branch:has-children:!has-siblings:closed,\\"
                                      "QTreeView::branch:closed:has-children:has-siblings{border-image: none; image: url(%s);}\\"
                                      "QTreeView::branch:open:has-children:!has-siblings,\\"
                                      "QTreeView::branch:open:has-children:has-siblings{border-image: none; image: url(%s);}" % (bgimg2, shrink, openpng))
        self.tabWidget.setGeometry(QtCore.QRect(300, 0, 581, 201))
        self.tableWidget.setGeometry(QtCore.QRect(-1, 0, 581, 181))
        self.tableWidget_2.setGeometry(QtCore.QRect(-1, 0, 581, 181))
        self.listView_2.setGeometry(QtCore.QRect(300, 470, 581, 201))
        self.treeWidget.setGeometry(QtCore.QRect(0, 0, 301, 671))
        self.label.setGeometry(QtCore.QRect(320, 490, 54, 12))
        self.lineEdit.setGeometry(QtCore.QRect(420, 490, 113, 20))
        self.label_2.setGeometry(QtCore.QRect(320, 550, 54, 12))
        self.label_3.setGeometry(QtCore.QRect(420, 580, 200, 12))
        self.label_4.setGeometry(QtCore.QRect(320, 620, 54, 12))
        self.toolButton.setGeometry(QtCore.QRect(420, 540, 81, 31))
        self.label_5.setGeometry(QtCore.QRect(420, 650, 250, 12))
        self.toolButton_3.setGeometry(QtCore.QRect(420, 610, 81, 31))
        self.toolButton_2.setGeometry(QtCore.QRect(650, 490, 71, 31))
        self.toolButton_4.setGeometry(QtCore.QRect(770, 490, 71, 31))
        self.toolButton_5.setGeometry(QtCore.QRect(650, 550, 71, 31))
        self.listWidget_2.setGeometry(QtCore.QRect(300, 200, 581, 271))

    def Position_new(self):
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 581, 201))
        self.tableWidget.setGeometry(QtCore.QRect(-1, 0, 572, 181))
        self.tableWidget_2.setGeometry(QtCore.QRect(-1, 0, 572, 181))
        self.listView_2.setGeometry(QtCore.QRect(0, 470, 581, 201))
        self.label.setGeometry(QtCore.QRect(20, 490, 54, 12))
        self.lineEdit.setGeometry(QtCore.QRect(120, 490, 113, 20))
        self.label_2.setGeometry(QtCore.QRect(20, 550, 54, 12))
        self.label_3.setGeometry(QtCore.QRect(120, 580, 200, 12))
        self.label_4.setGeometry(QtCore.QRect(20, 620, 54, 12))
        self.toolButton.setGeometry(QtCore.QRect(120, 540, 81, 31))
        self.label_5.setGeometry(QtCore.QRect(120, 650, 250, 12))
        self.toolButton_3.setGeometry(QtCore.QRect(120, 610, 81, 31))
        self.toolButton_2.setGeometry(QtCore.QRect(350, 490, 71, 31))
        self.toolButton_4.setGeometry(QtCore.QRect(470, 490, 71, 31))
        self.toolButton_5.setGeometry(QtCore.QRect(350, 550, 71, 31))
        self.listWidget_2.setGeometry(QtCore.QRect(0, 200, 581, 271))

    def ButtonBind(self):
        self.toolButton.clicked.connect(self.import_token_excel)
        self.toolButton_3.clicked.connect(self.import_case_excel)
        # 发射多线程弹窗信号
        self.message = message(self)
        self.message.signal.connect(self.box)

        # 开始执行
        self.toolButton_2.clicked.connect(self.runcase)
        self.toolButton_4.clicked.connect(self.exportclick)
        self.toolButton_5.clicked.connect(self.explainclick)

    def explainclick(self):
        """
        测试用例编写说明
        """
        # 选择保存测试用例的路径
        self.explain_path = QFileDialog.getExistingDirectory(self, "请选择编写说明保存路径")
        explainPath = EXCELTEMPLATE_PATH + '接口自动化测试脚本使用说明.doc'
        # 判断该文件是否已存在该路径下
        if self.explain_path != '':
            if os.path.exists(self.explain_path + '\\接口自动化测试脚本使用说明.doc'):
                shutil.copy2(explainPath, self.explain_path + '\\接口自动化测试脚本使用说明(%s).doc' % self.explainum)
                self.explainum = self.explainum + 1
            else:
                self.explainum = 1
                shutil.copy2(explainPath, self.explain_path + '\\接口自动化测试脚本使用说明.doc')
            self.msg = '接口自动化测试脚本使用说明导出成功！'
            self.message.start()

    def exportclick(self):
        try:
            # 选择保存测试用例的路径
            self.file_path = QFileDialog.getExistingDirectory(self, "请选择用例模板保存路径")
            excelPath = EXCELTEMPLATE_PATH + '接口测试模板.xlsx'
            tokenexcelPath = EXCELTEMPLATE_PATH + 'token.xlsx'
            # 判断该文件是否已存在该路径下
            if self.file_path != '':
                if os.path.exists(self.file_path + '\\接口测试模板.xlsx'):
                    shutil.copy2(excelPath, self.file_path + '\\接口测试模板(%s).xlsx' % self.filenum)
                    self.filenum = self.filenum+1
                else:
                    self.filenum = 1
                    shutil.copy2(excelPath, self.file_path + '\\接口测试模板.xlsx')
                if os.path.exists(self.file_path + '\\token.xlsx'):
                    shutil.copy2(tokenexcelPath, self.file_path + '\\token(%s).xlsx' % self.tokenfilenum)
                    self.tokenfilenum = self.tokenfilenum+1
                else:
                    self.tokenfilenum = 1
                    shutil.copy2(tokenexcelPath, self.file_path + '\\token.xlsx')
                self.msg = '用例模板导出成功！'
                self.message.start()
        except Exception as e:
            print(e)

    def box(self):
        # 重名提示
        box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr(self.msg), QMessageBox.NoButton, self)
        yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
        # 确定按钮绑定回车
        yr_btn.setShortcut(QtCore.Qt.Key_Return)
        box.exec_()

    def import_token_excel(self):
        """导入token文件"""
        try:
            self.filepath, filetype = QFileDialog.getOpenFileName(self, "请导入测试用例", '', 'Excel files(*.xlsx)')
            if self.filepath == '':
                if self.label_3.text() == '请选择excel.xlsx格式表格导入':
                    self.filepath = self.newfile
                    self.label_3.setText('请选择excel.xlsx格式表格导入')
            else:
                if self.label_5.text() != '请选择excel.xlsx格式表格导入':
                    if self.filepath in self.case_filepath:
                        self.msg = '导入的token文件已存在于测试用例中，请在用例文件中去除后重新导入'
                        self.message.start()
                        return
                self.newfile = self.filepath
                self.newfile = self.filepath.replace('/', '\\')
                # 将token文件地址写入yaml文件中
                YamlWrite(CONFIG_PATH+'interface_config\\'+'API_Parameter.yaml').Write_Yaml_Updata('TOKEN_FILE_PATH', self.filepath)
                self.label_3.setText(self.newfile)
        except Exception:
            self.msg = '导入失败，请重试'
            self.message.start()

    def import_case_excel(self):
        """导入token文件"""
        try:
            self.case_filepath, filetype = QFileDialog.getOpenFileNames(self, "请导入测试用例", '', 'Excel files(*.xlsx)')
            if not self.case_filepath:
                if self.label_5.text() == '请选择excel.xlsx格式表格导入':
                    self.label_5.setText('请选择excel.xlsx格式表格导入')
            else:
                if self.label_3.text() != '请选择excel.xlsx格式表格导入':
                    if self.filepath in self.case_filepath:
                        self.msg = '导入的测试用例包含token文件，请去除后重新导入'
                        self.message.start()
                        return
                self.label_5.setText(self.case_filepath[0].replace('/', '\\')+'...')
                # 将用例集合写入yaml文件中
                YamlWrite(CONFIG_PATH+'interface_config\\'+'API_Parameter.yaml').Write_Yaml_Updata('FILE_PATH', self.case_filepath)
                # 开启多线程读取excel，防止主程序卡死
                # self.Position_old()
                self.toolButton_2.setDisabled(True)
                self.toolButton_3.setDisabled(True)
                self.toolButton_2.setStyleSheet(
                    "QToolButton{font:11pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}"
                )
                self.toolButton_3.setStyleSheet(
                    "QToolButton{font:9pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}"
                )
                self.importexcelthread = threading.Thread(target=self.importExcelThread)
                self.importexcelthread.setDaemon(True)
                self.importexcelthread.start()
        except Exception:
            self.msg = '导入失败，请重试'
            self.message.start()

    def importExcelThread(self):
        try:
            # 初始化字典地址
            self.api_dict = {}
            self.treeWidget.clear()
            for index, file in enumerate(self.case_filepath):
                self.parseexcelxlrd = ParseExcelXlrd(file)
                self.paramread = ParamRead(self.parseexcelxlrd)
                # 判断导入的是否是接口文件
                templater = ['序号', '接口地址', '接口工作表', '请求类型', '是否执行', '是否生成用例', '执行结束时间', '通过', '失败']
                api_title = self.parseexcelxlrd.getRowValue(0, 2)
                if templater == api_title:
                    # 获取文件名称
                    file_name = os.path.splitext(os.path.basename(file))[0]
                    self.api_dict[file_name] = file
                    self.item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
                    # 增加悬浮提示
                    self.item_0.setToolTip(0, file)
                    self.treeWidget.topLevelItem(index+1 - 1).setText(0, file_name)
                    for api_index, api in enumerate(self.paramread.get_api_sheet()):
                        # 添加用例判断
                        if self.parseexcelxlrd.getCellValue(api, 1, 1) == '接口名称（描述）':
                            item_1 = QtWidgets.QTreeWidgetItem(self.item_0)
                            item_1.setToolTip(0, self.paramread.get_api_sheet_path_new(api))
                            self.treeWidget.topLevelItem(index+1 - 1).child(api_index+1 - 1).setText(0, api)
                else:
                    self.msg = '导入失败，请检查用例'
                    self.message.start()
                    self.label_5.setText('请选择excel.xlsx格式表格导入')
        except Exception as e:
            print(traceback.print_exc(e))
            self.msg = '导入失败，请检查用例'
            self.message.start()
            self.label_5.setText('请选择excel.xlsx格式表格导入')
        finally:
            self.toolButton_2.setDisabled(False)
            self.toolButton_3.setDisabled(False)
            self.toolButton_2.setStyleSheet(
                "QToolButton{font:11pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}"
            )
            self.toolButton_3.setStyleSheet(
                "QToolButton{font:9pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}"
            )

    def treeClick(self, item):
        '''
        根据父节点，子节点，获取相应的excel信息
        :return:
        '''
        # if self.label_16.text() == '正在读取用例文件，请勿进行其他操作...':
        #     self.messages = '请用例导入完成后重试！'
        #     self.message.start()
        #     return
        # 获取父节点索引
        F_node = item.parent()
        F_index = self.treeWidget.indexOfTopLevelItem(F_node)
        # 获取父节点内容
        # 判断点击层级，0为子节点,-1为父节点，点击父节点时只有收缩功能  or self.plainTextEdit.isSignalConnected() is True
        if F_index != -1:
            # self.treeWidget.setDisabled(True)
            # 禁用导入按钮
            # self.toolButton_3.setDisabled(True)
            self.F_text = self.treeWidget.topLevelItem(F_index).text(0)
            # 获取子节点内容
            self.S_text = self.treeWidget.currentItem().text(0)
            # 开启线程
            self.select_excle = SelectExcelThread(ParseExcelXlrd(self.api_dict[self.F_text]), self.S_text)
            self.select_excle.select_date.connect(self.readExcel)
            self.select_excle.run()

    def readExcel(self, data_dict, data, headres):
        '''
        开启多线程，防止读取excel文件卡顿
        :return:
        '''
        try:
            self.tableWidget_2.clear()
            item = QtWidgets.QTableWidgetItem()
            self.tableWidget_2.setHorizontalHeaderItem(0, item)
            item.setText('Key')
            item = QtWidgets.QTableWidgetItem()
            self.tableWidget_2.setHorizontalHeaderItem(1, item)
            item.setText('Value')
            if len(data) <= 5:
                if self.tableWidget_2.rowCount() > 5:
                    for num in range(self.tableWidget_2.rowCount()-5):
                        self.tableWidget_2.removeRow(self.tableWidget_2.rowCount()-1)
                for index, (d, v) in enumerate(data_dict.items()):
                    item = QtWidgets.QTableWidgetItem(str(d))
                    self.tableWidget_2.setItem(index, 0, item)
                    item = QtWidgets.QTableWidgetItem(str(v))
                    self.tableWidget_2.setItem(index, 1, item)
            else:
                for num in range(len(data)-5):
                    # 动态添加行
                    self.tableWidget_2.insertRow(self.tableWidget_2.rowCount())
                for index, (d, v) in enumerate(data_dict.items()):
                    item = QtWidgets.QTableWidgetItem(str(d))
                    self.tableWidget_2.setItem(index, 0, item)
                    item = QtWidgets.QTableWidgetItem(str(v))
                    self.tableWidget_2.setItem(index, 1, item)
            if len(headres.keys()) <= 5:
                if self.tableWidget.rowCount() > 5:
                    for num in range(self.tableWidget.rowCount()-5):
                        self.tableWidget.removeRow(self.tableWidget.rowCount()-1)
                for index, (d, v) in enumerate(headres.items()):
                    item = QtWidgets.QTableWidgetItem(str(d))
                    self.tableWidget.setItem(index, 0, item)
                    item = QtWidgets.QTableWidgetItem(str(v))
                    self.tableWidget.setItem(index, 1, item)
            else:
                for num in range(len(headres.keys())-5):
                    # 动态添加行
                    self.tableWidget.insertRow(self.tableWidget.rowCount())
                for index, (d, v) in enumerate(headres.items()):
                    item = QtWidgets.QTableWidgetItem(str(d))
                    self.tableWidget.setItem(index, 0, item)
                    item = QtWidgets.QTableWidgetItem(str(v))
                    self.tableWidget.setItem(index, 1, item)
        except Exception:
            self.msg = '读取失败，请重试'
            self.message.start()

    def runcase(self):
        "运行测试用例"
        if self.lineEdit.text() == '':
            self.msg = '测试地址不能为空！'
            self.message.start()
        elif re.match(r"^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$", self.lineEdit.text()) is None and re.match(
                r'[^\s]*[.com|.cn]', self.lineEdit.text()) is None:
            self.msg = '请输入正确格式的测试地址！'
            self.message.start()
        elif self.label_5.text() == '请选择excel.xlsx格式表格导入':
            self.msg = '请导入测试用例！'
            self.message.start()
        else:
            # try:
            #     self.parseexcelxlrd.writeCellValue(self.sheets[0], 50, 50, '测试')
            # except PermissionError:
            #     self.msg = "请先关闭用例文件，再运行测试用例"
            #     self.message.start()
            #     return
            # except Exception as e:
            #     print(e)
            #     self.msg = '运行错误，请重新运行或重启软件'
            #     self.message.start()
            #     return
            # 将IP地址写入yaml文件中
            # 清空日志框
            self.listWidget_2.clear()
            # 重置记录
            # self.num = 1
            YamlWrite(CONFIG_PATH + 'interface_config\\' + 'API_Parameter.yaml').Write_Yaml_Updata('API_PATH',
                                                                                                   self.lineEdit.text())
            # 增加线程运行测试用例
            self.runcasethread = threading.Thread(target=TestApi().test_api)
            self.runcasethread.setDaemon(True)
            self.runcasethread.start()
            self.timer1.timeout.connect(self.OutPut)
            self.timer1.start(200)
            self.toolButton.setDisabled(True)
            self.toolButton_2.setDisabled(True)
            self.toolButton_3.setDisabled(True)
            self.toolButton.setStyleSheet(
                "QToolButton{font:9pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}"
            )
            self.toolButton_2.setStyleSheet(
                "QToolButton{font:11pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}"
            )
            self.toolButton_3.setStyleSheet(
                "QToolButton{font:9pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}"
            )

    def OutPut(self):
        """
        持续输出日志
        :return:
        """
        try:
            # self.listWidget_2.clear()
            base_dir = APILOGS_PATH
            l = os.listdir(base_dir)
            l.sort(key=lambda fn: os.path.getmtime(base_dir + fn)
            if not os.path.isdir(base_dir + fn) else 0)
            if 'api-logger' not in l[-1]:
                print('')
            else:
                logpath = os.path.join(base_dir, l[-1])
                filesize = os.path.getsize(logpath)
                blocksize = 1024
                dat_file = open(logpath, 'rb')
                last_line = ""
                if filesize > blocksize:
                    maxseekpoint = (filesize // blocksize)
                    dat_file.seek((maxseekpoint - 1) * blocksize)
                elif filesize:
                    dat_file.seek(0, 0)
                lines = dat_file.read().splitlines()
                if self.num == 1:
                    self.old_lines = lines
                    for line in lines:
                        count = self.listWidget_2.count()
                        self.listWidget_2.insertItem(count, line.decode('utf-8'))
                        self.listWidget_2.scrollToBottom()
                        self.num += 1
                else:
                    # 计算此次获取到的集合与上一次的集合之差
                    new_lines = sorted(set(lines)-set(self.old_lines), key=lines.index)
                    for new in new_lines:
                        count = self.listWidget_2.count()
                        self.listWidget_2.insertItem(count, new.decode('utf-8'))
                        self.listWidget_2.scrollToBottom()
                    self.old_lines = lines
                if self.runcasethread.isAlive() is False:
                    self.timer1.stop()
                    self.toolButton_2.setDisabled(False)
                    self.toolButton.setDisabled(False)
                    self.toolButton_3.setDisabled(False)
                    self.toolButton.setStyleSheet(
                        "QToolButton{font:9pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}"
                    )
                    self.toolButton_2.setStyleSheet(
                        "QToolButton{font:11pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}"
                    )
                    self.toolButton_3.setStyleSheet(
                        "QToolButton{font:9pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}"
                    )
        except Exception as e:
            print(e)

    def readQssFile(self, filePath):
        with open(filePath, 'r', encoding='utf-8') as fileObj:
            styleSheet = fileObj.read()
        return styleSheet

    def closeEvent(self, event):
        """
        重写closeEvent方法，实现dialog窗体关闭时执行一些代码
        :param event: close()触发的事件
        :return: None
        """
        box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("您确定要退出吗？"), QMessageBox.NoButton, self)
        yr_btn = box.addButton(self.tr("确定"), QMessageBox.YesRole)
        box.addButton(self.tr("取消"), QMessageBox.NoRole)
        box.exec_()
        if box.clickedButton() == yr_btn:
            self.close()
        else:
            event.ignore()

class message(QThread):
    signal = pyqtSignal()

    def __init__(self, Interface):
        super(message, self).__init__()
        self.automaint = Interface

    def run(self):
        self.signal.emit()

class SelectExcelThread(QThread):
    # 通过类成员对象定义信号
    select_date = pyqtSignal(dict, list, dict)

    def __init__(self, book, s_text):
        super(SelectExcelThread, self).__init__()
        self.book = book
        self.s_text = s_text

    # 处理业务逻辑
    def run(self):
        # 读取点击的父节点路径
        paramread = ParamRead(self.book)
        # 获取请求参数
        data = paramread.get_api_sheet_data(self.s_text)
        # 获取正常请求中的参数
        describe = self.book.getMergeColumnValue(self.s_text, 1)
        describeindex = [i for i, x in enumerate(describe) if x == '描述']
        describe_row = int(describeindex[0]) + 2
        value = self.book.getRowValue(self.s_text, describe_row + 1)
        del value[0]
        del value[-4:]
        # 将请求参数和数据整合成字典
        data_dict = dict(zip(data, value))
        # 请求头
        headres = eval(paramread.get_api_sheet_header(self.s_text))
        self.select_date.emit(data_dict, data, headres)



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    widget = Interface()
    styleSheet = widget.readQssFile(QSS_PATH+'Interface.qss')
    widget.setStyleSheet(styleSheet)
    widget.show()
    sys.exit(app.exec_())
