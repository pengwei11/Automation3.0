#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: Automation.py
@time: 2020/1/6 14:12
@desc:
"""

from PyQt5.QtWidgets import (QWidget, QFileDialog, QMessageBox, QDesktopWidget, QMenu, QAction, QDialog, QApplication)
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QColor,QPalette,QPixmap,QBrush,QIcon, QStandardItem
from PyQt5.QtCore import pyqtSignal, QThread, QTimer
from Utils.ParseExcelWin import ParseExcelWin
from Utils.ConfigRead import *
from action.PageAction import *
from selenium.common.exceptions import *
from Utils.ParseExcelXlrd import ParseExcelXlrd 
from Utils.ParseExcel import ParseExcel
from test.TestOneCase import TestOneCase
from Utils.TimerCase import TimerCase
from Utils.WriteFile import YamlWrite
from test.testPaperless import TestPaperless
from Utils.ParseYaml import ParseYaml
from QtGui.gui import Choice
import sys, os, shutil, re
import threading
import ctypes
import inspect


class Automation(QWidget):

    def __init__(self, parent=None):
        super(Automation, self).__init__(parent)
        try:
            if os.path.exists(r'E:\Automation3.0\logs\mainlog\logger.log'):
                os.remove(r'E:\Automation3.0\logs\mainlog\logger.log')
            if os.path.exists(r'E:\Automation3.0\logs\apilog\case-logger.log'):
                os.remove(r'E:\Automation3.0\logs\apilog\case-logger.log')
        except Exception:
            pass
        # 初始化定时器
        self.timer1 = QTimer(self)
        self.timer2 = QTimer(self)
        self.timer3 = QTimer(self)
        self.timer4 = QTimer(self)
        # 记录重复导入用例的路径
        self.newfile = ''
        # 记录readexcel线程开启
        self.readexcel_num = 0
        # 记录旧用例的名称
        self.CreatUi()
        # 测试浏览器工具
        self.pageaction = PageAction()
        # 记录导入用例次数,同时记录是否导入文件，或新建文件
        self.import_num = 0
        # 记录新增或修改后的用例的主要内容
        self.newcase = {}
        # 关键字
        self.keyword = ['openBrowsers', 'quitBrowser', 'back', 'foword', 'refresh', 'js_scroll_top',
                        'js_scroll_end', 'getUrl', 'sleep', 'clear', 'inputValue', 'uploadFile', 'uploadFiles',
                        'assertTitle', 'assert_string_in_pageSource', 'assertEqule', 'assertElementEqule', 'assertLen',
                        'assertElement', 'assertUrl', 'getTitle', 'getPageSource', 'switchToFrame', 'switchToDefault',
                        'click', 'wait_find_element', 'not_wait_find_element', 'text_wait_find_element',
                        'not_text_wait_find_element', 'Enter', 'move_to_element', 'asserrSelect', 'if', 'break']
        # 记录单条测试用例log日志
        self.OneCaseLogList = []
        # 编写用例说明导出数量计算
        self.explainum = 1
        # 记录测试用例模板导出次数
        self.filenum = 1
        # 初始化yaml写入类
        self.writeyaml = YamlWrite(CONFIG_PATH+'Parameter.yaml')
        # 记录模块用例数量
        self.MoudleSum = []
        # 如果parameter.py中有IP地址，则键入
        if ParseYaml().ReadParameter('IP') is not None and  ParseYaml().ReadParameter('IP') != '暂停运行':
            self.lineEdit.setText(ParseYaml().ReadParameter('IP'))


    def CreatUi(self):
        # 设置窗口大小
        self.setFixedSize(940, 800)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        # 设置窗口永远置顶
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        # 获取屏幕大小，打开窗口默认在屏幕中间
        screen = QDesktopWidget().screenGeometry()
        self.setGeometry((screen.width() - 940) / 2, (screen.height() - 800) / 2, 940, 800)
        # 设置窗口名称
        self.setWindowTitle('Web自动化测试脚本')
        self.setWindowIcon(QIcon(RESOURSE_PATH + 'web-title.jpg'))
        palette = QPalette()
        palette.setBrush(QPalette.Background, QBrush(QPixmap(RESOURSE_PATH + 'bgimg.png')))
        self.setPalette(palette)
        # 运行窗口
        self.listView_3 = QtWidgets.QListView(self)
        self.listView_3.setObjectName('run_interface')
        
        # 用例编号显示框
        self.listView = QtWidgets.QListView(self)
        self.listView.setObjectName('case_header')

        # 用例标题显示框
        self.listView_4 = QtWidgets.QListView(self)
        self.listView_4.setObjectName('case_header')

        # 关键字辅助框提示
        self.label = QtWidgets.QLabel('关键字辅助框', self)
        self.label.setObjectName('main_tips')

        # 脚本编辑窗口提示
        self.label_13 = QtWidgets.QLabel('脚本编辑', self)
        self.label_13.setObjectName('main_tips_frame')
        
        # log查看窗口提示
        self.label_14 = QtWidgets.QLabel('log查看窗', self)
        self.label_14.setObjectName('main_tips')
        
        # 树状图测试用例提示
        self.label_20 = QtWidgets.QLabel('测试用例框', self)
        self.label_20.setObjectName('main_tips')
        
        # IP地址文字
        self.label_2 = QtWidgets.QLabel('测试地址', self)
        self.label_2.setObjectName('run_param_tips')

        # 浏览器文字
        self.label_3 = QtWidgets.QLabel('浏览器', self)
        self.label_3.setObjectName('run_param_tips')

        # 模块选择
        self.label_4 = QtWidgets.QLabel('模块选择', self)
        self.label_4.setObjectName('run_param_tips')

        # 循环次数
        self.label_5 = QtWidgets.QLabel('循环次数', self)
        self.label_5.setObjectName('run_param_tips')

        # 测试报告文字
        self.label_6 = QtWidgets.QLabel('测试报告', self)
        self.label_6.setObjectName('run_param_tips')

        # IP地址输入框
        self.lineEdit = QtWidgets.QLineEdit(self)
        self.lineEdit.setObjectName('ip_param')


        # 浏览器选择框
        self.comboBox = QtWidgets.QComboBox(self)
        self.comboBox.setObjectName('browser_param')
        self.comboBox.addItem('Google Chrome')
        self.comboBox.addItem('FireFox')
        # model = self.comboBox.model()
        # entry = QStandardItem('FireFox')
        # entry.setBackground(QColor('rgba(113,113,113,0.5)'))
        # entry.setForeground(QColor('white'))
        # model.appendRow(entry)
        # entry = QStandardItem('Google Chrome')
        # entry.setBackground(QColor('rgba(113,113,113,0.5)'))
        # entry.setForeground(QColor('white'))
        # model.appendRow(entry)
        self.comboBox.setCurrentText('Google Chrome')

        # 模块选择框
        self.comboBox_2 = QtWidgets.QComboBox(self)
        self.comboBox_2.setObjectName('module_param')


        # 执行用例循环测试选择框
        self.spinBox = QtWidgets.QSpinBox(self)
        self.spinBox.setRange(1, 5)  # 设置下界和上界
        # 不能输入值
        self.spinBox.setFocusPolicy(True)
        self.spinBox.setObjectName('loop_param')

        # 生成与不生成选择器
        self.radioButton = QtWidgets.QRadioButton('生成', self)
        self.radioButton.setObjectName('report_param')
        self.radioButton_2 = QtWidgets.QRadioButton('不生成', self)
        self.radioButton_2.setObjectName('report_param')
        self.radioButton_2.setChecked(True)

        # 报告生成路径
        self.label_15 = QtWidgets.QLabel('', self)
        # 设置路径默认隐藏
        self.label_15.setHidden(True)
        self.label_15.setObjectName('report_file_param')

        # 总用例数文字
        self.label_7 = QtWidgets.QLabel('总用例数', self)
        self.label_7.setObjectName('test_case_param')

        # 已运行数文字
        self.label_8 = QtWidgets.QLabel('已运行数', self)
        self.label_8.setObjectName('test_case_param')

        # 运行时间文字
        self.label_9 = QtWidgets.QLabel('运行时间', self)
        self.label_9.setObjectName('test_case_param')

        # 三根下划线
        self.line = QtWidgets.QFrame(self)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName('line')
        self.line_2 = QtWidgets.QFrame(self)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName('line')
        self.line_3 = QtWidgets.QFrame(self)
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName('line')

        # 总用例数显示
        self.label_10 = QtWidgets.QLabel('0', self)
        self.label_10.setObjectName('test_case_show')

        # 已运行数显示
        self.label_11 = QtWidgets.QLabel('0', self)
        self.label_11.setObjectName('test_case_show')

        # 运行时间显示
        self.label_12 = QtWidgets.QLabel('00:00:00', self)
        self.label_12.setObjectName('test_case_show')

        # 启动按钮
        self.toolButton = QtWidgets.QToolButton(self)
        self.toolButton.setText('导入')
        self.toolButton.setObjectName('button_enable')
        """点击事件"""

        # 导出模板按钮
        self.toolButton_15 = QtWidgets.QToolButton(self)
        self.toolButton_15.setText('导出模板')
        self.toolButton_15.setObjectName('button_enable')
        """点击事件"""

        # 编写说明按钮
        self.toolButton_16 = QtWidgets.QToolButton(self)
        self.toolButton_16.setText('编写说明')
        self.toolButton_16.setObjectName('button_enable')
        """点击事件"""

        # 暂停按钮
        self.toolButton_13 = QtWidgets.QToolButton(self)
        self.toolButton_13.setText('启动')
        self.toolButton_13.setObjectName('button_enable')
        """点击事件"""

        # 导入按钮
        self.toolButton_14 = QtWidgets.QToolButton(self)
        self.toolButton_14.setText('暂停')
        self.toolButton_14.setDisabled(True)
        self.toolButton_14.setObjectName('button_display')
        """点击事件"""

        # 同步用例按钮
        self.toolButton_17 = QtWidgets.QToolButton(self)
        self.toolButton_17.setText('结束测试')
        self.toolButton_17.setDisabled(True)
        self.toolButton_17.setObjectName('button_display')


        # 导入用例名称显示
        self.label_16 = QtWidgets.QLabel(self)
        self.label_16.setText('请选择excel.xlsx格式表格导入')
        self.label_16.setObjectName('case_name')


        # 用例树形显示框
        self.treeWidget = QtWidgets.QTreeWidget(self)
        self.treeWidget.headerItem().setText(0, '')
        bgimg2 = RESOURSE_PATH + 'bgimg2.png'
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
        
        # 设置头部标题隐藏
        self.treeWidget.setHeaderHidden(True)
        # 设置垂直滚动条隐藏
        self.treeWidget.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)


        # 关键字辅助窗口
        self.listWidget = MyCurrentQueue(self)
        self.listWidget.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.listWidget.setStyleSheet("MyCurrentQueue{font:10pt;color:#FFFFFF;font-weight:75;border-radius:3px;background-color: #363636;border:1px solid white;}")
        self.listWidget.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        for i in range(28):
            # 新建item（inputValue）
            item = QtWidgets.QListWidgetItem()
            item.setSizeHint(QtCore.QSize(14, 24))
            # 设置字体以及颜色
            item.setTextAlignment(QtCore.Qt.AlignLeft)
            self.listWidget.addItem(item)
            if i % 2 == 0:
                # 设置item背景颜色
                self.listWidget.item(i).setBackground(QColor('#696969'))
                self.listWidget.setSortingEnabled(False)
            else:
                # 设置item背景颜色
                self.listWidget.item(i).setBackground(QColor('#363636'))
                self.listWidget.setSortingEnabled(False)

        # 新建item（inputValue）
        item = self.listWidget.item(0)
        item.setText("inputValue")
        # 设置悬浮提示
        self.listWidget.item(0).setToolTip('<b>Key:</b>inputValue(输入框输入值)'
                                           '<br><b>param:</b>location(定位方式)'
                                           '<br><b>param:</b>locator(定位表达式)'
                                           '<br><b>param:</b>value(输入值)')
        # 新建item（click）
        item = self.listWidget.item(1)
        item.setText("click")
        # 设置悬浮提示
        self.listWidget.item(1).setToolTip('<b>Key:</b>click(点击事件)'
                                           '<br><b>param:</b>location(定位方式)'
                                           '<br><b>param:</b>locator(定位表达式)')

        # 新建item（openBrowsers）
        item = self.listWidget.item(2)
        item.setText("openBrowsers")
        # 设置悬浮提示
        self.listWidget.item(2).setToolTip('<b>Key:</b>openBrowsers(打开指定浏览器)'
                                           '<br><b>param:</b>value(浏览器类型:Google Chrome,FireFox)')

        # 新建item（getUrl）
        item = self.listWidget.item(3)
        item.setText("getUrl")
        # 设置悬浮提示
        self.listWidget.item(3).setToolTip('<b>Key:</b>getUrl(加载至指定网页)'
                                           '<br><b>param:</b>value(目标网址)')


        # 新建item（if）
        item = self.listWidget.item(4)
        item.setText("if")
        # 设置悬浮提示
        self.listWidget.item(4).setToolTip('<b>Key:</b>inputValue(输入框输入值)'
                                           '<br><b>param:</b>location(判断条件)'
                                           '<br><b>param:</b>locator(成立条件)'
                                           '<br><b>param:</b>value(满足条件)')

        # 新建item（sleep）
        item = self.listWidget.item(5)
        item.setText("sleep")
        # 设置悬浮提示
        self.listWidget.item(5).setToolTip('<b>Key:</b>sleep(强制等待)'
                                           '<br><b>param:</b>value(强制等待时间,单位:s)')

        # 新建item（assertEqule）
        item = self.listWidget.item(6)
        item.setText("assertEqule")
        # 设置悬浮提示
        self.listWidget.item(6).setToolTip('<b>Key:</b>assertEqule(检查目标元素与指定值是否相等)'
                                           '<br><b>param:</b>location(定位方式)'
                                           '<br><b>param:</b>locator(定位表达式)'
                                           '<br><b>param:</b>value(指定值)')

        # 新建item（assertLen）
        item = self.listWidget.item(7)
        item.setText("assertLen")
        # 设置悬浮提示
        self.listWidget.item(7).setToolTip('<b>Key:</b>assertLen(检查目标元素的值的长度与指定长度是否相等)'
                                           '<br><b>param:</b>location(定位方式)'
                                           '<br><b>param:</b>locator(定位表达式)'
                                           '<br><b>param:</b>value(指定长度)')

        # 新建item（assertElement）
        item = self.listWidget.item(8)
        item.setText("assertElement")
        # 设置悬浮提示
        self.listWidget.item(8).setToolTip('<b>Key:</b>assertElement(检查目标元素是否存在)'
                                           '<br><b>param:</b>location(定位方式)'
                                           '<br><b>param:</b>locator(定位表达式)'
                                           '<br><b>param:</b>value(True or False)')

        # 新建item（assertUrl）
        item = self.listWidget.item(9)
        item.setText("assertUrl")
        # 设置悬浮提示
        self.listWidget.item(9).setToolTip('<b>Key:</b>assertUrl(检查当前网址与目标网址是否一致)'
                                           '<br><b>param:</b>value(目标网址)')

        # 新建item（assertTitle）
        item = self.listWidget.item(10)
        item.setText("assertTitle")
        # 设置悬浮提示
        self.listWidget.item(10).setToolTip('<b>Key:</b>assertTitle(检查标题是否与目标标题相同)'
                                           '<br><b>param:</b>value(目标标题)')

        # 新建item（uploadFile）
        item = self.listWidget.item(11)
        item.setText("uploadFile")
        # 设置悬浮提示
        self.listWidget.item(11).setToolTip('<b>Key:</b>uploadFile(上传单个文件)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)'
                                            '<br><b>param:</b>value(文件本地地址,需要使用/或者\\\)')

        # 新建item（uploadFiles）
        item = self.listWidget.item(12)
        item.setText("uploadFiles")
        # 设置悬浮提示
        self.listWidget.item(12).setToolTip('<b>Key:</b>uploadFiles(上传多个文件)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)'
                                            '<br><b>param:</b>value(文件本地地址,需要使用/或者\\\)')

        # 新建item（clear）
        item = self.listWidget.item(13)
        item.setText("clear")
        # 设置悬浮提示
        self.listWidget.item(13).setToolTip('<b>Key:</b>clear(清空输入框)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)')

        # 新建item（wait_find_element）
        item = self.listWidget.item(14)
        item.setText("wait_find_...")
        # 设置悬浮提示
        self.listWidget.item(14).setToolTip('<b>Key:</b>wait_find_element(等待元素出现后进行下一步操作)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)')

        # 新建item（not_wait_find_element）
        item = self.listWidget.item(15)
        item.setText("not_wait_...")
        # 设置悬浮提示
        self.listWidget.item(15).setToolTip('<b>Key:</b>not_wait_find_element(等待元素消失后进行下一步操作)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)')

        # 新建item（text_wait_find_element）
        item = self.listWidget.item(16)
        item.setText("text_wait_...")
        # 设置悬浮提示
        self.listWidget.item(16).setToolTip('<b>Key:</b>text_wait_find_element(等待元素出现指定文字后进行下一步操作)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)'
                                            '<br><b>param:</b>value(指定文字)')

        # 新建item（not_text_wait_find_element）
        item = self.listWidget.item(17)
        item.setText("not_text_...")
        # 设置悬浮提示
        self.listWidget.item(17).setToolTip('<b>Key:</b>not_text_wait_find_element(等待元素指定文字消失后进行下一步操作)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)'
                                            '<br><b>param:</b>value(指定文字)')

        # 新建item（move_to_element）
        item = self.listWidget.item(18)
        item.setText("move_to_ele...")
        # 设置悬浮提示
        self.listWidget.item(18).setToolTip('<b>Key:</b>move_to_element(鼠标悬浮至指定元素)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)')

        # 新建item（Enter）
        item = self.listWidget.item(19)
        item.setText("Enter")
        # 设置悬浮提示
        self.listWidget.item(19).setToolTip('<b>Key:</b>Enter(在指定元素位置进行回车)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)')

        # 新建item（back）
        item = self.listWidget.item(20)
        item.setText("back")
        # 设置悬浮提示
        self.listWidget.item(20).setToolTip('<b>Key:</b>back(浏览器回退)')

        # 新建item（foword）
        item = self.listWidget.item(21)
        item.setText("foword")
        # 设置悬浮提示
        self.listWidget.item(21).setToolTip('<b>Key:</b>foword(浏览器前进)')

        # 新建item（refresh）
        item = self.listWidget.item(22)
        item.setText("refresh")
        # 设置悬浮提示
        self.listWidget.item(22).setToolTip('<b>Key:</b>refresh(刷新浏览器)')

        # 新建item（js_scroll_top）
        item = self.listWidget.item(23)
        item.setText("js_scroll_top")
        # 设置悬浮提示
        self.listWidget.item(23).setToolTip('<b>Key:</b>js_scroll_top(浏览器滚动条滚动至顶部)')

        # 新建item（js_scroll_end）
        item = self.listWidget.item(24)
        item.setText("js_scroll_end")
        # 设置悬浮提示
        self.listWidget.item(24).setToolTip('<b>Key:</b>js_scroll_end(浏览器滚动条滚动至底部)')

        # 新建item（switchToFrame）
        item = self.listWidget.item(25)
        item.setText("switchToFrame")
        # 设置悬浮提示
        self.listWidget.item(25).setToolTip('<b>Key:</b>switchToFrame(切换到frame页面内)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)')

        # 新建item（switchToDefault）
        item = self.listWidget.item(26)
        item.setText("switchToDefault")
        # 设置悬浮提示
        self.listWidget.item(26).setToolTip('<b>Key:</b>switchToDefault(切换到默认的frame页面)')

        # 新建item（asserrSelect）
        item = self.listWidget.item(27)
        item.setText("asserrSelect")
        # 设置悬浮提示
        self.listWidget.item(27).setToolTip('<b>Key:</b>asserrSelect(检查下拉框选择项与指定项是否相同)'
                                            '<br><b>param:</b>location(定位方式)'
                                            '<br><b>param:</b>locator(定位表达式)'
                                            '<br><b>param:</b>value(指定选项)')

        # 添加item点击时间
        self.listWidget.itemDoubleClicked.connect(self.keyWordClick)

        # 用例名称显示
        self.label_17 = QtWidgets.QLabel('用例编号', self)
        self.label_17.setObjectName('case_header_show')

        # 用例标题显示
        self.label_18 = QtWidgets.QLabel('用例标题:', self)
        self.label_18.setObjectName('case_header_show')

        # 用例标题文字显示
        self.label_19 = QtWidgets.QLabel(self)
        self.label_19.setObjectName('case_header_show')

        # 用例测试·
        self.toolButton_2 = QtWidgets.QToolButton(self)
        self.toolButton_2.setText('测试')
        self.toolButton_2.setObjectName('small_button_show')

        # 用例清除
        self.toolButton_3 = QtWidgets.QToolButton(self)
        self.toolButton_3.setText('清除')
        self.toolButton_3.setObjectName('small_button_show')

        # 用例保存按钮
        self.toolButton_18 = QtWidgets.QToolButton(self)
        self.toolButton_18.setText('保存')
        self.toolButton_18.setObjectName('small_button_show')

        # 用例编辑框
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self)
        # 增加例子
        self.plainTextEdit.insertPlainText("步骤描述:进入临时资料\n"
                                           "click('css', '#iframeLeftMune > div.muneList > div:nth-child(2) > ul > a:nth-child(5) > li > span.hideBox')\n"
                                           "步骤描述:点击上传文件按钮\n"
                                           "click('css', '#meetingDatum > div.ifrTableTool.clearfix > a:nth-child(3) > span')\n"
                                           "步骤描述:上传101个文件\n"
                                           "uploadFiles('css', '.webuploader-element-invisible', 'D:\\无纸化测试文件')\n"
                                           "步骤描述:等待文件上传成功\n"
                                           "text_wait_find_element('css', '.fl.badge', '已上传')\n"
                                           "步骤描述:确定上传\n"
                                           "click('css', '.div_btn.bc2.mainBtn')\n"
                                           "步骤描述:等待2s\n"
                                           "sleep('2')\n"
                                           "步骤描述:断言只上传了100个文件\n"
                                           "assertEqule('css', '#ifrPageBox > div > div > div > span.el-pagination__total', '共 100 条')")
        self.plainTextEdit.setObjectName('plaintext')
        self.plainTextEdit.setDisabled(True)

        # log显示框
        self.listWidget_2 = QtWidgets.QListWidget(self)
        self.listWidget_2.setObjectName('log_show')
        # self.listWidget_2.setWordWrap(True)
        self.listWidget_2.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        self.Position()

        self.ButtonBind()

    def Position(self):
        self.listView_3.setGeometry(QtCore.QRect(180, 490, 761, 311))
        self.label.setGeometry(QtCore.QRect(4, 0, 181, 20))
        self.label_2.setGeometry(QtCore.QRect(210, 510, 54, 28))
        self.lineEdit.setGeometry(QtCore.QRect(280, 510, 211, 28))
        self.label_3.setGeometry(QtCore.QRect(220, 570, 54, 28))
        self.comboBox.setGeometry(QtCore.QRect(280, 570, 211, 28))
        self.label_4.setGeometry(QtCore.QRect(210, 630, 54, 28))
        self.comboBox_2.setGeometry(QtCore.QRect(280, 630, 211, 28))
        self.label_5.setGeometry(QtCore.QRect(210, 690, 54, 28))
        self.spinBox.setGeometry(QtCore.QRect(280, 690, 42, 28))
        self.label_6.setGeometry(QtCore.QRect(210, 750, 54, 16))
        self.radioButton.setGeometry(QtCore.QRect(280, 750, 89, 16))
        self.radioButton_2.setGeometry(QtCore.QRect(350, 750, 89, 16))
        self.label_15.setGeometry(QtCore.QRect(280, 770, 181, 20))
        self.label_7.setGeometry(QtCore.QRect(570, 680, 61, 16))
        self.label_8.setGeometry(QtCore.QRect(570, 720, 61, 16))
        self.label_9.setGeometry(QtCore.QRect(570, 760, 54, 12))
        self.line.setGeometry(QtCore.QRect(650, 693, 118, 3))
        self.line_2.setGeometry(QtCore.QRect(650, 733, 118, 3))
        self.line_3.setGeometry(QtCore.QRect(650, 770, 118, 3))
        self.label_10.setGeometry(QtCore.QRect(700, 680, 54, 12))
        self.label_11.setGeometry(QtCore.QRect(700, 720, 54, 12))
        self.label_12.setGeometry(QtCore.QRect(680, 755, 62, 12))
        self.toolButton.setGeometry(QtCore.QRect(570, 510, 81, 31))
        self.toolButton_15.setGeometry(QtCore.QRect(700, 510, 81, 31))
        self.toolButton_16.setGeometry(QtCore.QRect(830, 510, 81, 31))
        self.toolButton_13.setGeometry(QtCore.QRect(570, 600, 81, 31))
        self.toolButton_14.setGeometry(QtCore.QRect(700, 600, 81, 31))
        self.label_17.setGeometry(QtCore.QRect(190, 27, 180, 14))
        self.label_16.setGeometry(QtCore.QRect(580, 560, 241, 16))
        self.label_14.setGeometry(QtCore.QRect(644, 0, 301, 20))
        self.label_13.setGeometry(QtCore.QRect(180, 0, 461, 21))
        self.treeWidget.setGeometry(QtCore.QRect(0, 380, 181, 421))
        self.listWidget.setGeometry(QtCore.QRect(0, 20, 181, 341))
        self.listView.setGeometry(QtCore.QRect(180, 20, 461, 28))
        # self.label_17.setGeometry(QtCore.QRect(190, 24, 150, 12))
        self.toolButton_2.setGeometry(QtCore.QRect(490, 25, 37, 18))
        self.toolButton_3.setGeometry(QtCore.QRect(590, 25, 37, 18))
        self.plainTextEdit.setGeometry(QtCore.QRect(180, 69, 461, 422))
        self.listWidget_2.setGeometry(QtCore.QRect(640, 20, 301, 471))
        self.toolButton_17.setGeometry(QtCore.QRect(830, 600, 81, 31))
        self.listView_4.setGeometry(QtCore.QRect(180, 46, 461, 26))
        self.label_18.setGeometry(QtCore.QRect(190, 54, 120, 14))
        self.label_19.setGeometry(QtCore.QRect(251, 54, 389, 14))
        self.label_20.setGeometry(QtCore.QRect(4, 360, 181, 21))
        self.toolButton_18.setGeometry(QtCore.QRect(540, 25, 37, 18))

    def ButtonBind(self):
        '''
        信号与槽绑定
        :return:
        '''
        # 导入 按钮信号绑定
        self.toolButton.clicked.connect(self.importExcel)

        # 关键字辅助框 设置item可进行内部拖拽
        self.listWidget.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.listWidget.setDragDropMode(self.listWidget.InternalMove)

        # 给树形窗口设置点击信号
        self.treeWidget.itemDoubleClicked.connect(self.treeClick)
        # 给树形窗口添加右键菜单
        self.treeWidget.setContextMenuPolicy(3)
        # 设置菜单
        self.treeWidget.customContextMenuRequested[QtCore.QPoint].connect(self.myListWidgetContext)
        # 绑定用例清除
        self.toolButton_3.clicked.connect(self.clearCase)
        # 绑定单个用例测试
        self.toolButton_2.clicked.connect(self.testCase)
        # 绑定单个用例测试按钮快捷操作
        self.toolButton_2.setShortcut(QtCore.Qt.Key_F5)
        # 设置最大追加数据
        self.plainTextEdit.setMaximumBlockCount(100)
        # 绑定用例保存按钮
        self.toolButton_18.clicked.connect(self.saveCase)
        # CTRL+S绑定保存事件
        self.toolButton_18.setShortcut(QtCore.Qt.CTRL + QtCore.Qt.Key_S)
        # 发射多线程弹窗信号
        self.message = message(self)
        self.message.signal.connect(self.box)
        # 双击弹窗log查看窗
        self.listWidget_2.doubleClicked.connect(self.maxlog)
        # 循环次数选择次数槽函数
        self.spinBox.valueChanged.connect(self.valueSpin)
        # 绑定编写说明槽函数
        self.toolButton_16.clicked.connect(self.explainclick)
        # 绑定报告生成与不生成的槽函数
        self.radioButton.clicked.connect(lambda: self.btnstate(self.radioButton))
        self.radioButton_2.clicked.connect(lambda: self.btnstate(self.radioButton_2))
        # 绑定用例模板导出槽函数
        self.toolButton_15.clicked.connect(self.exportclick)
        # 绑定用例运行槽函数
        self.toolButton_13.clicked.connect(self.runCase)
        # 绑定暂停槽函数
        self.toolButton_14.clicked.connect(self.Suspend)
        # self.toolButton_17.clicked.connect(self.openexcle)
        self.comboBox_2.currentIndexChanged.connect(self.ComboxValue)
        # 结束测试
        self.toolButton_17.clicked.connect(lambda :self.stop_thread(self.runtestthread))

    def openexcle(self):
        # 开启多线程读取excel，防止主程序卡死
        self.importexcelthread = threading.Thread(target=self.importExcelThread)
        self.importexcelthread.setDaemon(True)
        self.importexcelthread.start()

    def explainclick(self):
        """
        测试用例编写说明
        """
        # 选择保存测试用例的路径
        self.explain_path = QFileDialog.getExistingDirectory(self, "请选择编写说明保存路径")
        explainPath = EXCELTEMPLATE_PATH + '测试用例编写说明.doc'
        # 判断该文件是否已存在该路径下
        if self.explain_path != '':
            if os.path.exists(self.explain_path + '\\测试用例编写说明.doc'):
                shutil.copy2(explainPath, self.explain_path + '\\测试用例编写说明(%s).doc' % self.explainum)
                self.explainum = self.explainum + 1
            else:
                self.explainum = 1
                shutil.copy2(explainPath, self.explain_path + '\\测试用例编写说明.doc')
            self.messages = '用例编写说明导出成功！'
            self.message.start()

    def btnstate(self, btn):
        """
        测试报告生成路径
        """
        if btn.text() == "不生成":
            if btn.isChecked() == True:
                # 不生成测试报告，隐藏提示和路径label
                self.label_15.setHidden(True)

        if btn.text() == "生成":
            if btn.isChecked() == True:
                self.label_15.setText('')
                # 生成测试报告，显示提示和路径label
                self.label_15.setHidden(False)
                # 选择报告保存路径
                self.path = QFileDialog.getExistingDirectory(self, "请选择测试报告保存路径")
                self.path = self.path.replace('/', '\\')
                self.label_15.setText(self.path)
                # 取消或关闭选择窗口
                if self.path == '':
                    self.label_15.setHidden(True)
                    # 选中不生成测试报告
                    self.radioButton_2.setChecked(True)

        '''导出用例模板'''

    def exportclick(self):
        try:
            # 选择保存测试用例的路径
            self.file_path = QFileDialog.getExistingDirectory(self, "请选择用例模板保存路径")
            excelPath = EXCELTEMPLATE_PATH + '测试用例模板.xlsx'
            # 判断该文件是否已存在该路径下
            if self.file_path != '':
                if os.path.exists(self.file_path + '\\测试用例模板.xlsx'):
                    shutil.copy2(excelPath, self.file_path + '\\测试用例模板(%s).xlsx' % self.filenum)
                    self.filenum = self.filenum+1
                else:
                    self.filenum = 1
                    shutil.copy2(excelPath, self.file_path + '\\测试用例模板.xlsx')
                self.messages = '用例模板导出成功！'
                self.message.start()
        except Exception as e:
            print(e)

    def maxlog(self):
        tips3 = Tips_3(self)
        # 获取当前log查看窗的日志
        widgetres = []
        # 获取listwidget中条目数
        count = self.listWidget_2.count()
        # 遍历listwidget中的内容
        for i in range(count):
            widgetres.append(self.listWidget_2.item(i).text())
        # 设置子窗口未关闭时无法操作父窗口
        tips3.setWindowModality(QtCore.Qt.ApplicationModal)
        # 运行子窗口
        tips3.show()
        # 插入子窗口的listwidget中
        tips3.listwidget.addItems(widgetres)

    def myListWidgetContext(self, point):
        '''
        右键菜单，添加，删除，修改、
        注：未同步用例
        :param point:
        :return:
        '''
        # 获取treewidget对象
        self.item = self.treeWidget.itemAt(point)
        # 获取右键的节点位置
        rootIndex = self.treeWidget.indexOfTopLevelItem(self.item)
        # 判断右键是否在节点上
        if self.item is not None:
            popMenu = QMenu()
            # 判断右键节点位置（父节点 0  子节点 -1） 父节点右键时只有添加和重命名功能 子节点删除和重命名
            if rootIndex == -1:
                # 设置右键菜单
                insert = popMenu.addAction(QAction(u'删除', self))
                update = popMenu.addAction(QAction(u'重命名', self))
            else:
                delete = popMenu.addAction(QAction(u'添加', self))
                update = popMenu.addAction(QAction(u'重命名', self))
                # 获取右键菜单点击对象
            self.action = popMenu.exec_(self.treeWidget.mapToGlobal(point))
            # 删除功能
            # 判断是否点击右键菜单
            if self.action:
                # 判断右键菜单点击功能
                if self.action.text() == '删除':
                    # 删除增加提示
                    box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("确定删除该用例？"), QMessageBox.NoButton,
                                      self)
                    yr_btn = box.addButton(self.tr("确定"), QMessageBox.YesRole)
                    # 确定按钮绑定回车
                    yr_btn.setShortcut(QtCore.Qt.Key_Return)
                    box.addButton(self.tr("取消"), QMessageBox.NoRole)
                    box.exec_()
                    if box.clickedButton() == yr_btn:
                        self.parent1 = self.item.parent()
                        # 删除excel表格中的数据
                        self.deleteitem()
                    else:
                        pass
                elif self.action.text() == '添加':
                    if '*' in self.treeWidget.headerItem().text(0):
                        QMessageBox.about(self, '提示', '请先保存当前用例再新建')
                    else:
                        # 添加时需要输入名称，新增子窗口Tips_2
                        # 初始化子窗口
                        tips2 = Tips_2(self)
                        tips2.PointTips_1()
                        # 设置子窗口标题
                        tips2.setWindowTitle(self.action.text())
                        # 获取该父节点下的所有子节点
                        ChildList = self.book.getColumnValue(self.item.text(0), testStep_Num)
                        # 获取所有前置条件编号,加入下拉框列表
                        PreList = [x for x in ChildList if 'test_pre' in str(x)]
                        tips2.lineEdit_3.addItems(PreList)
                        # 设置子窗口未关闭时无法操作父窗口
                        tips2.setWindowModality(QtCore.Qt.ApplicationModal)
                        tips2.show()
                        # 发射信号
                        tips2.dialogSignel_2.connect(self.additem)
                elif self.action.text() == '重命名':
                    # 添加时需要输入名称，新增子窗口Tips_2
                    # 初始化子窗口
                    tips2 = Tips_2(self)
                    tips2.PointTips_2()
                    # 设置子窗口标题
                    tips2.setWindowTitle(self.action.text())
                    # 设置子窗口未关闭时无法操作父窗口
                    tips2.setWindowModality(QtCore.Qt.ApplicationModal)
                    tips2.show()
                    # 发射信号
                    tips2.dialogSignel_2.connect(self.updateitem)
                    # 输入框获取右键的item的名称
                    tips2.lineEdit.setText(self.item.text(0))

    def additem(self, mc, caseTitle, preNode):
        '''
        右键父节点，点击添加弹出子窗口，新增给tree子节点
        :param mc:
        :return:
        '''
        # 判断是否重名
        # 获取所有节点的名称
        itemvalue = QtWidgets.QTreeWidgetItemIterator(self.treeWidget)
        self.jdlist = []
        while itemvalue.value():
            items = itemvalue.value()
            columnCount = items.columnCount()
            for i in range(columnCount):
                text = items.text(i)
                if i == columnCount - 1:
                    self.jdlist.append(text)
                else:
                    self.jdlist.append(text)
            itemvalue.__iadd__(1)
        if mc not in self.jdlist:
            try:
                book = ParseExcelXlrd(self.filepath)
                winbook = ParseExcelWin(self.filepath)
                # 新增子节点
                addChild = QtWidgets.QTreeWidgetItem()
                # 设置子节点位置
                addChild.setText(0, mc)
                # 将用例按照格式写入excel表格中
                # 获取最大行数，按照规定列数，将用例编号，用例标题写入
                maxrow = book.wookbook.sheet_by_name(self.item.text(0)).nrows
                # 写入用例编号
                winbook.writeCellValue(self.item.text(0), maxrow + 1, testStep_Num, str(mc))
                # 写入工作表
                winbook.writeCellValue(self.item.text(0), maxrow + 1, testStep_Moudle, self.item.text(0))
                # 写入预置条件
                winbook.writeCellValue(self.item.text(0), maxrow + 1, testStep_Preset, preNode)
                # 写入用例标题
                winbook.writeCellValue(self.item.text(0), maxrow + 1, testStep_Title, caseTitle)
                # 设置单元格外边框
                winbook.borderAround(self.item.text(0), maxrow + 1, testStep_Num, maxrow+1, testStep_Expect)
                # self.label_16.setText('正在更新用例文件，请稍等...')
                winbook.save()
                # 在右键父节点位置添加子节点
                self.item.addChild(addChild)
                # 清除用例编辑框
                self.plainTextEdit.clear()
                # 重置用例编号框
                self.label_17.setText('')
                # 重置用例标题
                self.label_19.setText('')
                # 设置树形标签选中状态
                self.treeWidget.setCurrentItem(addChild)
            # # 禁用导入文件按钮
            # self.toolButton.setDisabled(True)
            # # 设置读取提示颜色
            # self.label_16.setStyleSheet("QLabel{color:red}")
            # # 更新导入文件
            # p = threading.Thread(target=self.importExcelThread)
            # p.setDaemon(True)
            # p.start()
            except PermissionError:
                # 重名提示
                box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("请关闭用例后重试"), QMessageBox.NoButton,
                                  self)
                yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
                # 确定按钮绑定回车
                yr_btn.setShortcut(QtCore.Qt.Key_Return)
                box.exec_()
            except Exception as e:
                print(e)
                box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("新增失败，请重新导入用例后重试"),
                                  QMessageBox.NoButton,
                                  self)
                yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
                # 确定按钮绑定回车
                yr_btn.setShortcut(QtCore.Qt.Key_Return)
                box.exec_()
                # 禁用导入文件按钮
                self.toolButton.setDisabled(False)
                # 设置读取提示颜色
                self.label_16.setText(os.path.basename(self.newfile))
                self.label_16.setStyleSheet("QLabel{color:white}")
            finally:
                winbook.close()
        else:
            # 重名提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("%s用例编号已存在，请重新输入" % mc),
                              QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.exec_()
            if box.clickedButton() == yr_btn:
                # 初始化子窗口
                tips2 = Tips_2(self)
                tips2.PointTips_1()
                # 设置子窗口标题
                tips2.setWindowTitle(self.action.text())
                # 设置子窗口未关闭时无法操作父窗口
                tips2.setWindowModality(QtCore.Qt.ApplicationModal)
                tips2.show()
                # 发射信号
                tips2.dialogSignel_2.connect(self.additem)
            else:
                pass

    def updateitem(self, caseNum1, caseNum2):
        '''
        右键父节点或者子节点，对用例或工作表进行重命名
        :param mc:
        :return:
        '''
        # 获取素有节点的名称
        itemvalue = QtWidgets.QTreeWidgetItemIterator(self.treeWidget)
        # 用于储存所有名称
        self.jdlist = []
        while itemvalue.value():
            items = itemvalue.value()
            columnCount = items.columnCount()
            for i in range(columnCount):
                text = items.text(i)
                if i == columnCount - 1:
                    self.jdlist.append(text)
                else:
                    self.jdlist.append(text)
            itemvalue.__iadd__(1)
        # 判断是否重名
        if caseNum2 not in self.jdlist:
            # 获取右键节点的父节点的值和类型
            Pnode = self.item.parent()
            PnodeIndex = self.treeWidget.indexOfTopLevelItem(Pnode)
            # 判断修改的是工作表还是单元格的数据
            try:
                if PnodeIndex == -1:
                    # 修改工作表名称
                    self.parseexcel.wb[self.item.text(0)].title = caseNum2
                    self.parseexcel.wb.save(self.filepath)
                    self.parseexcel.wb.close()
                else:
                    PnodeText = Pnode.text(0)
                    # 获取需要修改的文件编号在集合中的下标
                    # print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
                    CaseNodeList = self.book.getColumnValue(PnodeText, testStep_Num)
                    UpdateValueIndex = CaseNodeList.index(self.item.text(0))
                    # 修改数据表中的值
                    self.parseexcel.wb[PnodeText].cell(UpdateValueIndex + 1, testStep_Num, caseNum2)
                    # 修改树形标签名称显示
                self.item.setText(0, caseNum2)
            except PermissionError:
                # 重名提示
                box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("请关闭用例后重试"), QMessageBox.NoButton,
                                  self)
                yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
                # 确定按钮绑定回车
                yr_btn.setShortcut(QtCore.Qt.Key_Return)
                box.exec_()
            except Exception as e:
                box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("重命名失败，请重新导入用例后重试"),
                                  QMessageBox.NoButton,
                                  self)
                yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
                # 确定按钮绑定回车
                yr_btn.setShortcut(QtCore.Qt.Key_Return)
                box.exec_()
                # 禁用导入文件按钮
                self.toolButton.setDisabled(False)
                # 设置读取提示颜色
                self.label_16.setText(os.path.basename(self.newfile))
                self.label_16.setStyleSheet("QLabel{color:white}")

        else:
            # 重名提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("%s用例编号已存在，请重新输入" % caseNum2),
                              QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.exec_()
            if box.clickedButton() == yr_btn:
                # 初始化子窗口
                tips2 = Tips_2(self)
                tips2.PointTips_2()
                # 设置子窗口标题
                tips2.setWindowTitle(self.action.text())
                # 设置子窗口未关闭时无法操作父窗口
                tips2.setWindowModality(QtCore.Qt.ApplicationModal)
                tips2.show()
                # 发射信号
                tips2.dialogSignel_2.connect(self.updateitem)
                # 输入框获取右键的item的名称
                tips2.lineEdit.setText(self.item.text(0))
            else:
                pass

    def deleteitem(self):
        '''
        删除excle表格中的数据
        '''
        # 开启多线程写入
        deleteitemthread = threading.Thread(target=self.deleteitemThread)
        deleteitemthread.setDaemon(True)
        deleteitemthread.start()

    def deleteitemThread(self):
        winbook = ParseExcelWin(self.newfile)
        book = ParseExcelXlrd(self.newfile)
        # 获取被删除的case名称和父元素名称
        deleteCaseId = self.item.text(0)
        PnodeText = self.parent1.text(0)
        # 获取excel表格中符合被删除的用例集合
        deleteCaseList = book.getMergeColumnValue(PnodeText, testStep_Num)
        # 找到集合中的用例下表
        deleteCaseListIndex = []
        if deleteCaseList is not None:
            deleteCaseListIndex = [i for i in range(len(deleteCaseList)) if str(deleteCaseList[i]) == str(deleteCaseId)]
        # 循环所有下标，并删除指定行
        if len(deleteCaseListIndex) != 0:
            try:
                winbook.deleteRows(PnodeText, deleteCaseListIndex[0]+2, len(deleteCaseListIndex))
                # for i in deleteCaseListIndex:
                #     self.parseexcel.wb[PnodeText].delete_rows(i + 2)
                # self.parseexcel.wb.save(self.filepath)
                # 删除选中的treewidget对象
                self.parent1.removeChild(self.item)
                winbook.save()
                winbook.close()
            except Exception as e:
                self.messages = '删除失败，请重试'
                self.message.start()
                # # 禁用导入文件按钮
                # self.toolButton.setDisabled(False)
                # # 设置读取提示颜色
                # self.label_16.setText(os.path.basename(self.newfile))
                # self.label_16.setStyleSheet("QLabel{color:white}")
        else:
            self.messages =  '删除失败，请重试'
            self.message.start()

    def keyWordClick(self, item):
        '''
        打开关键字子窗口
        :param item:
        :return:
        '''
        if self.import_num == 0:
            # 重名提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("请导入用例文件后重试"), QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.exec_()
            return
        # 获取item文字，赋值子窗口标题
        self.itemText = item.text()
        # 初始化子窗口
        tips = Tips_1(self)
        # 分类子窗口参数类别（4,3,2,1）
        if self.itemText == 'inputValue' or self.itemText == 'assertEqule' or self.itemText == 'assertLen' or self.itemText == 'uploadFile' \
                or self.itemText == 'uploadFiles' or self.itemText == 'text_wait_...' or \
                self.itemText == 'not_text_...' or self.itemText == 'asserrSelect' or self.itemText == 'assertElement' :
            tips.PointTips_1()
        elif self.itemText == 'click' or self.itemText == 'clear' or self.itemText == 'wait_find_...' \
                or self.itemText == 'not_wait_...' or self.itemText == 'move_to_ele...' or self.itemText == 'Enter' \
                or self.itemText == 'switchToFrame':
            tips.PointTips_2()
        elif self.itemText == 'openBrowsers' or self.itemText == 'getUrl' or self.itemText == 'sleep' or self.itemText == 'assertUrl' \
                or self.itemText == 'assertTitle':
            tips.PointTips_3()
        elif self.itemText == 'back' or self.itemText == 'foword' or self.itemText == 'refresh' or self.itemText == 'js_scroll_top' \
                or self.itemText == 'js_scroll_end' or self.itemText == 'switchToDefault':
            tips.PointTips_4()
        elif self.itemText == 'if':
            tips.PointTips_5()
        # 设置子窗口标题
        tips.setWindowTitle(self.itemText)
        # 设置子窗口未关闭时无法操作父窗口
        tips.setWindowModality(QtCore.Qt.ApplicationModal)
        # 运行子窗口
        tips.show()
        # 发射子窗口数据
        tips.dialogSignel.connect(self.pltext)

    def pltext(self, bz, by, location, value):
        if self.plainTextEdit.toPlainText() != '':
            self.plainTextEdit.insertPlainText('\n')
        # 赋值编辑框, 四种情况
        if self.itemText == 'if':
            self.plainTextEdit.insertPlainText(
                "步骤描述:%s\n%s('%s', '%s', '%s')" % (bz, self.itemText, location, by, value))
            self.plainTextEdit.insertPlainText(
                "\n步骤描述:结束判断\nbreak()")
            return
        if by and location and value:
            self.plainTextEdit.insertPlainText(
                "步骤描述:%s\n%s('%s', '%s', '%s')" % (bz, self.itemText, by, location, value))
            return
        elif by and location:
            self.plainTextEdit.insertPlainText(
                "步骤描述:%s\n%s('%s', '%s')" % (bz, self.itemText, by, location))
            return
        elif value:
            self.plainTextEdit.insertPlainText(
                "步骤描述:%s\n%s('%s')" % (bz, self.itemText, value))
            return
        else:
            self.plainTextEdit.insertPlainText(
                "步骤描述:%s\n%s()" % (bz, self.itemText))
            return

    '''导入测试用例'''
    def importExcel(self):
        try:
            # 导入文件名称
            self.filepath, filetype = QFileDialog.getOpenFileName(self, "请导入测试用例", '', 'Excel files(*.xlsx)')
            # self.filepath, filetype = QFileDialog.getExistingDirectory(self, "请导入测试用例", './',)
            # self.filepath, filetype = QFileDialog.getOpenFileNames(self, "请导入测试用例", '', 'Excel files(*.xlsx)')
            if self.filepath == '':
                if self.label_16.text() == '请选择excel.xlsx格式表格导入':
                    self.filepath = self.newfile
                    self.label_16.setText('请选择excel.xlsx格式表格导入')
            else:
                self.newfile = self.filepath
                self.newfile = self.filepath.replace('/', '\\')
                # 增加读取提示，防止用户勿操作
                self.label_16.setText('正在读取用例文件，请勿进行其他操作...')
                self.label_11.setText('0')
                # 设置读取提示颜色
                self.label_16.setStyleSheet("QLabel{color:red}")
                # 开启多线程读取excel，防止主程序卡死
                self.importexcelthread = threading.Thread(target=self.importExcelThread)
                self.importexcelthread.setDaemon(True)
                self.importexcelthread.start()
                # 不是第一次导入用例时，先清空编辑框
                if self.import_num != 0:
                    # 清除用例编辑框
                    self.plainTextEdit.clear()
                    # 重置用例编号框
                    self.label_17.setText('用例编号')
                    # 重置用例标题
                    self.label_19.setText('')
                self.import_num += 1
        except Exception:
            pass

    def importExcelThread(self):
        '''
        开启多线程导入文件，防止界面卡顿
        :return:
        '''
        try:
            # 禁用导入文件按钮
            self.toolButton.setDisabled(True)
            # 设置导入按钮为灰色
            self.toolButton.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
            # 禁用导入文件按钮
            self.toolButton_13.setDisabled(True)
            # 设置导入按钮为灰色
            self.toolButton_13.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
            # 禁用导入文件按钮
            self.toolButton_2.setDisabled(True)
            # 设置导入按钮为灰色
            self.toolButton_2.setStyleSheet(
                "QToolButton{font:9pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
            # 打开excel
            self.book = ParseExcelXlrd(self.newfile)
            # 获取工作表标签
            self.sheets = self.book.wookbook.sheet_names()

            # 获取第一个工作表的第一行数据
            CaseBook = ParseExcelXlrd(self.filepath).getRowValue(0, 2)
            CaseBook = list(filter(None, CaseBook))
            CaseBookTemplate = ['序号', '用例编号', '用例工作表', '用例标题', '预期结果', '是否执行', '执行结束时间', '执行结果1', '执行结果2', '执行结果3', '执行结果4', '执行结果5']
            if CaseBook != CaseBookTemplate:
                self.messages = '导入失败，请检查用例'
                self.message.start()
                self.label_16.setText('请选择exlce.xlsx格式表格导入')
                self.label_16.setStyleSheet("QLabel{color:white}")
                return
            ''' treeWidget列表更新'''
            # 清空treewidget列表
            self.treeWidget.clear()
            # 通过 工作表 -- 用例编号进行树形排列
            # 循环所有工作表，根据工作表创建父节点
            for i, v in enumerate(self.sheets):
                # 获取工作表的第一行数据
                CaseSheet = ParseExcelXlrd(self.filepath).getRowValue(v, 1)
                CaseSheet = list(filter(None, CaseSheet))
                CaseSheetTemplate = ['用例编号', '工作表', '预置条件编号', '用例标题', '预期结果', '测试步骤描述', '关键字', '操作元素的定位方式', '操作元素的定位表达式', '操作值']
                if i != 0 and set(CaseSheetTemplate) < set(CaseSheet):
                    self.item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
                    self.treeWidget.topLevelItem(i - 1).setText(0, v)
                else:
                    continue
                # 每读取一个工作表，读取所有的用例编号，并创建子节点
                for j, d in enumerate(list(filter(None, self.book.getColumnValue(v, testStep_Num)))):
                    if j != 0:
                        item_1 = QtWidgets.QTreeWidgetItem(self.item_0)
                        if type(d) is float:
                            d = str(int(d))
                        self.treeWidget.topLevelItem(i - 1).child(j - 1).setText(0, d)
                    else:
                        continue
            # 计算总用例数量
            CaseSum = self.book.getColumnValue(self.sheets[0], testCase_Isimplement)
            self.CaseSumIndex = [i for i in range(len(CaseSum)) if CaseSum[i].lower() == 'y']
            self.label_10.setText(str(len(self.CaseSumIndex)))
            # 模块下拉框赋值
            self.comboBox_2.clear()
            self.comboBox_2.addItem('全部')
            self.comboBox_2.setCurrentText('全部')
            if len(self.sheets) > 1:
                sheetlist = self.sheets
                del sheetlist[0]
                self.comboBox_2.addItems(sheetlist)
            else:
                pass
            self.parseexcel = ParseExcel(self.newfile)
            # 显示文件名称
            self.label_16.setText(os.path.basename(self.newfile))
            self.label_16.setStyleSheet("QLabel{color:white}")
        except Exception as e:
            print(e)
            # 绑定弹窗
            self.messages = '导入失败，请检查用例'
            self.message.start()
            self.label_16.setText('请选择exlce.xlsx格式表格导入')
            self.label_16.setStyleSheet("QLabel{color:white}")
        finally:
            # 启用导入按钮
            self.toolButton.setDisabled(False)
            # 设置导入按钮为灰色
            self.toolButton.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
            # 启用导入按钮
            self.toolButton_13.setDisabled(False)
            # 设置导入按钮为灰色
            self.toolButton_13.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
            # 启用导入按钮
            self.toolButton_2.setDisabled(False)
            # 设置导入按钮为灰色
            self.toolButton_2.setStyleSheet(
                "QToolButton{font:9pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
            bgimg3 = RESOURSE_PATH + 'bgimg3.png'
            bgimg3 = bgimg3.replace('\\', '/')
            self.plainTextEdit.clear()
            self.plainTextEdit.setStyleSheet("QPlainTextEdit{background-image: url(%s);color:#FFFFFF;}" % bgimg3)
            self.plainTextEdit.setEnabled(True)

    def box(self):
        # 重名提示
        box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr(self.messages), QMessageBox.NoButton, self)
        yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
        # 确定按钮绑定回车
        yr_btn.setShortcut(QtCore.Qt.Key_Return)
        box.exec_()

    def treeClick(self, item):
        '''
        根据父节点，子节点，获取相应的excel信息
        :return:
        '''
        if self.label_16.text() == '正在读取用例文件，请勿进行其他操作...':
            self.messages = '请用例导入完成后重试！'
            self.message.start()
            return
        # 获取父节点索引
        F_node = item.parent()
        F_index = self.treeWidget.indexOfTopLevelItem(F_node)
        # 获取父节点内容
        # 判断点击层级，0为子节点,-1为父节点，点击父节点时只有收缩功能  or self.plainTextEdit.isSignalConnected() is True
        if F_index != -1:
            # 清除用例编辑框
            self.plainTextEdit.clear()
            # 重置用例编号框
            self.label_17.setText('用例编号')
            # 重置用例标题
            self.label_19.setText('')
            self.treeWidget.setDisabled(True)
            # 禁用编辑框
            self.plainTextEdit.setDisabled(True)
            # 禁用导入按钮
            self.toolButton.setDisabled(True)
            self.F_text = self.treeWidget.topLevelItem(F_index).text(0)
            # 获取子节点内容
            self.S_text = self.treeWidget.currentItem().text(0)
            # 读取Excel表中对应的工作表内容
            # 读取Excel表中对应的工作表内容
            self.readexcel = threading.Thread(target=self.readExcel)
            self.readexcel.setDaemon(True)
            if self.readexcel.isAlive() is False:
                self.readexcel.start()

    def readExcel(self):
        '''
        开启多线程，防止读取excel文件卡顿
        :return:
        '''
        try:
            # 重新读取用例文件
            book = ParseExcelXlrd(self.newfile)
            # 记录用例标题是否写入
            step_title_num = 0
            # 获取所有的用例编号
            step_num = book.getMergeColumnValue(self.F_text, testStep_Num)
            # 通过值找到所有符合用例编号的下表
            CaseIndex = []
            if step_num is not None:
                CaseIndex = [i for i in range(len(step_num)) if str(step_num[i]) == str(self.S_text)]
                self.label_17.setText('正在读取内容，请稍等...')
            # 循环所有下标，并获取其他数据
            plainList = []
            for i in CaseIndex:
                # 获取用例标题，输入至编辑框内, 只输入一次
                if step_title_num == 0:
                    self.step_title = book.getCellValue(self.F_text, i + 2, testStep_Title)
                    self.step_title.replace('\n', '')
                    self.label_19.setText(str(self.step_title))
                    step_title_num += 1
                # 获取用例步骤
                self.step_describe = book.getCellValue(self.F_text, i + 2, testStep_Describe)
                # 获取关键字
                self.step_keyword = book.getCellValue(self.F_text, i + 2, testStep_KeyWord)
                # 获取定位方式
                self.step_loaction = book.getCellValue(self.F_text, i + 2, testStep_Location)
                # 获取定位表达式
                self.step_locator = book.getCellValue(self.F_text, i + 2, testStep_Locator)
                # 获取输入值
                self.step_value = book.getCellValue(self.F_text, i + 2, testStep_Value)
                # 写入编辑框, 分四种情况
                # 赋值编辑框, 四种情况
                fun = ''
                if self.step_keyword and self.step_loaction and self.step_locator and self.step_value:
                    fun = "%s('%s', '%s', '%s')" % (self.step_keyword, self.step_loaction,
                                                    self.step_locator, self.step_value)
                    plainList.append('步骤描述:%s'%self.step_describe)
                    plainList.append(fun)
                elif self.step_keyword and self.step_loaction and self.step_locator:
                    fun = "%s('%s', '%s')" % (self.step_keyword, self.step_loaction,
                                              self.step_locator)
                    plainList.append('步骤描述:%s'%self.step_describe)
                    plainList.append(fun)
                elif self.step_keyword and self.step_value:
                    fun = "%s('%s')" % (self.step_keyword, self.step_value)
                    plainList.append('步骤描述:%s'%self.step_describe)
                    plainList.append(fun)
                elif self.step_keyword:
                    fun = "%s()" % self.step_keyword
                    plainList.append('步骤描述:%s'%self.step_describe)
                    plainList.append(fun)
            plainvalue = "\n".join(str(i) for i in plainList)
            self.plainTextEdit.appendPlainText(plainvalue)
            pretest = book.getCellValue(self.F_text, CaseIndex[0]+2, testStep_Preset)
            if pretest != '' and pretest is not None:
                # 设置编辑框显示
                self.label_17.setText('%s+%s'%(pretest, self.S_text))
            else:
                self.label_17.setText(self.S_text)
            # 启用树形表
            self.treeWidget.setDisabled(False)
            # 启用编辑框
            self.plainTextEdit.setDisabled(False)
            # 启用导入按钮
            self.toolButton.setDisabled(False)
        except Exception as e:
            print(e)

    def clearCase(self):
        '''
        清空用例编辑框
        :return:
        '''
        if self.plainTextEdit.toPlainText() != '':
            # 删除增加提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("是否清除用例，清除后无法撤回"), QMessageBox.NoButton,
                              self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.YesRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.addButton(self.tr("取消"), QMessageBox.NoRole)
            box.exec_()
            if box.clickedButton() == yr_btn:
                # 清除用例编辑框
                self.plainTextEdit.clear()
                # 重置用例编号框
                self.label_17.setText('用例编号')
                # 重置用例标题
                self.label_19.setText('')
                # 清空log框
                self.listWidget_2.clear()
            else:
                pass

    def testCase(self):
        '''
        测试单条用例
        :return:
        '''
        if self.import_num == 0:
            # 重名提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("请导入用例文件后重试"), QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.exec_()
            return
        elif self.lineEdit.text() == '':
            self.messages = '测试地址不能为空！'
            self.message.start()
        elif re.match(r"^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$", self.lineEdit.text()) is None and re.match(r'[^\s]*[.com|.cn]', self.lineEdit.text()) is None:
            self.messages = '请输入正确格式的测试地址！'
            self.message.start()
        elif self.plainTextEdit.toPlainText() == '':
            # 重名提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("编辑框为空"), QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.exec_()
            return
        elif self.label_17.text() == '用例编号':
            # 重名提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("未选择用例，请重试"), QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.exec_()
            return
        else:
            try:
                self.parseexcel.writeCellValue(self.sheets[0], 50, 50, '测试')
            except PermissionError:
                self.messages = "请先关闭用例文件，再运行测试用例"
                self.message.start()
                return
            except Exception as e:
                self.messages = '运行错误，请重新运行或重启软件'
                self.message.start()
                return
            if self.comboBox.currentText() == 'Google Chrome':
                VersionYaml = ParseYaml().ReadParameter('Version')
                if VersionYaml == "":
                    # 获取浏览器的版本号
                    driver = webdriver.Chrome(
                        executable_path=DRIVERS_PATH + 'chrome\\' + '70.0.3538.97\\chromedriver.exe')
                    driver.get('chrome://version/')
                    version = driver.find_element_by_css_selector('#version > span:nth-child(1)').text
                    driver.quit()
                    VersionList = ['70', '71', '72', '73', '74', '75', '76', '77', '78']
                    for i in VersionList:
                        if i == version[:2]:
                            self.writeyaml.Write_Yaml_Updata('Version', version[:2])
                        elif int(version[:2]) < 70 or int(version[:2]) > 78:
                            self.messages = "浏览器版本不符合，请更新浏览器版本号"
                            self.message.starts()
                            return
                else:
                    pass
            self.writeyaml.Write_Yaml_Updata('IP', self.lineEdit.text())
            self.writeyaml.Write_Yaml_Updata('Browser', self.comboBox.currentText())
            # 禁用导入,启动和测试按钮
            self.toolButton_2.setDisabled(True)
            self.toolButton_13.setDisabled(True)
            self.toolButton.setDisabled(True)
            self.toolButton_2.setStyleSheet(
                "QToolButton{font:9pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
            """点击事件"""
            self.toolButton_13.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
            """点击事件"""
            self.toolButton.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
            """点击事件"""
            self.listWidget_2.clear()
            # 清空log框
            self.listWidget_2.setGeometry(QtCore.QRect(180, 20, 762, 471))
            # log框置顶
            self.listWidget_2.raise_()
            # 初始化定时器
            self.timercase = TimerCase()
            self.testcasethread = threading.Thread(target=TestOneCase(self.newfile, self.F_text, self.S_text).TestCase)
            self.testcasethread.setDaemon(True)
            self.testcasethread.start()
            self.timer3.timeout.connect(self.OneCaseOutLog)
            self.timer3.start(500)
            self.timer4.timeout.connect(self.Onedisplaycase)
            self.timer4.start(1000)

    def Onedisplaycase(self):
        """
        单条用例运行时间显示
        """
        if self.testcasethread.isAlive():
            self.label_12.setText(self.timercase.Timer())
        else:
            self.toolButton_2.setDisabled(False)
            self.toolButton_13.setDisabled(False)
            self.toolButton.setDisabled(False)
            self.toolButton_2.setStyleSheet(
                "QToolButton{font:9pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
            """点击事件"""
            self.toolButton_13.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
            """点击事件"""
            self.toolButton.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
            """点击事件"""
            self.listWidget_2.setGeometry(QtCore.QRect(640, 20, 301, 471))
            self.timer4.disconnect()
            self.timer4.stop()

    def saveCase(self):
        """
        保存用例
        :return:
        """
        if self.import_num == 0:
            # 重名提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("请导入用例文件后重试"), QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.exec_()
            return
        elif self.plainTextEdit.toPlainText() == '':
            # 重名提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("编辑框为空"), QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.exec_()
            return
        elif self.label_17.text() == '用例编号':
            # 重名提示
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("未选择用例，请重试"), QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.NoRole)
            # 确定按钮绑定回车
            yr_btn.setShortcut(QtCore.Qt.Key_Return)
            box.exec_()
            return
        else:
                # 开启多线程写入
                keywordthread = threading.Thread(target=self.keywordThread)
                keywordthread.setDaemon(True)
                keywordthread.start()

    def keywordThread(self):
        try:
            oldbook = ParseExcelXlrd(self.newfile)
            winbook = ParseExcelWin(self.newfile)
            # 获取所有excel用例编号
            caseList = oldbook.getMergeColumnValue(self.F_text, testStep_Num)
            caseListIndex = [i for i in range(len(caseList)) if str(caseList[i]) == self.S_text]
            # 获取预置条件编号，用例标题，预期结果，用于合并单元格时赋值
            # 预置条件编号
            presetNum = oldbook.getCellValue(self.F_text, caseListIndex[0] + 2, testStep_Preset)
            # 用例标题
            title = oldbook.getCellValue(self.F_text, caseListIndex[0] + 2, testStep_Title)
            # 预期结果
            expect = oldbook.getCellValue(self.F_text, caseListIndex[0] + 2, testStep_Expect)
            # 获取用例编辑框的每行数据
            self.plainTextList = self.plainTextEdit.toPlainText().split('\n')
            # 过滤步骤描述
            keywordList = [s for s in range(len(self.plainTextList)) if '步骤描述' not in self.plainTextList[s][0:4]]
            # 判断用例编辑框关键字数据和excel表格中的数据是否对应，用例编辑框关键字数据条数大于excle用例条数则插入行，否则删除行
            if len(keywordList) > len(caseListIndex):
                insert_num = len(keywordList) - len(caseListIndex)
                winbook.insertRows(self.F_text, caseListIndex[-1] + 3, insert_num)
                # 合并用例编号单元格
                winbook.mergecells(self.F_text, caseListIndex[0] + 2, testStep_Num, caseListIndex[-1] + 2 + insert_num, testStep_Num)
                winbook.writeCellValue(self.F_text, caseListIndex[0] + 2, testStep_Num, self.S_text)
                # 合并工作表单元格
                winbook.mergecells(self.F_text, caseListIndex[0] + 2, testStep_Moudle, caseListIndex[-1] + 2 + insert_num, testStep_Moudle)
                winbook.writeCellValue(self.F_text, caseListIndex[0] + 2, testStep_Moudle, self.F_text)
                # 合并预置条件单元格
                winbook.mergecells(self.F_text, caseListIndex[0] + 2, testStep_Preset, caseListIndex[-1] + 2 + insert_num, testStep_Preset)
                winbook.writeCellValue(self.F_text, caseListIndex[0] + 2, testStep_Preset, presetNum)
                # 合并用例标题单元格
                winbook.mergecells(self.F_text, caseListIndex[0] + 2, testStep_Title, caseListIndex[-1] + 2 + insert_num, testStep_Title)
                winbook.writeCellValue(self.F_text, caseListIndex[0] + 2, testStep_Title, title)
                # 合并预期结果单元格
                winbook.mergecells(self.F_text, caseListIndex[0] + 2, testStep_Expect, caseListIndex[-1] + 2 + insert_num, testStep_Expect)
                winbook.writeCellValue(self.F_text, caseListIndex[0] + 2, testStep_Expect, expect)
                winbook.borderAround(self.F_text, caseListIndex[0]+2, testStep_Num, caseListIndex[-1]+2+insert_num, testStep_Picture)
            elif len(keywordList) < len(caseListIndex):
                # 删除行
                delete_num = len(caseListIndex) - len(keywordList)
                winbook.deleteRows(self.F_text, caseListIndex[-delete_num] + 2, delete_num)
            else:
                pass
            winbook.save()
            # 更新导入文件
            newbook = ParseExcelXlrd(self.filepath)
            # 获取所有excel用例编号
            caseList1 = newbook.getMergeColumnValue(self.F_text, testStep_Num)
            # 获取对应的用例编号的位置
            caseListIndex1 = [i for i in range(len(caseList1)) if str(caseList1[i]) == self.S_text]
            # 步骤描述内容获取
            describeList = [v for i, v in enumerate(self.plainTextList) if self.plainTextList[i][0:4] == '步骤描述']
            # 关键字内容获取
            keywordValueList = [v for i, v in enumerate(self.plainTextList) if v.split('(')[0] in self.keyword]
            # 写入步骤描述
            for i, v in enumerate(describeList):
                # 截取步骤描述
                describe = v[5:]
                # 写入步骤描述
                winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_Describe, describe)
            for i, v in enumerate(keywordValueList):
                # 关键字
                keyword = v.split('(', 1)[0]
                # 参数
                parameter = v.split('(', 1)[1]
                # 去除最后一个字符
                parameter = parameter[:-1]
                # 截取，前后的字符串
                parameterList = []
                if keyword == 'if' and '(' in parameter:
                    location = parameter.split(')', 1)[0]+')'
                    location = location[1:]
                    locatorValue = parameter.split(')', 1)[1]
                    locatorValue = locatorValue[2:]
                    locatorValue = locatorValue.replace("'", '')
                    locatorValue = locatorValue.strip()
                    parameterList = locatorValue.split(',')
                    parameterList.insert(0, location)
                else:
                    # 去除 '
                    parameter = parameter.replace("'", '')
                    if parameter != '':
                        parameterList = parameter.split(',', 2)
                if len(parameterList) == 0:
                    winbook.clearKeyWordValue(self.F_text, caseListIndex1[i] + 2)
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_KeyWord,
                                                         keyword)
                elif len(parameterList) == 1:
                    winbook.clearKeyWordValue(self.F_text, caseListIndex1[i] + 2)
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_KeyWord,
                                                         keyword)
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_Value,
                                                         parameterList[0].strip())
                elif len(parameterList) == 2:
                    winbook.clearKeyWordValue(self.F_text, caseListIndex1[i] + 2)
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_KeyWord,
                                                         keyword)
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_Location,
                                                         parameterList[0].strip())
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_Locator,
                                                         parameterList[1].strip())
                elif len(parameterList) == 3:
                    winbook.clearKeyWordValue(self.F_text, caseListIndex1[i] + 2)
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_KeyWord,
                                                         keyword)
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_Location,
                                                         parameterList[0].strip())
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_Locator,
                                                         parameterList[1].strip())
                    winbook.writeCellValue(self.F_text, caseListIndex1[i] + 2, testStep_Value,
                                                         parameterList[2].strip())
                    # if，whele for 等三个函数最后完成
            winbook.save()
            self.messages = '保存成功'
            self.message.start()
        except PermissionError:
            self.messages = '请关闭用例后重试'
            self.message.start()
        except Exception as e:
            print(e)
            self.messages = '保存失败，请重试'
            self.message.start()
            # 禁用导入文件按钮
            self.toolButton.setDisabled(False)
            # 设置读取提示颜色
            self.label_16.setText(os.path.basename(self.newfile))
            self.label_16.setStyleSheet("QLabel{color:white}")
        finally:
            winbook.close()

    def OneCaseOutLog(self):
        """
        持续输出日志
        :return:
        """
        try:
            base_dir = CASELOGS_PATH
            l = os.listdir(base_dir)
            l.sort(key=lambda fn: os.path.getmtime(base_dir + fn)
            if not os.path.isdir(base_dir + fn) else 0)
            if 'logger' not in l[-1]:
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
                lines = dat_file.readlines()
                if lines:
                    newlines = []
                    for i in lines:
                        i = i.decode('utf-8')
                        i = i.strip()
                        newlines.append(i)
                    dat_file.close()
                    # 获取listwidget中条目数
                    count = self.listWidget_2.count()
                    # 遍历listwidget中的内容
                    for i in range(count):
                        self.OneCaseLogList.append(self.listWidget_2.item(i).text())
                    # 获取log和listwidget中未重复的数据
                    notsetList = sorted(set(newlines)-set(self.OneCaseLogList), key=newlines.index)
                    if notsetList:
                        self.listWidget_2.addItems(notsetList)
                        self.listWidget_2.scrollToBottom()
                    if self.toolButton_13.isEnabled():
                        self.timer2.stop()
                        self.timer4.stop()
        except Exception as e:
            print(e)

    def OutPut(self):
        """
        持续输出日志
        :return:
        """
        try:
            base_dir = LOGS_PATH
            l = os.listdir(base_dir)
            l.sort(key=lambda fn: os.path.getmtime(base_dir + fn)
            if not os.path.isdir(base_dir + fn) else 0)
            if 'logger' not in l[-1]:
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
                lines = dat_file.readlines()
                if lines:
                    last_line = lines[-1].strip()
                dat_file.close()
                if last_line == '' or last_line is None:
                    print('')
                else:
                    widgetres = []
                    # 获取listwidget中条目数
                    count = self.listWidget_2.count()
                    # 遍历listwidget中的内容
                    for i in range(count):
                        widgetres.append(self.listWidget_2.item(i).text())
                    if last_line.decode('utf-8') not in widgetres:
                        self.listWidget_2.insertItem(len(widgetres), last_line.decode('utf-8'))
                        self.listWidget_2.scrollToBottom()
                    if self.displaycasethread.isAlive() is False and self.toolButton_13.isEnabled() is False:
                        self.displaycasethread = threading.Thread(target=self.displaycase)
                        self.displaycasethread.start()
        except Exception as e:
            print(e)

    def valueSpin(self):
        """
        用于计算用例总数量
        """
        if self.import_num != 0:
            if self.comboBox_2.currentText() != '全部':
                loop = int(self.spinBox.value())
                self.loopcase = len(self.MoudleSum)*loop
                self.label_10.setText(str(self.loopcase))
            else:
                loop = int(self.spinBox.value())
                self.loopcase = len(self.CaseSumIndex)*loop
                self.label_10.setText(str(self.loopcase))

    def runCase(self):
        """
        测试用例
        """
        try:
            if self.lineEdit.text() == '':
                self.messages = '测试地址不能为空！'
                self.message.start()
            elif re.match(r"^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$", self.lineEdit.text()) is None and re.match(r'[^\s]*[.com|.cn]', self.lineEdit.text()) is None:
                self.messages = '请输入正确格式的测试地址！'
                self.message.start()
            elif self.newfile == '':
                self.messages = '请导入测试用例！'
                self.message.start()
            else:
                try:
                    self.parseexcel.writeCellValue(self.sheets[0], 50, 50, '测试')
                except PermissionError:
                    self.messages = "请先关闭用例文件，再运行测试用例"
                    self.message.start()
                    return
                except Exception as e:
                    self.messages = '运行错误，请重新运行或重启软件'
                    self.message.start()
                    return
                if len(self.CaseSumIndex) != 0:
                    if os.path.exists(r'D:\自动化测试截图'):
                        shutil.rmtree(r'D:\自动化测试截图')
                    if self.comboBox.currentText() == 'Google Chrome':
                        VersionYaml = ParseYaml().ReadParameter('Version')
                        if VersionYaml == "":
                            # 获取浏览器的版本号
                            driver = webdriver.Chrome(
                                executable_path=DRIVERS_PATH + 'chrome\\' + '70.0.3538.97\\chromedriver.exe')
                            driver.get('chrome://version/')
                            version = driver.find_element_by_css_selector('#version > span:nth-child(1)').text
                            driver.quit()
                            VersionList = ['70', '71', '72', '73', '74', '75', '76', '77', '78']
                            for i in VersionList:
                                if i == version[:2]:
                                    # 写入相关数据
                                    self.writeyaml.Write_Yaml_Updata('Version', version[:2])  # 将value写入ip.yaml文件中
                                elif int(version[:2]) < 70 or int(version[:2]) > 78:
                                    self.messages = "浏览器版本不符合，请更新浏览器版本号"
                                    self.message.starts()
                                    return
                        else:
                            pass
                        # 清除一些稍后写入的数据
                        self.writeyaml.Write_Yaml_Updata('CaseNum', 0)
                        # 写入IP
                        self.writeyaml.Write_Yaml_Updata('IP', self.lineEdit.text())
                        # 写入浏览器类型
                        self.writeyaml.Write_Yaml_Updata('Browser', self.comboBox.currentText())
                        # 写入是否生成测试报告
                        self.writeyaml.Write_Yaml_Updata('ReportAddress', self.label_15.text())
                        # 写入循环次数
                        self.writeyaml.Write_Yaml_Updata('loop', self.spinBox.text())
                        # 获取导入的测试用例路径
                        self.writeyaml.Write_Yaml_Updata('ImportAddress', self.newfile)
                        # 获取模块信息
                        self.writeyaml.Write_Yaml_Updata('Moudle', self.comboBox_2.currentText())
                        # 清空log框
                        self.listWidget_2.clear()
                        self.listWidget_2.setGeometry(QtCore.QRect(180, 20, 762, 471))
                        # log框置顶
                        self.listWidget_2.raise_()
                        # 初始化定时器
                        self.timercase = TimerCase()
                        # 运行测试用例
                        if self.radioButton.isChecked():
                            self.runtestthread = threading.Thread(target=TestPaperless().RunReport)
                        else:
                            self.runtestthread = threading.Thread(target=TestPaperless().TestCase)
                        self.displaycasethread = threading.Thread(target=self.displaycase)

                        if self.runtestthread.isAlive() is False:
                            self.runtestthread.setDaemon(True)
                            self.runtestthread.start()
                        self.timer2.timeout.connect(self.OutPut)
                        self.timer2.start(200)
                        # 设置按钮状态
                        self.toolButton_2.setDisabled(True)
                        self.toolButton_13.setDisabled(True)
                        self.toolButton.setDisabled(True)
                        self.toolButton_14.setDisabled(False)
                        self.listWidget.setDisabled(True)
                        self.toolButton_17.setDisabled(False)
                        self.toolButton_17.setStyleSheet(
                            "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}"
                        )
                        self.toolButton_2.setStyleSheet(
                            "QToolButton{font:9pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
                        self.toolButton_13.setStyleSheet(
                            "QToolButton{font:12pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
                        self.toolButton.setStyleSheet(
                            "QToolButton{font:12pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
                        self.toolButton_14.setStyleSheet(
                            "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
                        self.displaycasethread.setDaemon(True)
                        self.displaycasethread.start()
        except WebDriverException:
            self.messages = '浏览器驱动异常，无法正常运行，清检查浏览器驱动'
            self.message.start()
        except Exception as e:
            print(e)
            # logger.info(e)

    def displaycase(self):
        """
        用例, 运行时间显示
        """
        while self.toolButton_13.isEnabled() is False:
            if self.runtestthread.isAlive():
                # 读取用例已运行数量
                self.all_already_case = ParseYaml().ReadParameter('CaseNum')
                self.label_11.setText(str(self.all_already_case))
            if self.toolButton_13.isEnabled() is False:
                # 当启动按钮在禁用状态时，开始计时
                self.label_12.setText(self.timercase.Timer())
            if self.runtestthread.isAlive() is False and self.toolButton_13.isEnabled() is False:
                self.toolButton_2.setDisabled(False)
                self.toolButton_13.setDisabled(False)
                self.toolButton.setDisabled(False)
                self.toolButton_14.setDisabled(True)
                self.listWidget.setDisabled(False)
                self.toolButton_13.disconnect()
                self.toolButton_13.clicked.connect(self.runCase)
                self.toolButton_2.setStyleSheet(
                    "QToolButton{font:9pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
                self.toolButton_13.setStyleSheet(
                    "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
                self.toolButton.setStyleSheet(
                    "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
                self.toolButton_14.setStyleSheet(
                    "QToolButton{font:12pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
                self.listWidget_2.setGeometry(QtCore.QRect(640, 20, 301, 471))
            time.sleep(1)

    def Suspend(self):
        """
        暂停用例运行，需要运行完当期用例
        :return:
        """
        self.writeyaml.Write_Yaml_Updata('IP', '暂停运行')
        # 启动 按钮点击后，启用
        self.toolButton_13.setEnabled(True)
        self.toolButton_13.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")
        self.toolButton_13.disconnect()
        self.toolButton_13.clicked.connect(self.Continue)
        # 暂停 按钮变成灰色，禁用用
        self.toolButton_14.setEnabled(False)
        self.toolButton_14.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")

    def Continue(self):
        """
        继续用例的运行
        :return:
        """
        self.writeyaml.Write_Yaml_Updata('IP', self.lineEdit.text())
        # 启动 按钮点击后，禁用，并置灰
        self.toolButton_13.setEnabled(False)
        self.toolButton_13.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#8e7f7c;color: white;border-radius:3px;}")
        self.timer2.start(1000)
        # 暂停 按钮变成蓝色，启用
        self.toolButton_14.setEnabled(True)
        self.toolButton_14.setStyleSheet(
                "QToolButton{font:12pt '宋体';background-color:#419AD9;color: white;border-radius:3px;}")

    def ComboxValue(self):
        '''
        下拉框事件，有模板导入时选择模块进行用例显示
        :param tag:
        :return:
        '''
        if self.import_num != 0:
            self.spinBox.setValue(1)
            # 总用例数显示
            if self.comboBox_2.currentText() == '全部':
                self.label_10.setText(str(len(self.CaseSumIndex)))
            else:
                SheetList = self.book.getColumnValue(0, testCase_Sheet)
                self.MoudleSum = [i for i in self.CaseSumIndex if SheetList[i] == self.comboBox_2.currentText()]
                self.label_10.setText(str(len(self.MoudleSum)))

    def closeEvent(self, event):
        """
        重写closeEvent方法，实现dialog窗体关闭时执行一些代码
        :param event: close()触发的事件
        :return: None
        """
        if self.toolButton_14.isEnabled() is True:
            QMessageBox.about(self, "提示", "请暂停用例运行后再退出")
            event.ignore()
        else:
            box = QMessageBox(QMessageBox.Question, self.tr("提示"), self.tr("您确定要退出吗？"), QMessageBox.NoButton, self)
            yr_btn = box.addButton(self.tr("确定"), QMessageBox.YesRole)
            box.addButton(self.tr("取消"), QMessageBox.NoRole)
            box.exec_()
            if box.clickedButton() == yr_btn:
                self.writeyaml.Write_Yaml_Updata('Version', '')
                self.close()
            else:
                event.ignore()

    def readQssFile(self, filePath):
        with open(filePath, 'r', encoding='utf-8') as fileObj:
            styleSheet = fileObj.read()
        return styleSheet

    def _async_raise(self, tid, exctype):
        """raises the exception, performs cleanup if needed"""
        tid = ctypes.c_long(tid)
        if not inspect.isclass(exctype):
            exctype = type(exctype)
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(exctype))
        if res == 0:
            raise ValueError("invalid thread id")
        elif res != 1:
            # """if it returns a number greater than one, you're in trouble,
            # and you should call it again with exc=NULL to revert the effect"""
            ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
            raise SystemError("PyThreadState_SetAsyncExc failed")

    def stop_thread(self, thread):
        self._async_raise(thread.ident, SystemExit)


class Tips_1(QDialog):
    '''
    关键字主要数据输入框
    '''
    dialogSignel = pyqtSignal(str, str, str, str)

    def __init__(self, parent=None):
        super(Tips_1, self).__init__(parent)
        # 设置窗口大小
        self.setFixedSize(321, 231)

        # 主窗口
        self.listView = QtWidgets.QListView(self)
        # 步骤描述
        self.label = QtWidgets.QLabel('步骤描述', self)
        # 步骤描述输入框
        self.lineEdit = QtWidgets.QLineEdit(self)
        # 定位方式
        self.label_2 = QtWidgets.QLabel('定位方式', self)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        # 定位方式输入框
        self.lineEdit_2 = QtWidgets.QComboBox(self)
        self.lineEdit_2.addItems(['css', 'xpath', 'id', 'name', 'class', 'link', 'link_text', 'tag'])
        # 表达式
        self.label_3 = QtWidgets.QLabel('表达式', self)
        self.label_3.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.label_3.setMouseTracking(True)
        self.label_3.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        # 表达式输入框
        self.lineEdit_3 = QtWidgets.QLineEdit(self)
        # 操作值
        self.label_4 = QtWidgets.QLabel('操作值', self)
        self.label_4.setAlignment(QtCore.Qt.AlignCenter)
        # 操作值输入框
        self.lineEdit_4 = QtWidgets.QLineEdit(self)
        # 确定
        self.toolButton = QtWidgets.QToolButton(self)
        self.toolButton.setText('确定')
        # 确定按钮绑定输入事件
        self.toolButton.clicked.connect(self.inputTips)
        # 确定按钮绑定回车
        self.toolButton.setShortcut(QtCore.Qt.Key_Return)
        # 取消
        self.toolButton_2 = QtWidgets.QToolButton(self)
        self.toolButton_2.setText('取消')
        # 取消按钮绑定关闭事件
        self.toolButton_2.clicked.connect(self.CloseTips)

    def PointTips_1(self):
        '''
        设置控件坐标，步骤描述，定位方式，表达式，操作值
        :return:
        '''
        self.listView.setGeometry(QtCore.QRect(0, 0, 321, 231))
        self.label.setGeometry(QtCore.QRect(30, 30, 54, 20))
        self.lineEdit.setGeometry(QtCore.QRect(100, 30, 171, 20))
        self.label_2.setGeometry(QtCore.QRect(30, 70, 54, 20))
        self.lineEdit_2.setGeometry(QtCore.QRect(100, 70, 171, 20))
        self.label_3.setGeometry(QtCore.QRect(30, 110, 54, 20))
        self.lineEdit_3.setGeometry(QtCore.QRect(100, 110, 171, 20))
        self.label_4.setGeometry(QtCore.QRect(30, 150, 54, 20))
        self.lineEdit_4.setGeometry(QtCore.QRect(100, 150, 171, 20))
        self.toolButton.setGeometry(QtCore.QRect(130, 190, 51, 25))
        self.toolButton_2.setGeometry(QtCore.QRect(220, 190, 51, 25))

    def PointTips_2(self):
        '''
        设置控件坐标，步骤描述，定位方式，表达式
        :return:
        '''
        self.setFixedSize(321, 187)
        self.label_4.setVisible(False)
        self.lineEdit_4.setVisible(False)
        self.listView.setGeometry(QtCore.QRect(0, 0, 321, 187))
        self.label.setGeometry(QtCore.QRect(30, 30, 54, 20))
        self.lineEdit.setGeometry(QtCore.QRect(100, 30, 171, 20))
        self.label_2.setGeometry(QtCore.QRect(30, 70, 54, 20))
        self.lineEdit_2.setGeometry(QtCore.QRect(100, 70, 171, 20))
        self.label_3.setGeometry(QtCore.QRect(30, 110, 54, 20))
        self.lineEdit_3.setGeometry(QtCore.QRect(100, 110, 171, 20))
        self.toolButton.setGeometry(QtCore.QRect(130, 150, 51, 25))
        self.toolButton_2.setGeometry(QtCore.QRect(220, 150, 51, 25))

    def PointTips_3(self):
        '''
        设置控件坐标，步骤描述，操作值
        :return:
        '''
        self.setFixedSize(321, 157)
        self.label_2.setVisible(False)
        self.lineEdit_2.setVisible(False)
        self.label_3.setVisible(False)
        self.lineEdit_3.setVisible(False)
        self.listView.setGeometry(QtCore.QRect(0, 0, 321, 157))
        self.label.setGeometry(QtCore.QRect(30, 30, 54, 20))
        self.lineEdit.setGeometry(QtCore.QRect(100, 30, 171, 20))
        self.label_4.setGeometry(QtCore.QRect(30, 70, 54, 20))
        self.lineEdit_4.setGeometry(QtCore.QRect(100, 70, 171, 20))
        self.toolButton.setGeometry(QtCore.QRect(130, 120, 51, 25))
        self.toolButton_2.setGeometry(QtCore.QRect(220, 120, 51, 25))

    def PointTips_4(self):
        '''
        设置控件坐标，步骤描述
        :return:
        '''
        self.setFixedSize(321, 127)
        self.label_2.setVisible(False)
        self.lineEdit_2.setVisible(False)
        self.label_3.setVisible(False)
        self.lineEdit_3.setVisible(False)
        self.label_4.setVisible(False)
        self.lineEdit_4.setVisible(False)
        self.listView.setGeometry(QtCore.QRect(0, 0, 321, 157))
        self.label.setGeometry(QtCore.QRect(30, 30, 54, 20))
        self.lineEdit.setGeometry(QtCore.QRect(100, 30, 171, 20))
        self.toolButton.setGeometry(QtCore.QRect(130, 90, 51, 25))
        self.toolButton_2.setGeometry(QtCore.QRect(220, 90, 51, 25))

    def PointTips_5(self):
        '''
        设置控件坐标，步骤描述
        :return:
        '''
        self.lineEdit_2.clear()
        self.listView.setGeometry(QtCore.QRect(0, 0, 321, 231))
        self.label.setGeometry(QtCore.QRect(30, 30, 54, 20))
        self.lineEdit.setGeometry(QtCore.QRect(100, 30, 171, 20))
        self.label_3.setGeometry(QtCore.QRect(30, 70, 54, 20))
        self.label_3.setText('判断条件')
        self.lineEdit_3.setGeometry(QtCore.QRect(100, 70, 171, 20))
        self.label_2.setGeometry(QtCore.QRect(30, 110, 54, 20))
        self.label_2.setText('成立条件')
        self.lineEdit_2.setGeometry(QtCore.QRect(100, 110, 171, 20))
        self.label_4.setGeometry(QtCore.QRect(30, 150, 54, 20))
        self.label_4.setText('满足条件')
        self.lineEdit_4.setGeometry(QtCore.QRect(100, 150, 171, 20))
        self.toolButton.setGeometry(QtCore.QRect(130, 190, 51, 25))
        self.toolButton_2.setGeometry(QtCore.QRect(220, 190, 51, 25))
        self.lineEdit_2.addItems(['=', '!=', '>', '< ', '>=', '<='])

    def CloseTips(self):
        self.close()

    def inputTips(self):
        '''
        传递子窗口输入数据，将输入显示到编辑框中
        :return:
        '''
        # 获取步骤描述
        self.bz = self.lineEdit.text()
        # 获取定位方式
        self.by = self.lineEdit_2.currentText()
        # 获取定位表达式
        self.location = self.lineEdit_3.text()
        # 输入值
        self.value = self.lineEdit_4.text()
        if self.bz == '' and self.lineEdit.isVisible() is True or \
                self.by == '' and self.lineEdit_2.isVisible() is True or \
                self.location == '' and self.lineEdit_3.isVisible() is True:
            QMessageBox.about(self, "提示", "输入不能为空")
        else:
            self.dialogSignel.emit(self.bz, self.by, self.location, self.value)
            self.close()

class Tips_2(QDialog):
    '''用于重命名时弹窗'''
    # 发射子窗口的值的信号
    dialogSignel_2 = pyqtSignal(str, str, str)

    def __init__(self, parent=None):
        super(Tips_2, self).__init__(parent)
        self.setAcceptDrops(True)
        self.listView = QtWidgets.QListView(self)
        self.label = QtWidgets.QLabel('用例编号', self)
        self.lineEdit = QtWidgets.QLineEdit(self)
        self.label_2 = QtWidgets.QLabel('用例标题', self)
        self.lineEdit_2 = QtWidgets.QLineEdit(self)
        self.label_3 = QtWidgets.QLabel('预置用例', self)
        self.lineEdit_3 = QtWidgets.QComboBox(self)
        self.toolButton = QtWidgets.QToolButton(self)
        self.toolButton.setText('确定')
        self.toolButton.clicked.connect(self.inputTips)
        # 确定按钮绑定回车
        self.toolButton.setShortcut(QtCore.Qt.Key_Return)
        self.toolButton_2 = QtWidgets.QToolButton(self)
        self.toolButton_2.setText('取消')
        self.toolButton_2.clicked.connect(self.CloseTips)

    def PointTips_1(self):
        # 添加新的用例布局
        self.setFixedSize(321, 187)
        self.listView.setGeometry(QtCore.QRect(0, 0, 321, 187))
        self.label.setGeometry(QtCore.QRect(30, 30, 54, 20))
        self.lineEdit.setGeometry(QtCore.QRect(100, 30, 171, 20))
        self.label_2.setGeometry(QtCore.QRect(30, 70, 54, 20))
        self.lineEdit_2.setGeometry(QtCore.QRect(100, 70, 171, 20))
        self.label_3.setGeometry(QtCore.QRect(30, 110, 54, 20))
        self.lineEdit_3.setGeometry(QtCore.QRect(100, 110, 171, 20))
        self.toolButton.setGeometry(QtCore.QRect(130, 150, 51, 25))
        self.toolButton_2.setGeometry(QtCore.QRect(220, 150, 51, 25))

    def PointTips_2(self):
        self.setFixedSize(321, 157)
        self.listView.setGeometry(QtCore.QRect(0, 0, 321, 157))
        self.label.setGeometry(QtCore.QRect(30, 30, 54, 20))
        self.lineEdit.setGeometry(QtCore.QRect(100, 30, 171, 20))
        self.label_2.setGeometry(QtCore.QRect(30, 70, 54, 20))
        self.lineEdit_2.setGeometry(QtCore.QRect(100, 70, 171, 20))
        self.toolButton.setGeometry(QtCore.QRect(130, 120, 51, 25))
        self.toolButton_2.setGeometry(QtCore.QRect(220, 120, 51, 25))
        self.label_3.setVisible(False)
        self.lineEdit_3.setVisible(False)
        self.label_2.setText('用例编号')
        # 设置旧用例编号禁止写入
        self.lineEdit.setDisabled(True)

    def CloseTips(self):
        self.close()

    def inputTips(self):
        '''
        传递子窗口输入数据，将输入显示到编辑框中
        :return:
        '''
        # 获取用例名称
        self.mc = self.lineEdit.text()
        self.caseTitle = self.lineEdit_2.text()
        self.preNode = self.lineEdit_3.currentText()
        if self.mc == '' or self.caseTitle == '':
            QMessageBox.about(self, "提示", "用例编号输入不能为空")
        else:
            self.dialogSignel_2.emit(self.mc, self.caseTitle, self.preNode)
            self.close()

class Tips_3(QDialog):
    # 发射子窗口的值的信号
    dialogSignel_2 = pyqtSignal()

    def __init__(self, parent=None):
        super(Tips_3, self).__init__(parent)
        self.setWindowTitle('log查看窗')
        self.setAcceptDrops(True)
        self.listwidget = QtWidgets.QListWidget(self)
        bgimg3 = RESOURSE_PATH + 'bgimg3.png'
        bgimg3 = bgimg3.replace('\\', '/')
        self.listwidget.setStyleSheet("QListWidget{background-image: url(%s);color:#FFFFFF;}" % bgimg3)
        self.listwidget.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.PointTips_1()

    def PointTips_1(self):
        # 添加新的用例布局
        self.setFixedSize(600, 653)
        self.listwidget.setGeometry(QtCore.QRect(0, 0, 600, 653))

class MyCurrentQueue(QtWidgets.QListWidget):
    '''
    由于InternalMove拖拽方法无法触发dropEvent方法
    需要继承QListWidget重写dropEvent方法
    '''

    def __init__(self, parent=None):
        super(MyCurrentQueue, self).__init__(parent)

    def dropEvent(self, event):
        # print('%d '%self.currentRow(),end = '')#用于打印拖拽前后目标item的索引值，以便观察
        super(MyCurrentQueue, self).dropEvent(event)  # 如果不调用父类的构造方法，拖拽操作将无法正常进行
        # print(self.currentRow())
        for i in range(27):
            if i % 2 == 0:
                # 设置item背景颜色
                self.item(i).setBackground(QColor('#696969'))
                self.setSortingEnabled(False)
            else:
                # 设置item背景颜色
                self.item(i).setBackground(QColor('#363636'))
                self.setSortingEnabled(False)

class message(QThread):
    signal = pyqtSignal()

    def __init__(self, Automation):
        super(message, self).__init__()
        self.automaint = Automation

    def run(self):
        self.signal.emit()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    widget = Automation()
    styleSheet = widget.readQssFile(QSS_PATH+'Automator.qss')
    widget.setStyleSheet(styleSheet)
    widget.show()

    app.exec_()
    sys.exit(app.exec_())
