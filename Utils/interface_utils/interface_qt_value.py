#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: interface_qt_value.py
@time: 2020/5/26 9:43
@desc: 用于读取qt界面中导入的数据，并获取接口地址和接口参数
"""
from Utils.ParseExcelXlrd import ParseExcelXlrd
from Utils.interface_utils.interface_param import ParamRead
from Utils.interface_utils.interface_config import *
import os


class InterFace_Qt_Value(object):

    def __init__(self, files):
        self.files = files
        self.parseexcelxlrd = ""
        self.paramread = ""
        self.tree_dict = {}

    def select_api_sheet(self):
        """
        读取用例文件列表，返回接口地址字典{'name': ['api_name:api_1', 'api_name:api_2']}
        """
        if isinstance(self.files, list):
            for file in self.files:
                # 实例化excel读取功能
                self.parseexcelxlrd = ParseExcelXlrd(file)
                self.paramread = ParamRead(self.parseexcelxlrd)
                # 获取文件名称, 接口地址，工作表名
                file_name = os.path.splitext(os.path.basename(file))[0]
                # api_path = self.paramread.get_api_path()
                sheet_name = self.paramread.get_api_sheet()
                self.tree_dict.setdefault(file_name, sheet_name)
            return self.tree_dict

    def select_api_name(self):
        """
        读取用例文件列表，返回接口地址字典{'name': ['api_name:api_1', 'api_name:api_2']}
        """
        if isinstance(self.files, list):
            for file in self.files:
                # 实例化excel读取功能
                self.parseexcelxlrd = ParseExcelXlrd(file)
                self.paramread = ParamRead(self.parseexcelxlrd)
                # 获取文件名称, 接口地址，工作表名
                file_name = os.path.splitext(os.path.basename(file))[0]
                # api_path = self.paramread.get_api_path()
                sheet_name = self.paramread.get_api_sheet()
                self.tree_dict.setdefault(file_name, sheet_name)
            return self.tree_dict



if __name__ == '__main__':
    print(InterFace_Qt_Value(['E:/接口测试用例/ce/会议室接口.xlsx']).select_api_sheet())