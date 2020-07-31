#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: AddCase.py
@time: 2020/4/13 10:30
@desc: 生成用例
"""

from Utils.interface_utils.interface_param import ParamRead
from Utils.ParseExcelWin import ParseExcelWin
from Utils.ParseYaml import ParseYaml
from Utils.ParseExcelXlrd import ParseExcelXlrd
from Utils.Logger import Logger
from Utils.ConfigRead import *
from faker import Faker
import random
import copy
fake = Faker('zh_cn')


class AddCase(object):

    def __init__(self, parseexcelxlrd, parseexcelwin, moudel):
        self.parseyaml = ParseYaml()
        # self.excel_path = self.parseyaml.ReadAPI_Paramter('TEST_PATH')
        self.moudle = moudel
        self.parseexcelxlrd = parseexcelxlrd
        self.parseexcelwin = parseexcelwin
        self.paramread = ParamRead(self.parseexcelxlrd)
        # 正常请求数据
        self.request_data = {}
        # 开始写入行
        self.start_row = 0
        # 接口请求参数
        self.params = []
        self.api_logger = Logger('api-logger', APILOGS_PATH).getlog()

    def addcase(self):
        """读取表格参数，写入用例"""
        try:
            if self.paramread.get_api_sheet_data(self.moudle) != []:
                # 获取正常请求中的参数
                self.params = self.paramread.get_api_sheet_data(self.moudle)
                # 获取“描述”行数据
                describe = self.parseexcelxlrd.getMergeColumnValue(self.moudle, 1)
                describeindex = [i for i, x in enumerate(describe) if x == '描述']
                # 描述行具体参数
                represvalue = self.parseexcelxlrd.getRowValue(self.moudle, int(describeindex[0])+2)
                represindex = [i for i, x in enumerate(represvalue) if x in self.params]
                # 将请求参数与正常请求的值对应写入字典中，用于之后生成用例所用到的参数
                for i, x in enumerate(represindex):
                    if self.parseexcelxlrd.getCellValue(self.moudle, int(describeindex[0])+2, int(x)+1) == self.params[i]:
                        self.request_data[self.params[i]] = \
                            self.parseexcelxlrd.getCellValue(self.moudle, int(describeindex[0])+2+1, int(x)+1)
                # 用例开始写入行数
                self.start_row = int(describeindex[0])+4
                self.write_case()
                self.write_case_title()
        except Exception:
            self.api_logger.info('用例新增失败，请重试!')

    def fake_value(self):
        """生成随机值"""
        fakes = ['', ' ']
        # 随机字符串
        random_string = fake.sentence(nb_words=4, variable_nb_words=True)
        fakes.append(random_string)
        # 随机小数点后两位小数
        random_float = round(random.uniform(1, 10), 2)
        fakes.append(random_float)
        # 随机整数
        random_int = fake.random_number(3)
        fakes.append(random_int)
        # 随机负数
        random_neg = -fake.random_number(2)
        fakes.append(random_neg)
        # 随机日期
        random_data = fake.date(pattern="%Y-%m-%d")
        fakes.append(random_data)
        return fakes

    def gen_case(self):
        """
        生成用例， list存放格式为s = [[],[]]
        """
        # 用例集合
        cases = []
        # 每个参数按照特定格式生成用例
        for i, v in self.request_data.items():
            # 通过深拷贝复制正常请求中的字典
            copy_cases = copy.deepcopy(self.request_data)
            for f in self.fake_value():
                # 通过生成的随机值替换正常请求中的值，到达用例的目的
                copy_cases[i] = f
                # 用于存放单条用例
                one_case = []
                for s, c in copy_cases.items():
                    one_case.append(c)
                cases.append(one_case)
        return cases

    def write_case(self):
        """将用例写入excel表格中"""
        try:
            for rowno, cases in enumerate(self.gen_case()):
                for columnno, case in enumerate(cases):
                    self.parseexcelwin.writeCellValue(self.moudle, self.start_row+rowno, columnno+2, case)
            self.parseexcelwin.save()
        except Exception:
            self.api_logger.info('用例写入失败，请重试!')

    def write_case_title(self):
        """生成用例标题，并写入excel表格中"""
        try:
            title = ['为空', '为空格', '为字符串', '为小数', '为整数', '为负数', '为日期格式']
            add_row = 0
            for rowno, param in enumerate(self.params):
                for columnno, t in enumerate(title):
                    self.parseexcelwin.writeCellValue(self.moudle,
                                                      self.start_row+add_row, 1, '%s%s' % (param, t))
                    add_row += 1
            self.parseexcelwin.borderAround(self.moudle, 1, 1,
                                            self.parseexcelxlrd.wookbook.sheet_by_name(self.moudle).nrows,
                                            self.parseexcelxlrd.wookbook.sheet_by_name(self.moudle).ncols)
            self.parseexcelwin.save()
            self.parseexcelwin.close()
        except Exception:
            self.api_logger.info('用例写入失败，请重试!')


if __name__ == '__main__':
    import time
    print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
    p = ParseExcelXlrd(r'E:/接口测试用例/ce/token.xlsx')
    w = ParseExcelWin(r'E:/接口测试用例/ce/token.xlsx')
    AddCase(p, w, '编辑参会人员').addcase()
    print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
