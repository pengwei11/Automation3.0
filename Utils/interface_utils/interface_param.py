#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: interface_param.py
@time: 2020/4/13 8:42
@desc: 接口参数读取
"""
from Utils.ParseExcelXlrd import ParseExcelXlrd
from Utils.ConfigRead import *
from Utils.interface_utils.interface_config import *
from Utils.ParseYaml import ParseYaml
from Utils.Logger import Logger
import json


class ParamRead(object):

    def __init__(self, parseexcel):
        # parseexcel = ParseExcelXlrd(self.excel_path)
        self.parseexcel = parseexcel
        self.parseyaml = ParseYaml()
        self.api_logger = Logger('api-logger', APILOGS_PATH).getlog()

    """
    读取第一个工作表中的参数
    """

    def get_api_path(self):
        """ 接口地址"""
        api_path = self.parseexcel.getColumnValue(0, API_PATH)
        del api_path[0:2]
        return api_path

    def get_api_sheet(self):
        """接口工作表"""
        api_sheet = self.parseexcel.getColumnValue(0, API_SHEET)
        del api_sheet[0:2]
        return api_sheet

    def get_api_method(self):
        """接口请求类型"""
        api_method = self.parseexcel.getColumnValue(0, API_METHOD)
        del api_method[0:2]
        return api_method

    def get_api_isrun(self):
        """接口是否执行"""
        api_isrun = self.parseexcel.getColumnValue(0, API_ISRUN)
        del api_isrun[0:2]
        return api_isrun

    def get_api_iscase(self):
        """接口是否生成用例"""
        get_api_iscase = self.parseexcel.getColumnValue(0, API_ISCASE)
        del get_api_iscase[0:2]
        return get_api_iscase

    """
    读取接口工作表中的参数
    """
    def get_api_sheet_token(self, sheetname):
        """获取制定excel表格中的token获取方法"""
        try:
            describe = self.parseexcel.getMergeColumnValue(sheetname, 1)
            describeindex = [i for i, x in enumerate(describe) if x == 'token提取']
            token = self.parseexcel.getCellValue(sheetname, describeindex[0]+2, 2)
            if "，" in token:
                token.replace("，", ",")
            token = token.split(',')
            return token
        except Exception as e:
            self.api_logger.info('token读取失败，请检查用例文件')

    def get_api_sheet_name(self, sheetname):
        """接口名称"""
        sheet_api_name = self.parseexcel.getCellValue(sheetname, API_SHEET_NAME, 2)
        return sheet_api_name

    def get_api_sheet_id(self, sheetname):
        """接口编号"""
        sheet_api_id = self.parseexcel.getCellValue(sheetname, API_SHEET_ID, 2)
        return sheet_api_id

    def get_api_sheet_path(self, sheetname):
        """接口地址（增加前缀）"""
        sheet_api_path = self.parseexcel.getCellValue(sheetname, API_SHEET_PATH, 2)
        sheet_api_path = 'http://'+self.parseyaml.ReadAPI_Paramter('API_PATH')+sheet_api_path
        return sheet_api_path

    def get_api_sheet_path_new(self, sheetname):
        """接口地址"""
        sheet_api_path = self.parseexcel.getCellValue(sheetname, API_SHEET_PATH, 2)
        return sheet_api_path

    def get_api_sheet_method(self, sheetname):
        """接口请求类型"""
        api_sheet_method = self.parseexcel.getCellValue(sheetname, API_SHEET_METHOD, 2)
        return api_sheet_method

    def get_api_sheet_code(self, sheetname):
        """接口返回值状态码"""
        try:
            # 获取接口参数描述列
            describe = self.parseexcel.getMergeColumnValue(sheetname, 1)
            describeindex = [i for i, x in enumerate(describe) if x == '返回值状态码']
            # 获取+1列的值
            code = self.parseexcel.getCellValue(sheetname, describeindex[0]+2, 2)
            if code is None:
                self.api_logger.info('预期返回值状态码为空, 请注意填写，否则结果无法写入')
            return code
        except ValueError:
            self.api_logger.info('接口请求头读取失败，请检查用例文件')

    def get_api_sheet_header(self, sheetname):
        """接口请求头"""
        try:
            # 获取接口参数描述列
            describe = self.parseexcel.getMergeColumnValue(sheetname, 1)
            describeindex = [i for i, x in enumerate(describe) if x == '请求头']
            headers = {}
            # 获取对应请求头的信息存入字典
            for d in describeindex:
                header_type = self.parseexcel.getCellValue(sheetname, d+2, 2)
                if header_type == "":
                    continue
                header_param = self.parseexcel.getCellValue(sheetname, d+2, 3)
                headers[header_type] = header_param
            if headers != {}:
                headers = json.dumps(headers)
                return headers
            else:
                return None
        except ValueError:
            self.api_logger.info('接口请求头读取失败，请检查用例文件')

    def get_api_sheet_data(self, sheetname):
        """接口请求参数"""
        try:
            # 获取接口参数描述列
            describe = self.parseexcel.getMergeColumnValue(sheetname, 1)
            describeindex = [i for i, x in enumerate(describe) if x == '请求参数']
            datas = []
            # 获取对应请求头的信息存入字典
            for d in describeindex:
                data_value = self.parseexcel.getCellValue(sheetname, d + 2, 3)
                if data_value == "":
                    continue
                else:
                    datas.append(data_value)
            return datas
        except ValueError:
            self.api_logger.info('接口请求参数获取失败，请检查用例文件')

    def get_api_sheet_except(self, sheetname):
        """预期返回值"""
        try:
            # 获取接口参数描述列
            describe = self.parseexcel.getMergeColumnValue(sheetname, 1)
            except_row = describe.index('预期返回值')
            excepts = self.parseexcel.getCellValue(sheetname, except_row+2, 2)
            return excepts
        except ValueError:
            self.api_logger.info('预期返回值获取失败，请检查用例文件')

    def get_api_sheet_param(self, sheetname):
        """获取生成的接口参数的值"""
        try:
            describe = self.parseexcel.getMergeColumnValue(sheetname, 1)
            describeindex = [i for i, x in enumerate(describe) if x == '描述']
            describe_row = int(describeindex[0])+2
            max_row = self.parseexcel.wookbook.sheet_by_name(sheetname).nrows
            params = []
            for i in range(max_row-describe_row):
                # 从描述行的下一行开始逐行获取数据，并删除每行中的最后四位和第一位
                p_row = self.parseexcel.getRowValue(sheetname, describe_row+i+1)
                del p_row[0]
                del p_row[-4:]
                # 将data与param合并成字典
                p_row_dict = dict(zip(self.get_api_sheet_data(sheetname), p_row))
                p_row_json = json.dumps(p_row_dict, ensure_ascii=False)
                params.append(p_row_json)
                # 将字典转换为json
            return params
        except Exception:
            self.api_logger.info('生成接口参数值获取失败，请检查用例文件')

if __name__ == '__main__':
    pa = ParseExcelXlrd(r'E:\智慧校园接口测试用例\0登录.xlsx')
    # print(ParamRead().get_api_path())
    # print(ParamRead().get_api_sheet())
    # print(ParamRead().get_api_method())
    # print(ParamRead().get_api_isrun())
    # print(ParamRead().get_api_sheet_name('查询离线会议室服务器'))
    # print(ParamRead().get_api_sheet_id('查询离线会议室服务器'))
    # print(ParamRead().get_api_sheet_path('查询离线会议室服务器'))
    # print(ParamRead(pa).get_api_sheet_data('编辑参会人员'))
    # print(ParamRead().get_api_sheet_data('查询离线会议室服务器'))
    # print(ParamRead().get_api_sheet_except('查询离线会议室服务器'))
    print(ParamRead(pa).get_api_sheet_param('登录'))
    # print(type(eval(ParamRead(pa).get_api_sheet_header('获取token'))))
    # print(ParamRead(pa).get_api_sheet_path('获取token'))
