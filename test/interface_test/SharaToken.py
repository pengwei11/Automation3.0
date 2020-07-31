#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: SharaToken.py
@time: 2020/5/19 14:35
@desc: 验证是否有token接口，并传递token值
"""
from action.interface_action.RunMethod import RunMethod
from Utils.ConfigRead import *
from Utils.interface_utils.interface_config import *
from Utils.interface_utils.interface_param import ParamRead
from Utils.ParseExcelWin import ParseExcelWin
from Utils.ParseYaml import ParseYaml
from Utils.ParseExcelXlrd import ParseExcelXlrd
from Utils.Logger import Logger
from datetime import datetime
import traceback
import ast


class SharaToken(object):

    def __init__(self):
        self.parseyaml = ParseYaml()
        self.file_path = self.parseyaml.ReadAPI_Paramter('TOKEN_FILE_PATH')
        if self.file_path is not None:
            self.parseexcelxlrd = ParseExcelXlrd(self.file_path)
            self.parseexcelwin = ParseExcelWin(self.file_path)
            # 用例读取
            self.paramread = ParamRead(self.parseexcelxlrd)
        # 接口测试类型
        self.runmethod = RunMethod()
        # 将状态码，实际结果，测试时间保存至字典中（测试结果通过实际结果或状态码判断）
        self.api_code = []
        self.api_result = []
        self.api_testtime = []
        self.api_logger = Logger('api-logger', APILOGS_PATH).getlog()

    def token(self):
        """通过测试token接口，获取token值"""
        try:
            if self.file_path is not None:
                print("*************获取token*************")
                self.api_logger.info("*************获取token*************")
                self.api_sheet = self.parseexcelxlrd.getCellValue(0, 3, API_SHEET)
                # 获取测试行
                describe = self.parseexcelxlrd.getMergeColumnValue(self.api_sheet, 1)
                describeindex = [i for i, x in enumerate(describe) if x == '描述']
                describe_row = int(describeindex[0]) + 1
                response = ''
                # 请求方式
                method = self.paramread.get_api_sheet_method(self.api_sheet).lower()
                # 请求值
                value = self.paramread.get_api_sheet_param(self.api_sheet)[0]
                # 请求地址
                path = self.paramread.get_api_sheet_path(self.api_sheet)
                # 请求头
                hearder = self.paramread.get_api_sheet_header(self.api_sheet)
                # Get请求
                if method == 'get':
                    response = self.runmethod.get_main(path, ast.literal_eval(value), ast.literal_eval(hearder))
                # POST请求
                elif method == 'post':
                    response = self.runmethod.post_main(path, ast.literal_eval(value), ast.literal_eval(hearder))
                # delete请求
                elif method == 'delete':
                    response = self.runmethod.delete_main(path, ast.literal_eval(value), ast.literal_eval(hearder))
                # 获取token提取字段
                token_field = self.paramread.get_api_sheet_token(self.api_sheet)
                # 写入测试结果
                self.api_code.append(response.status_code)
                self.api_result.append(response.text)
                times = datetime.now()
                times.strftime('%Y:%m:%d %H:%M:%S')
                self.api_testtime.append(times)
                self.write_result()
                token = ''
                if len(token_field) == 1:
                    token = response.json()[token_field[0]]
                else:
                    token = response.json()
                    try:
                        for token_data in token_field:
                            token = token[token_data]
                    except Exception as e:
                        self.api_logger.info("token值提取字段错误")
                if token != '':
                    print("获取成功，本次token值为: \n%s" % token)
                    self.api_logger.info("获取成功，本次token值为: \n%s" % token)
                    return token
                else:
                    print("获取失败")
                    self.api_logger.info("获取失败")
            else:
                print('跳过token获取')
                self.api_logger.info('跳过token获取')
                return None
        except Exception:
            print(traceback.print_exc())
            self.api_logger.info(traceback.print_exc())
            print('接口请求错误，请检查接口地址')
            self.api_logger.info('接口请求错误，请检查接口地址')

    def write_result(self):
        """写入结果"""
        # 获取开始写入行
        describe = self.parseexcelxlrd.getMergeColumnValue(self.api_sheet, 1)
        describeindex = [i for i, x in enumerate(describe) if x == '描述']
        describe_row = int(describeindex[0]) + 2
        # 获取各个结果的写入列
        result_col_value = self.parseexcelxlrd.getRowValue(self.api_sheet, describe_row)
        for col, result in enumerate(result_col_value):
            if result == '状态码':
                code_col = col+1
            elif result == '实际结果':
                real_col = col+1
            elif result == '测试时间':
                testtime_col = col+1
            elif result == '测试结果':
                fruit_col = col+1
        # 写入状态码和测试结果，如果状态码为200则pass
        for row, code in enumerate(self.api_code):
            self.parseexcelwin.writeCellValue(self.api_sheet, describe_row+row+1, code_col, code)
            if code == '200' or code == 200:
                self.parseexcelwin.writeCellValue(self.api_sheet, describe_row + row + 1, fruit_col, 'Pass')
                self.parseexcelwin.fontcolor(self.api_sheet, describe_row + row + 1, fruit_col, 4)
            else:
                self.parseexcelwin.writeCellValue(self.api_sheet, describe_row + row + 1, fruit_col, 'Failed')
                self.parseexcelwin.fontcolor(self.api_sheet, describe_row + row + 1, fruit_col, 3)
        # 写入实际结果，如果结果为html，则写入html所保存的地址
        for row, real in enumerate(self.api_result):
            self.parseexcelwin.writeCellValue(self.api_sheet, describe_row+row+1, real_col, str(real))
        # 写入测试时间
        for row, times in enumerate(self.api_testtime):
            self.parseexcelwin.writeCellValue(self.api_sheet, describe_row+row+1, testtime_col, times)
        self.parseexcelwin.save()
        self.parseexcelwin.close()



if __name__ == '__main__':
    SharaToken().token()


