#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: TestApi.py
@time: 2020/4/13 10:29
@desc: 运行程序
"""
from action.interface_action.RunMethod import RunMethod
from Utils.ConfigRead import *
from Utils.interface_utils.interface_config import *
from Utils.interface_utils.interface_param import ParamRead
from Utils.ParseExcelWin import ParseExcelWin
from Utils.ParseYaml import ParseYaml
from Utils.WriteFile import YamlWrite
from Utils.ParseExcelXlrd import ParseExcelXlrd
from test.interface_test.AddCase import AddCase
from Utils.Logger import Logger
from datetime import datetime
from test.interface_test.SharaToken import SharaToken
import time
import traceback
import os
import ast
import shutil


class TestApi(object):

    def __init__(self):
        self.parseyaml = ParseYaml()
        self.file_path = self.parseyaml.ReadAPI_Paramter('FILE_PATH')
        self.parseexcelxlrd = ''
        self.parseexcelwin = ''
        # 接口测试类型
        self.runmethod = RunMethod()
        # 用例读取
        self.paramread = ''
        # 接口yaml文件路径
        self.api_parameter = CONFIG_PATH+'interface_config\\'+'API_Parameter.yaml'
        # 将状态码，实际结果，测试时间保存至字典中（测试结果通过实际结果或状态码判断）
        self.api_code = []
        self.api_result = []
        self.api_testtime = []
        # 删除api-result文件下的所有文件夹
        if os.path.exists(API_RESULT):
            shutil.rmtree(API_RESULT)
            os.makedirs(API_RESULT)
        else:
            os.makedirs(API_RESULT)

    def test_api(self):
        """
        通过导入的用例，读取接口地址、接口类型、接口请求头、接口参数等数据，整合
        成一个接口请求数据，并对接口进行请求，获取返回值中的code，json进行预期值
        对比，从而进行接口测试接口写入。
        """
        try:
            api_logger = Logger('api-logger', APILOGS_PATH).getlog()
            # 获取token
            token = SharaToken().token()
            for dir_path in self.file_path:
                # 从yaml文件集合中读取测试文件
                self.parseexcelxlrd = ParseExcelXlrd(dir_path)
                self.paramread = ParamRead(self.parseexcelxlrd)
                # 获取执行的接口的行数(+3)
                runcase_rowno = [i for i in range(len(self.paramread.get_api_isrun()))
                                 if self.paramread.get_api_isrun()[i].lower() == 'y']
                # 通过‘执行’接口行数，获取接口工作表
                if not runcase_rowno:
                    print('**********无 测 试 用 例 ！**********')
                    api_logger.info('**********无 测 试 用 例 ！**********')
                    return
                else:
                    print('**********测 试 开 始**********')
                    api_logger.info('**********测 试 开 始**********')
                    print('**********当 前 测 试 文 件:%s' % dir_path)
                    api_logger.info('**********当 前 测 试 文 件:%s**********' % dir_path)
                    for r in runcase_rowno:
                        # 初始化测试结果集合
                        self.api_code = []
                        self.api_result = []
                        self.api_testtime = []
                        self.parseexcelwin = ParseExcelWin(dir_path)
                        # 获取工作表，将需要生成用例的工作表写入yaml，并调用用例生成模块
                        self.api_sheet = self.parseexcelxlrd.getCellValue(0, r+3, API_SHEET)
                        is_case = self.parseexcelxlrd.getCellValue(0, r+3, API_ISCASE)
                        if is_case.lower() == 'y':
                            print('"%s"接口新增用例中...' % self.api_sheet)
                            api_logger.info('"%s"接口新增用例中...' % self.api_sheet)
                            # YamlWrite(self.api_parameter).Write_Yaml_Updata('MOUDEL', self.api_sheet)
                            AddCase(self.parseexcelxlrd, self.parseexcelwin, self.api_sheet).addcase()
                            # 重新读取excel表格
                            self.parseexcelxlrd = ParseExcelXlrd(dir_path)
                            self.parseexcelwin = ParseExcelWin(dir_path)
                            self.paramread = ParamRead(self.parseexcelxlrd)
                        describe = self.parseexcelxlrd.getMergeColumnValue(self.api_sheet, 1)
                        describeindex = [i for i, x in enumerate(describe) if x == '描述']
                        describe_row = int(describeindex[0]) + 2
                        for index, value in enumerate(self.paramread.get_api_sheet_param(self.api_sheet)):
                            title = self.parseexcelxlrd.getCellValue(self.api_sheet, describe_row+index+1, 1)
                            print('正在测试"%s"请求，测试参数为:\n%s' % (title, value))
                            api_logger.info('正在测试"%s"请求，测试参数为:\n%s' % (title, value))
                            response = ''
                            # 请求方式
                            method = self.paramread.get_api_sheet_method(self.api_sheet).lower()
                            # 请求地址
                            path = self.paramread.get_api_sheet_path(self.api_sheet)
                            # 请求头
                            header = ast.literal_eval(self.paramread.get_api_sheet_header(self.api_sheet))
                            if token is not None:
                                # 添加token进请求头
                                header['token'] = token
                            # 智慧校园特殊请求头
                            # header['Authorization'] = 'Bearer mmySNbD0ZDDT7b5PyFM3YByUInyfALrG'
                            # Get请求
                            print(header)
                            if method == 'get':
                                response = self.runmethod.get_main(path, ast.literal_eval(value), header)
                            # POST请求
                            elif method == 'post':
                                response = self.runmethod.post_main(path, ast.literal_eval(value), header)
                            # delete请求
                            elif method == 'delete':
                                response = self.runmethod.delete_main(path, ast.literal_eval(value), header)
                            elif method == 'put':
                                response = self.runmethod.put_main(path, ast.literal_eval(value), header)
                            else:
                                print('暂未开启该类型的请求测试')
                                api_logger.info('暂未开启该类型的请求测试')
                            code = self.paramread.get_api_sheet_code(self.api_sheet)
                            try:
                                response_value = response.json()
                            except Exception:
                                try:
                                    response_value = response.text
                                except Exception:
                                    response_value = None
                            try:
                                if isinstance(response_value, dict):
                                    api_logger.info(
                                        '"%s"请求测试完成：\n状态码：%s\n实际结果：%s' % (title, response_value[code], response_value))
                                elif response_value is None:
                                    api_logger.info('本次请求为空，请检查接口与参数')
                            except Exception as e:
                                api_logger.info(traceback.print_exc(e))
                            else:
                                print('"%s"请求测试完成' % (title))
                            # try:
                            #     if response is not None:
                            #         print('"%s"请求测试完成：\n状态码：%s\n实际结果：%s' % (title, response.status_code, response.json()))
                            #         api_logger.info('"%s"请求测试完成：\n状态码：%s\n实际结果：%s' % (title, response.status_code, response.json()))
                            #     else:
                            #         print('本次请求为空，请检查接口与参数')
                            #         api_logger.info('本次请求为空，请检查接口与参数')
                            # except Exception:
                            #     print('"%s"请求测试完成' % (title))
                            #     api_logger.info('"%s"请求测试完成' % (title))
                            # 将接口测试结果写入字典中
                            if response_value is not None and isinstance(response_value, dict):
                                self.api_code.append(response_value[code])
                                self.api_result.append(response_value)
                            elif response_value is None:
                                self.api_code.append('None')
                                self.api_result.append('请求失败，请检查接口参数')
                            else:
                                self.api_code.append(response.status_code)
                                api_result_path = API_RESULT + self.api_sheet
                                if not os.path.exists(api_result_path):
                                    os.makedirs(api_result_path)
                                with open(api_result_path+'\\'+self.paramread.get_api_sheet_id(self.api_sheet)+'-'+title+'.txt',
                                          'w', encoding="utf-8") as file:
                                    file.write(response.text)
                                self.api_result.append\
                                    (api_result_path+'\\'+self.paramread.get_api_sheet_id(self.api_sheet)+'-'+title+'.txt')
                            # if response != '' and response is not None:
                            #     try:
                            #         self.api_code.append(response_value[code])
                            #     except Exception:
                            #         self.api_code.append(response.status_code)
                            #     try:
                            #         self.api_result.append(response_value)
                            #     except Exception:
                            #         # 如果测试结果不为json格式，则保存至api_result/接口名称 文件夹中，命名方式为'请求描述'.txt
                            #         api_result_path = API_RESULT+self.api_sheet
                            #         if not os.path.exists(api_result_path):
                            #             os.makedirs(api_result_path)
                            #         # else:
                            #         #     # 文件夹存在则删除该文件，再重新创建
                            #         #     shutil.rmtree(api_result_path)
                            #         #     os.makedirs(api_result_path)
                            #         with open(api_result_path+'\\'+self.paramread.get_api_sheet_id(self.api_sheet)+'-'+title+'.txt',
                            #                   'w', encoding="utf-8") as file:
                            #             file.write(response.text)
                            #         self.api_result.append\
                            #             (api_result_path+'\\'+self.paramread.get_api_sheet_id(self.api_sheet)+'-'+title+'.txt')
                                # 将测试时间存入字典中
                            times = datetime.now()
                            times.strftime('%Y:%m:%d %H:%M:%S')
                            self.api_testtime.append(str(times))
                            print("------------------------------------------------------------------")
                            api_logger.info("------------------------------------------------------------------")
                        self.write_result(r)
                        print('**********测 试 结 束**********')
                        api_logger.info('**********测 试 结 束**********')
        except Exception:
            print(traceback.print_exc())
            api_logger.info(traceback.print_exc())
            print('接口请求错误，请检查接口地址')
            api_logger.info('接口请求错误，请检查接口地址')

    def write_result(self, case_num):
        print(case_num)
        """写入结果"""
        # 获取开始写入行
        pass_num = 0
        failed_num = 0
        test_time = ''
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
            if int(code) == 200:
                self.parseexcelwin.writeCellValue(self.api_sheet, describe_row + row + 1, fruit_col, 'Pass')
                self.parseexcelwin.fontcolor(self.api_sheet, describe_row + row + 1, fruit_col, 4)
                pass_num += 1
            else:
                self.parseexcelwin.writeCellValue(self.api_sheet, describe_row + row + 1, fruit_col, 'Failed')
                self.parseexcelwin.fontcolor(self.api_sheet, describe_row + row + 1, fruit_col, 3)
                failed_num += 1
        # 写入实际结果，如果结果为html，则写入html所保存的地址
        for row, real in enumerate(self.api_result):
            self.parseexcelwin.writeCellValue(self.api_sheet, describe_row+row+1, real_col, str(real))
        # 写入测试时间
        for row, times in enumerate(self.api_testtime):
            self.parseexcelwin.writeCellValue(self.api_sheet, describe_row+row+1, testtime_col, times)
            if row+1 == len(self.api_testtime):
                test_time = times
        # 写入接口用例中的通过与不通过数据
        self.parseexcelwin.writeCellValue('接口用例', case_num+3, API_PASS, pass_num)
        self.parseexcelwin.writeCellValue('接口用例', case_num+3, API_FAIL, failed_num)
        self.parseexcelwin.writeCellValue('接口用例', case_num+3, API_END_TIME, test_time)
        self.parseexcelwin.save()
        self.parseexcelwin.close()


if __name__ == '__main__':
    TestApi().test_api()

