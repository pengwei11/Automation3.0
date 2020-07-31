#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: TestOneCase.py
@time: 2020/3/5 15:44
@desc: 根据给定的测试用例名称，测试单条用例
"""

from Utils.ParseExcelXlrd import ParseExcelXlrd
from Utils.ConfigRead import *
from action.PageAction import PageAction
from selenium.common.exceptions import *   # 导入所有异常类
from Utils.Logger import Logger
from Utils.ParseYaml import ParseYaml
from Utils.WriteFile import YamlWrite
from test.HTMLTestRunner_cn import HTMLTestRunner
from datetime import datetime
from openpyxl.styles import Font
import traceback
import time
import re
import unittest


class TestOneCase(object):

    def __init__(self, filename, sheetname, test_num):
        self.filename = filename
        self.test_num = test_num
        self.sheetname = sheetname
        self.book = ParseExcelXlrd(self.filename)
        self.pre_ifnum = 0
        self.ifnum = 0
        self.pageaction = PageAction()

    def TestCase(self):
        '''
        测试单条用例
        '''
        logger = Logger('case-logger', CASELOGS_PATH).getlog()
        try:
            url = ParseYaml().ReadParameter('IP')
            self.pageaction.openBrowser()
            self.pageaction.getUrl('http://%s' % url)
            CaseList = self.book.getMergeColumnValue(self.sheetname, testStep_Num)
            # 获取用例名称在excel表格中对应的行数
            CaseListIndex = [i for i in range(len(CaseList)) if str(CaseList[i]) == self.test_num]
            # 获取前置条件
            PreNum = self.book.getCellValue(self.sheetname, CaseListIndex[0]+2, testStep_Preset)
            if PreNum != '' and PreNum is not None:
                # 获取前置条件的关键字行数
                PreListIndex = [s for s in range(len(CaseList)) if str(CaseList[s]) == PreNum]
                # 执行前置条件用例
                for i in PreListIndex:
                    print(222)
                    # 获取前置条件关键字等数据
                    # 用例执行步骤
                    pre_stepname = self.book.getCellValue(self.sheetname, i + 2, testStep_Describe)
                    pre_keyword = self.book.getCellValue(self.sheetname, i+2, testStep_KeyWord)
                    if pre_keyword is not None:
                        pre_keyword = pre_keyword.strip()
                    pre_location = self.book.getCellValue(self.sheetname, i+2, testStep_Location)
                    # 去除前后空格
                    if pre_location is not None and type(pre_location) is not int:
                        pre_location = pre_location.strip()
                    pre_locator = self.book.getCellValue(self.sheetname, i+2, testStep_Locator)
                    if type(pre_locator) is int:
                        pre_locator = str(pre_locator)
                    pre_value = self.book.getCellValue(self.sheetname, i+2, testStep_Value)
                    if pre_value is not None and type(pre_value) is not str:
                        pre_value = str(pre_value)

                    '''IF关键字条件判断'''
                    if pre_keyword == 'if':
                        pre_location = str(pre_location)
                        if pre_location == '=':
                            pre_location = '=='
                        if 'byelement' in pre_location or 'iselement' in pre_location:
                            pre_location.replace("’", "'")
                            pre_location.replace("‘", "'")
                            pre_location.replace("（", "(")
                            pre_location.replace("）", ")")
                            pre_location = 'self.pageaction' + '.' + pre_location
                        if 'byelement' in pre_value or 'iselement' in pre_value:
                            pre_value.replace("’", "'")
                            pre_value.replace("‘", "'")
                            pre_value.replace("（", "(")
                            pre_value.replace("）", ")")
                            pre_value = 'self.pagetion' + '.' + pre_value
                        pre_iffun = '%s %s %s' % (pre_location, pre_locator, pre_value)
                        # 条件成立或不成立都跳出此次循环
                        if eval(pre_iffun):
                            logger.info('判断通过，执行判断内操作')
                            continue
                        else:
                            logger.info('判断失败，条件不成立')
                            self.pre_ifnum += 1
                            continue
                    if self.pre_ifnum != 0:
                        # 判断关键字是否为break，如果不为break则跳出循环
                        if pre_keyword != 'break':
                            continue
                    if pre_keyword == 'break':
                        logger.info('判断结束')
                        # 初始化if失败次数记录
                        self.pre_ifnum = 0
                        continue
                    if pre_keyword and pre_location and pre_locator and pre_value:
                        pre_fun = 'self.pageaction' + '.' + pre_keyword + '(' + '"' + pre_location + '"' + ', ' + '"' + pre_locator + '"' + ', ' + '"' + \
                                  pre_value + '"' + ')'
                    elif pre_keyword and pre_value and (pre_location is None or pre_location == '') \
                            and (pre_locator is None or pre_locator == ''):
                        pre_fun = 'self.pageaction' + '.' + pre_keyword + '(' + '"' + pre_value + '"' + ')'
                    elif pre_keyword and pre_location and pre_locator and (pre_value is None or pre_value == ''):
                        pre_fun = 'self.pageaction' + '.' + pre_keyword + '(' + '"' + pre_location + '"' + ', ' + '"' + pre_locator + '"' + ')'
                    elif pre_keyword and (pre_location is None or pre_location == '') and (pre_locator is None or pre_locator == '') and (pre_value is None or pre_value == ''):
                        pre_fun = 'self.pageaction' + '.' + pre_keyword + '(' + ')'
                    elif (pre_keyword is None or pre_keyword == '') and (pre_location is None or pre_location == '') \
                            and (pre_locator is None or pre_locator == '') and (pre_value is None or pre_value == ''):
                        continue
                    else:
                        logger.info('关键字对应参数错误')
                        continue
                    try:
                        # eval 将字符串转换为可执行的python语句
                        eval(pre_fun)
                    except TypeError:
                        logger.info('步骤"{}"执行失败'.format(pre_stepname))
                        logger.info('关键字参数个数错误，请检查参数')
                    except TimeoutException:
                        # 写入测试时间，测试结果，错误信息，错误截图
                        logger.info('步骤"{}"执行失败'.format(pre_stepname))
                        logger.info('元素定位超时，请检查上一步是否执行成功，或元素定位方式')
                    except TimeoutError as e:
                        logger.info('步骤"{}"执行失败'.format(pre_stepname))
                        logger.info('元素查找超时，请检查上一步是否执行成功，或元素定位方式')
                    except AttributeError as e:
                        logger.info('步骤"{}"执行失败'.format(pre_stepname))
                        logger.info(e)
                    except AssertionError:
                        logger.info('步骤"{}"执行失败'.format(pre_stepname))
                        print('步骤"{}"执行失败'.format(pre_stepname))
                    except WebDriverException:
                        logger.info('步骤"{}"执行失败'.format(pre_stepname))
                        logger.info('浏览器异常，请检查浏览器驱动或运行过程中是否被强制关闭')
                    except Exception:
                        error_info = traceback.format_exc()
                        # # 写入测试时间，测试结果，错误信息，错误截图
                        logger.info('步骤"{}"执行失败'.format(pre_stepname))
                        logger.info(error_info)
                    else:
                        logger.info('步骤"{}"执行成功'.format(pre_stepname))
                        print('步骤"{}"执行成功'.format(pre_stepname))

            # 执行用例
            for i in CaseListIndex:
                # 用例执行步骤
                stepname = self.book.getCellValue(self.sheetname, i + 2, testStep_Describe)
                # 获取关键字
                keyword = self.book.getCellValue(self.sheetname, i + 2, testStep_KeyWord)
                # 去除前后空格
                if keyword is not None:
                    keyword = keyword.strip()
                # 获取定位方式
                location = self.book.getCellValue(self.sheetname, i + 2, testStep_Location)
                # 去除前后空格
                if location is not None and type(location) is not int:
                    location = location.strip()
                # 获取定位表达式
                locator = self.book.getCellValue(self.sheetname, i + 2, testStep_Locator)
                if type(locator) is int:
                    locator = str(locator)
                # 获取输入值
                testvalue = self.book.getCellValue(self.sheetname, i + 2, testStep_Value)
                # 如果输入值为 int 类型，则强转为 str 类型，用于字符串拼接
                if testvalue is not None and type(testvalue) is not str:
                    testvalue = str(testvalue)

                '''IF关键字条件判断'''
                if keyword == 'if':
                    location = str(location)
                    if locator == '=':
                        locator = '=='
                    if 'byelement' in location or 'iselement' in location:
                        location.replace("’", "'")
                        location.replace("‘", "'")
                        location.replace("（", "(")
                        location.replace("）", ")")
                        location = 'self.pageaction' + '.' + location
                    if 'byelement' in testvalue or 'iselement' in testvalue:
                        testvalue.replace("’", "'")
                        testvalue.replace("‘", "'")
                        testvalue.replace("（", "(")
                        testvalue.replace("）", ")")
                        testvalue = 'self.pagetion' + '.' + testvalue
                    iffun = '%s %s %s' % (location, locator, testvalue)
                    # 条件成立或不成立都跳出此次循环
                    if eval(iffun):
                        logger.info('判断通过，执行判断内操作')
                        continue
                    else:
                        logger.info('判断失败，条件不成立')
                        # 如果不成立，则需要一直跳出循环，直至break关键字出现
                        self.ifnum += 1
                        continue
                if self.ifnum != 0:
                    # 判断关键字是否为break，如果不为break则跳出循环
                    if keyword != 'break':
                        continue
                if keyword == 'break':
                    logger.info('判断结束')
                    # 初始化if失败次数记录
                    self.ifnum = 0
                    continue
                if keyword and location and locator and testvalue:
                    print(111)
                    fun = 'self.pageaction' + '.' + keyword + '(' + '"' + location + '"' + ', ' + '"' + locator + '"' + ', ' + '"' + \
                          testvalue + '"' + ')'
                elif keyword and testvalue and (location is None or location == '') \
                        and (locator is None or locator == ''):
                    fun = 'self.pageaction' + '.' + keyword + '(' + '"' + testvalue + '"' + ')'
                elif keyword and location and locator and (testvalue is None or testvalue == ''):
                    fun = 'self.pageaction' + '.' + keyword + '(' + '"' + location + '"' + ', ' + '"' + locator + '"' + ')'
                elif keyword and (location is None or location == '') and (locator is None or locator == '') and (testvalue is None or testvalue == ''):
                    fun = 'self.pageaction' + '.' + keyword + '(' + ')'
                elif (keyword is None or keyword == '') and (location is None or location == '') \
                        and (locator is None or locator == '') and (testvalue is None or testvalue == ''):
                    continue
                else:
                    logger.info('关键字对应参数错误')
                    continue
                try:
                    # eval 将字符串转换为可执行的python语句
                    eval(fun)
                # 抛出异常的情况，将失败结果写入excel表格中
                except TypeError as e:
                    logger.info('步骤"{}"执行失败'.format(stepname))
                    logger.info('关键字参数个数错误，请检查参数')
                    logger.info(e)
                except TimeoutException as e:
                    logger.info('步骤"{}"执行失败'.format(stepname))
                    logger.info('元素定位超时，请检查上一步是否执行成功，或元素定位方式')
                    logger.info(e)
                except TimeoutError as e:
                    logger.info('步骤"{}"执行失败'.format(stepname))
                    logger.info(e)
                except AttributeError as e:
                    logger.info('步骤"{}"执行失败'.format(stepname))
                    logger.info(e)
                except AssertionError as e:
                    logger.info('步骤"{}"执行失败'.format(stepname))
                    logger.info(e)
                except WebDriverException as e:
                    # 写入测试时间，测试结果，错误信息，错误截图
                    logger.info('步骤"{}"执行失败'.format(stepname))
                    logger.info('浏览器异常，请检查浏览器驱动或运行过程中是否被强制关闭')
                except Exception:
                    error_info = traceback.format_exc()
                    logger.info('步骤"{}"执行失败'.format(stepname))
                    logger.info(error_info)
                else:
                    logger.info('步骤"{}"执行成功'.format(stepname))
                    print('步骤"{}"执行成功'.format(stepname))
        except Exception as e:
            logger.info('用例执行失败，请检查后重试')
            logger.info(e)
        finally:
            self.pageaction.quitBrowser()

if __name__ == '__main__':
    TestOneCase(r'E:\Automation3.0\百度测试用例.xlsx', '测试', 'test_ceshi_2').TestCase()