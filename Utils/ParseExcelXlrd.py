#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: ParseExcelXlrd.py
@time: 2020/1/13 10:48
@desc: 更换读取库为xlrd
"""

from Utils.Logger import Logger
from Utils.ConfigRead import *
from xlrd import xldate_as_tuple
from datetime import datetime
import traceback
import xlrd



class ParseExcelXlrd(object):

    def __init__(self, filename):
        # 读取excel文件地址
        self.filename = filename
        # 读取excel文件
        self.wookbook = xlrd.open_workbook(self.filename)
        self.logger = Logger('logger', LOGS_PATH).getlog()

    def getRowValue(self, sheetname, rowno):
        """
        获取excel某一行的数据
        :param sheetname:
        :param rowno:
        :return:
        """
        try:
            if type(sheetname) is int:
                sheetnames = self.wookbook.sheet_by_index(sheetname)
            else:
                # 获取sheetname对象
                sheetnames = self.wookbook.sheet_by_name(sheetname)
            row_values = []
            for i, v in enumerate(sheetnames.row_values(rowno-1)):
                CellType = sheetnames.cell(rowno - 1, i).ctype
                CellValue = sheetnames.cell_value(rowno - 1, i)
                if CellType == 4:
                    if CellValue == 1:
                        row_values.append('True')
                    elif CellValue == 0:
                        row_values.append('False')
                elif CellType == 3:
                    date = datetime(*xldate_as_tuple(CellValue, 0))
                    CellValue = date.strftime('%Y-%m-%d')
                    row_values.append(CellValue)
                elif type(CellValue) is float:
                    row_values.append(int(CellValue))
                else:
                    row_values.append(str(CellValue))
            return row_values
        except Exception as e:
            self.logger.info(traceback.print_exc())
            self.logger.info('读取失败，请检查工作表名以及行，列号')

    def getColumnValue(self, sheetname, columnno):
        '''
        获取excel某一列的数据
        :param sheetname:
        :param rowno:
        :return:
        '''
        try:
            if type(sheetname) is int:
                sheetnames = self.wookbook.sheet_by_index(sheetname)
            else:
                sheetnames = self.wookbook.sheet_by_name(sheetname)
            return sheetnames.col_values(columnno-1)
        except Exception as e:
            self.logger.info(traceback.print_exc())
            self.logger.info('读取失败，请检查工作表名以及行，列号')

    def getMergeColumnValue(self, sheetname, columnno):
        """
        读取合并单元格的数据
        :param sheetname: 工作表
        :param columnno: 列号
        :return:
        """
        try:

            # 获取所有sheet名字
            sheet_name = self.wookbook.sheet_by_name(sheetname)
            # 获取总行数
            nrows = sheet_name.nrows  # 包括标题
            # 获取总列数
            ncols = sheet_name.ncols
            # 计算出合并的单元格有哪些
            colspan = {}
            # 如果sheet是合并的单元格 则获取合并单元格的值，并将第一行的数据赋值给合并单元格中的空值
            if sheet_name.merged_cells:
                for item in sheet_name.merged_cells:
                    for row in range(item[0], item[1]):
                        for col in range(item[2], item[3]):
                            # 合并单元格的首格是有值的，所以在这里进行了去重
                            if (row, col) != (item[0], item[2]):
                                colspan.update({(row, col): (item[0], item[2])})

                col = []
                for i in range(1, nrows):
                    if colspan.get((i, columnno-1)):
                        value = sheet_name.cell_value(*colspan.get((i, columnno-1)))
                        if type(value) is float:
                            value = str(int(value))
                        col.append(value)
                    else:
                        value1 = sheet_name.cell_value(i, columnno-1)
                        if type(sheet_name.cell_value(i, columnno-1)) is float:
                            value1 = str(int(sheet_name.cell_value(i, columnno-1)))
                        col.append(value1)
                return col
        except Exception as e:
            self.logger.info(e)
            self.logger.info('合并单元格读取错误')

    def getCellValue(self, sheetname, rowno, columnno):
        """
        获取excel某一单元格的数据
        :param sheetname:
        :param rowno:
        :return:
        """
        try:
            if type(sheetname) is int:
                sheetnames = self.wookbook.sheet_by_index(sheetname)
            else:
                sheetnames = self.wookbook.sheet_by_name(sheetname)
            CellValue = sheetnames.cell_value(rowno-1, columnno-1)
            CellType = sheetnames.cell(rowno-1, columnno-1).ctype
            if CellType == 4:
                if CellValue == 1:
                    return 'True'
                elif CellValue == 0:
                    return 'False'
            elif CellType == 3:
                date = datetime(*xldate_as_tuple(CellValue, 0))
                CellValue = date.strftime('%Y-%m-%d')
                return CellValue
            elif type(CellValue) is float:
                return int(CellValue)
            else:
                return str(CellValue)
        except Exception as e:
            self.logger.info(traceback.print_exc())
            self.logger.info('读取失败，请检查工作表名以及行，列号')

    def ismerge(self, sheetname):
        """
        判断'工作表'内是否有合并单元格
        :param sheetname:
        :return:
        """
        sheetnames = self.wookbook.sheet_by_name(sheetname)
        merge = sheetnames.merged_cells
        return merge




if __name__ == '__main__':
    import time
    print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
    w = ParseExcelXlrd(r'E:\智慧校园接口测试用例\0登录.xlsx')
    # templater = ['序号', '接口地址', '接口工作表', '请求类型', '是否执行', '是否生成用例', '执行结束时间', '通过', '失败']
    # print(w.getRowValue(0, 2))
    # print(templater == w.getRowValue(0, 2))
    print(w.getCellValue('登录', 1, 2))
    # print(w.wookbook.)
    # print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
    # # print(w.getMergeColumnValue('搜索', 1))
    # # print(w.wookbook.sheet_by_name('测试').nrows)
    # print(w.getRowValue('查询离线会议室服务器', 16))
    # print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))

