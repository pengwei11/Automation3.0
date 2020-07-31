#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: ParseExcelWin.py
@time: 2020/2/28 14:15
@desc:
"""
from Utils.ConfigRead import *
from Utils.Logger import Logger
from Utils.ParseYaml import ParseYaml
from win32com.client import Dispatch
import xlrd
import win32com.client
import pythoncom

class ParseExcelWin(object):

    '''
    解析EXCEL文档
    '''
    def __init__(self, filename):
        self.filename = filename
        self.parseyaml = ParseYaml()
        pythoncom.CoInitialize()
        self.app = win32com.client.Dispatch('Excel.Application')
        # 读取excel文件
        self.wb = self.app.Workbooks.Open(filename)
        pythoncom.CoInitialize()
        self.logger = Logger('logger', LOGS_PATH).getlog()

    def save(self):
        self.wb.Save()

    def close(self):
        self.wb.Close(SaveChanges=0)
        del self.app

    def getCellValue(self, sheet, row, col):  # 获取单元格的数据
        "Get value of one cell"
        sht = self.wb.Worksheets(sheet)
        return sht.Cells(row, col).Value


    def getColumnValue(self, sheetname, col):
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            info = sheetnames.UsedRange
            nrows = info.Rows.Count
            columnValueList = []
            for i in range(2, nrows+1):
                value = sheetnames.Cells(i, col).Value
                columnValueList.append(value)
            return columnValueList
        except Exception as e:
            print(e)
            self.logger.info('读取错误')

    def getRowValue(self, sheetname, row):
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            info = sheetnames.UsedRange
            ncols = info.Columns.Count
            rowValueList = []
            for i in range(1, ncols+1):
                value = sheetnames.Cells(row, i).Value
                rowValueList.append(value)
            return rowValueList
        except Exception as e:
            print(e)
            self.logger.info('读取错误')

    def getMergeColumnValue(self, sheetname, columnno):
        """
        读取合并单元格的数据
        :param sheetname: 工作表
        :param columnno: 列号
        :return:
        """
        try:
            # 获取数据
            data = xlrd.open_workbook(self.filename)
            # 获取所有sheet名字
            sheet_name = data.sheet_by_name(sheetname)
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
                        col.append(value)
                    else:
                        col.append(sheet_name.cell_value(i, columnno-1))
                return col
        except Exception as e:
            self.logger.info(e)
            self.logger.info('合并单元格读取错误')


    def sheetnames(self):
        """
        获取所有的sheet名称
        """
        sheetcount = self.wb.Sheets.Count
        sheetnames = []
        for i in range(sheetcount):
            sheetnames.append(self.wb.Worksheets(i+1).Name)
        return sheetnames


    def writeCellValue(self, sheetname, rowno, columnno, value):
        """
        向excel某一单元格写入数据
        :param sheetname:
        :param rowno:
        :return:
        """
        sheetnames = self.wb.Worksheets(sheetname)
        sheetnames.Cells(rowno, columnno).Value = value

    def insertRows(self, sheetname, row, num=1):
        """
        row:插入行
        num:插入行数量
        """
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            for i in range(num):
                sheetnames.Rows(row).Insert()
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
            raise
        except Exception as e:
            self.logger.info(e)
            self.logger.info('写入失败，请检查工作表名以及行，列号')

    def insertCols(self, sheetname, col, num=1):
        """
        row:插入列
        num:插入列数量
        """
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            for i in range(num):
                sheetnames.Columns(col).Insert()
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
            raise
        except Exception as e:
            self.logger.info(e)
            self.logger.info('写入失败，请检查工作表名以及行，列号')

    def deleteRows(self, sheetname, row, num=1):
        """
        row:删除行
        num:删除行数量
        """
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            for i in range(num):
                sheetnames.Rows(row).Delete()
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
            raise
        except Exception as e:
            self.logger.info(e)
            self.logger.info('写入失败，请检查工作表名以及行，列号')

    def deleteCols(self, sheetname, col, num=1):
        """
        row:删除列
        num:删除列数量
        """
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            for i in range(num):
                sheetnames.Columns(col).Delete()
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
            raise
        except Exception as e:
            self.logger.info(e)
            self.logger.info('写入失败，请检查工作表名以及行，列号')

    def mergecells(self, sheetname, start_row, start_col, end_row, end_col):
        """
        start_row:合并单元格开始行
        start_col:合并单元格开始列
        start_row:合并单元格结束行
        start_col:合并单元格结束列
        """
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Merge()
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
            raise
        except Exception as e:
            self.logger.info(e)
            self.logger.info('合并失败，请检查工作表名以及行，列号')

    def unmergecells(self, sheetname, start_row, start_col, end_row, end_col):
        """
        start_row:拆分单元格开始行
        start_col:拆分单元格开始列
        start_row:拆分单元格结束行
        start_col:拆分单元格结束列
        """
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            if sheetnames.Cells(start_row, start_col).MergeCells is True:
                sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).UnMerge()
            else:
                pass
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
            raise
        except Exception as e:
            self.logger.info(e)
            self.logger.info('拆分失败，请检查工作表名以及行，列号')


    def borderAround(self, sheetname, start_row, start_col, end_row, end_col):
        """
        start_row:拆分单元格开始行
        start_col:拆分单元格开始列
        start_row:拆分单元格结束行
        start_col:拆分单元格结束列
        """
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            xlContinuous = 1
            xlEdgeLeft = 7
            xlEdgeTop = 8
            xlEdgeBottom = 9
            xlEdgeRight = 10
            xlInsideVertical = 11
            xlInsideHorizontal = 12
            xlThin = 2
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlEdgeLeft).Weight = xlThin
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlEdgeTop).LineStyle = xlContinuous
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlEdgeTop).Weight = xlThin
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlEdgeBottom).Weight = xlThin
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlEdgeRight).LineStyle = xlContinuous
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlEdgeRight).Weight = xlThin
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlInsideVertical).LineStyle = xlContinuous
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlInsideVertical).Weight = xlThin
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).Borders(xlInsideHorizontal).Weight = xlThin
            # sheetnames.Range(sheetnames.Cells(start_row, start_col), sheetnames.Cells(end_row, end_col)).BorderAround(9, 12)
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
            raise
        except Exception as e:
            self.logger.info(e)
            self.logger.info('边框修改失败，请检查工作表名以及行，列号')


    def fontcolor(self, sheetname, start_row, start_col, color=1):
        '''
        设置字体颜色
        '''
        try:
            sheetnames = self.wb.Worksheets(sheetname)
            sheetnames.Cells(start_row, start_col).Font.ColorIndex = color
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
            raise
        except Exception as e:
            self.logger.info(e)
            self.logger.info('字体颜色修改失败，请检查工作表名以及行，列号')


    def clearStepColumnValue(self, sheetname):
        """
        清除测试步骤表中的执行时间，错误结果，错误信息，错误截图信息
        :param sheetname:
        :param columno:
        :return:
        """
        try:
            self.logger.info('清除"%s"工作表测试结果中，请稍等...' % self.wb.Worksheets(sheetname).Name)
            # 获取最大行
            sheetnames = self.wb.Worksheets(sheetname)
            info = sheetnames.UsedRange
            nrows = info.Rows.Count
            # 清除指定列数据
            sheetnames.Range(sheetnames.Cells(2, testStep_EndTime), sheetnames.Cells(nrows, testStep_EndTime)).ClearContents()
            sheetnames.Range(sheetnames.Cells(2, testStep_Result), sheetnames.Cells(nrows, testStep_Result)).ClearContents()
            sheetnames.Range(sheetnames.Cells(2, testStep_Result+1), sheetnames.Cells(nrows, testStep_Result+1)).ClearContents()
            sheetnames.Range(sheetnames.Cells(2, testStep_Result+2), sheetnames.Cells(nrows, testStep_Result+2)).ClearContents()
            sheetnames.Range(sheetnames.Cells(2, testStep_Result+3), sheetnames.Cells(nrows, testStep_Result+3)).ClearContents()
            sheetnames.Range(sheetnames.Cells(2, testStep_Result+4), sheetnames.Cells(nrows, testStep_Result+4)).ClearContents()
            sheetnames.Range(sheetnames.Cells(2, testStep_Error), sheetnames.Cells(nrows, testStep_Error)).ClearContents()
            sheetnames.Range(sheetnames.Cells(2, testStep_Picture), sheetnames.Cells(nrows, testStep_Picture)).ClearContents()
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
        except Exception as e:
            self.logger.info(e)
            self.logger.info('数据清空失败')

    def clearCaseColumnValue(self, sheetname):
        """
        清除执行时间，错误结果
        :param sheetname:
        :param columno:
        :return:
        """
        try:
            self.logger.info('清除"%s"工作表测试结果中，请稍等....' % self.wb.Worksheets(sheetname).Name)
            sheetnames = self.wb.Worksheets(sheetname)
            for i, v in enumerate(self.getColValue(sheetname, testCase_EndTime)):
                if v == '执行结束时间' or sheetnames.Cells(i+2, testCase_EndTime).MergeCells is True or v == '' or v == None:
                    continue
                else:
                    sheetnames.Cells(i+2, testCase_EndTime).Value = ''

            # 清除用例测试结果
            for o in range(5):
                for s, d in enumerate(self.getColumnValue(sheetname, testCase_Result+o)):
                    if d == '执行结果1' or d == '执行结果2' or d == '执行结果3'\
                            or d == '执行结果4' or d == '执行结果5' or d == '' or d == None:
                        continue
                    elif sheetnames.Cells(s+2, testCase_Result+o).MergeCells is True:
                        continue
                    else:
                        sheetnames.Cells(s+2, testCase_Result+o).Value = ''
        except PermissionError:
            self.logger.info('请先关闭用例文件，再运行测试用例')
        except Exception as e:
            self.logger.info(e)
            self.logger.info('数据清空失败')

    def clearKeyWordValue(self, sheetname, row):
        '''
        清空关键字，定位方式，定位元素，value的值
        '''
        self.writeCellValue(sheetname, row, testStep_KeyWord, '')
        self.writeCellValue(sheetname, row, testStep_Location, '')
        self.writeCellValue(sheetname, row, testStep_Locator, '')
        self.writeCellValue(sheetname, row, testStep_Value, '')


if __name__ == '__main__':
    import time
    print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
    p = ParseExcelWin(r'E:\Automation3.0\接口测试模板.xlsx')
    print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
    # p.writeCellValue('创建会议', 1, 1, "{'code': 80001, 'message': '授权账号或密码不对，请再确认清楚', 'result': ''}")
    p.writeCellValue('接口用例', 3, 8, r'=HYPERLINK("E:\promise\机械键盘\1.jpg", "E:\promise\机械键盘\1.jpg")')
    p.save()
    p.wb.Close(SaveChanges=0)
    # p.app.quit()
    print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))