#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: ParseYaml.py
@time: 2019/11/14 16:39
@desc: 读取Yaml文件
"""

from Utils.ReadFile import YamlRead
from Utils.ConfigRead import *


class ParseYaml(object):

    PARAMETER_PATH = CONFIG_PATH+'Parameter.yaml'
    GUISELECTVALUE_PATH = CONFIG_PATH+'GuiSelectValue.yaml'
    TIMEWAIT_PATH = CONFIG_PATH+'TimeWait.yaml'
    API_PARAMETER_PATH = CONFIG_PATH+'interface_config\\'+'API_Parameter.yaml'

    def __init__(self, parameter=PARAMETER_PATH, guiselectvalue=GUISELECTVALUE_PATH, timewait=TIMEWAIT_PATH, api_parameter=API_PARAMETER_PATH):
        self.parameter = YamlRead(parameter).data
        self.guiselectvalue = YamlRead(guiselectvalue).data
        self.timewait = YamlRead(timewait).data
        self.api_parameter = YamlRead(api_parameter).data

    def ReadParameter(self, element, index=0):
        return self.parameter[index].get(element)

    def ReadGuiSelectValue(self, element, index=0):
        return self.guiselectvalue[index].get(element)

    def ReadTimeWait(self, element, index=0):
        return self.timewait[index].get(element)

    def ReadAPI_Paramter(self, element, index = 0):
        return self.api_parameter[index].get(element)


if __name__ == '__main__':
    # print(ParseYaml().ReadGuiSelectValue('BrowserType').get('Chrome'))
    # print(ParseYaml().ReadParameter('Moudle'))
    print(ParseYaml().ReadAPI_Paramter('FILE_PATH'))
    # print(type(ParseYaml().ReadTimeWait('casetime')))
    # print(bytes("查询离线会议室服务器"))

