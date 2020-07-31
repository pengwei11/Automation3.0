#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: WriteFile.py
@time: 2019/11/14 16:58
@desc:
"""
import os
from ruamel import yaml
from Utils.ConfigRead import *


class YamlWrite(object):

    def __init__(self, filename):
        self.filename = filename

    def Write_Yaml(self, value):
        try:
            if self.filename in '\\':
                self.filename.replace('\\', '/')
                if not os.path.exists(self.filename):
                    os.system(r'type nul>{}'.format(self.filename))
                    # logger.info('新建文件：%s'%filename)
        finally:
            with open(self.filename, 'w+', encoding='utf-8') as f:
                yaml.dump(value, f, Dumper=yaml.RoundTripDumper)
                f.close()

    # 追加写入
    def Write_Yaml_A(self, value):
        try:
            if self.filename in '\\':
                self.filename.replace('\\', '/')
                if not os.path.exists(self.filename):
                    os.system(r'type nul>{}'.format(self.filename)) 
                    # logger.info('新建文件：%s'%filename)
        finally:
            with open(self.filename, 'a', encoding='utf-8') as f:
                yaml.dump(value, f, Dumper=yaml.RoundTripDumper)
                f.close()

    # 修改参数
    def Write_Yaml_Updata(self, key, value):
        with open(self.filename) as f:
            content = yaml.safe_load(f)
            content[key] = value
            f.close()
        with open(self.filename, 'w+', encoding='utf-8') as f:
            yaml.dump(content, f, Dumper=yaml.RoundTripDumper)
            f.close()


if __name__ == '__main__':
    import time
    p = r'E:\Automation3.0\config\interface_config\API_Parameter.yaml'
    # YamlWrite(p).Write_Yaml_Updata('MOUDEL', '获取token')
    YamlWrite(p).Write_Yaml_Updata('FILE_PATH', ['E:/接口测试用例/ce/会议室接口.xlsx'])