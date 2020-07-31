#!/usr/bin/env python
# encoding: utf-8
"""
@author: caopeng
@license: (C) Copyright 2013-2017, Node Supply Chain Manager Corporation Limited.
@contact: 1249294960@qq.com
@software: pengwei
@file: PageAction.py
@time: 2019/11/8 9:32
@desc:
"""
from Utils.Logger import Logger
from Utils.ConfigRead import *
import requests
import json
import ast
import traceback



class RunMethod(object):

    def __init__(self):
        self.api_logger = Logger('api-logger', APILOGS_PATH).getlog()

    """
    封装测试接口类型（Get, POST, DELETE）
    """
    def get_main(self, url, data=None, header=None):
        res = None
        if header:
            res = requests.get(url=url, params=data, headers=header)
        else:
            res = requests.get(url=url, params=data)
        self.api_logger.info('发送get请求')
        return res

    def post_main(self, url, data=None, header=None):
        """
        注：
        情况一：post带json参数      N
        情况二：post带data参数      Y
        情况三：post带单文件上传 {"file": open('test.txt', 'rb')} files = file   N
        情况四：post带多文件上传 file = [
                                ('file1',('test.txt',open('test.txt', 'rb'))),
                                ('file2', ('test.png', open('test.png', 'rb')))
                                ]
                                files = file    N
        """
        res = None
        try:
            if header:
                res = requests.post(url=url, data=data, headers=header)
            else:
                res = requests.post(url=url, data=data)
            return res
        except Exception:
            return res
        finally:
            self.api_logger.info('发送post请求')

    def delete_main(self, url, data, header=None):
        res = None
        try:
            if header:
                res = requests.delete(url=url, params=data, headers=header)
            else:
                res = requests.delete(url=url, params=data)
            return res
        except Exception:
            return res
        finally:
            self.api_logger.info('发送delete请求')

    def put_main(self, url, data, header=None):
        res = None
        try:
            if header:
                res = requests.put(url=url, params=data, headers=header)
            else:
                res = requests.put(url=url, params=data)
            return res
        except Exception:
            return res
        finally:
            self.api_logger.info('发送put请求')

if __name__ == '__main__':
    # url = 'http://172.16.45.15/api/login/login'
    url = 'http://172.16.45.15/appapi/video/video_note_list'
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': 'Bearer mmySNbD0ZDDT7b5PyFM3YByUInyfALrG',
        'token': 'xptnoODXnWebVqypm6ial1qgoXFibKWcmnFT2ciqq8vUz57Kh5%2FWa5RmcVhnaGNlbcqaZ2hlk5TCbGrLl2uelcbEbJebyZdikZqda1hxrg%3D%3D'
    }
    # param = {
    #     'user_id': 2,
    #     'res_id': 39,
    #     'origin': 0,
    #     'score':10,
    #     'content':'测试评论',
    # }
    param = {"videoid": "39", 'userid': ''}

    # print(param==d)
    # param = {"m_id": 124, "u_id": 260, "username": "hgySB", "unit": "ITC", "dept": "研发测试部", "position": "软件测试工程师", "salutatory": "欢迎ITC指导工作", "t_id": 1, "is_broadcast": 1, "company": "BL", "device_name": "8501R-V1.2"}
    res = RunMethod().get_main(url, param, header=headers)
    # res = requests.post(url, data=param, headers=headers)
    # print(res)
    # print(res.status_code)
    # s = res.json()['data']
    # d = s['token']
    # print(res.json())

    print(res.json())