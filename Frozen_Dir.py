#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: Frozen_Dir.py
@time: 2019/11/14 10:56
@desc:
"""
import sys
import os


def app_path():
    if hasattr(sys, 'frozen'):
        if 'dist' in os.path.dirname(sys.executable):
            return os.path.dirname(sys.executable)[0:os.path.dirname(sys.executable).rfind('dist', 1)-1]
        else:
            return os.path.dirname(sys.executable)
    return os.path.dirname(__file__)
