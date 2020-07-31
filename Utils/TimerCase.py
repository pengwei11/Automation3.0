#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: WriteFile.py
@time: 2020/3/8 16:58
@desc: 计时器
"""

class TimerCase(object):

	def __init__(self):
		self.Hour = 0
		self.Minute = 0
		self.Sceond = 0

	def Timer(self):
		"""
		计时器
		"""
		self.Sceond = self.Sceond+1
		if self.Sceond == 60:
			self.Minute = self.Minute+1
			self.Sceond = 0
		if self.Minute == 60:
			self.Hour = self.Hour+1
			self.Minute = 0
		if self.Minute < 10:
			times = str(self.Hour) + ':0' + str(self.Minute) + ':' + str(self.Sceond)
		else:
			times = str(self.Hour) + ':' + str(self.Minute) + ':' + str(self.Sceond)
		return times
