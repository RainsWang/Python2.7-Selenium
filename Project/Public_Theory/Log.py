#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/8 17:08
当前项目名称  ：620Calc
功能          ：输出自定义的日志信息
-------------------------------------漂亮的分割线----------------------------------------------'''
import logging,time
from Public_Theory.GetProjectFilePath import GetProjectFilePath

class LogOutput:
    def __init__(self,title):
        self.ProjectFilePath = GetProjectFilePath()
        self.day = time.strftime("%Y%m%d",time.localtime())
        self.logger = logging.getLogger(title)
        self.logger.setLevel(logging.INFO)
        if not self.logger.handlers:
            self.formatter = logging.Formatter('%(asctime)s %(levelname)-8s: %(message)s')
            self.file_handler = logging.FileHandler(self.ProjectFilePath + '\EagleProjecttest\Log\%s.log' % self.day)
            self.file_handler.setFormatter(self.formatter)
            self.control = logging.StreamHandler()
            self.control.setFormatter(self.formatter)
            self.logger.addHandler(self.file_handler)
            self.logger.addHandler(self.control)
    def debug_log(self,msg):
        self.logger.debug(msg)
    def info_log(self,msg):
        self.logger.info(msg)
    def warning_log(self,msg):
        self.logger.warning(msg)
    def error_log(self,msg):
        self.logger.error(msg)
'''
Try = LogOutput('尝试输出')
Try.info_log('This is an output message!!!')'''