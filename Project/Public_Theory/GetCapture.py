#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/8 16:51
当前项目名称  ：620Calc
功能          ：出现错误是进行截图，并将截图命名放到测试工程错误文件夹下
-------------------------------------漂亮的分割线----------------------------------------------'''
from Public_Theory import GetProjectFilePath
import time
def GetCapture(driver):
    ProjectFilePath = GetProjectFilePath.GetProjectFilePath()
    #self.driver = webdriver.Chrome()
    ErrorCapturePath = ProjectFilePath+'\EagleProjecttest\Error\\'
    now = time.strftime("%Y-%m-%d_%H_%M_%S",time.localtime())
    driver.get_screenshot_as_file(ErrorCapturePath+now+".png")