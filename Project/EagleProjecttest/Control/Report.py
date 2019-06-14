#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/16 20:46
当前项目名称  ：620Calc
功能          ：读取case，并生成相应的报告
-------------------------------------漂亮的分割线----------------------------------------------'''
import HTMLTestRunner
import time
import unittest
from Public_Theory.GetProjectFilePath import GetProjectFilePath
import os


def ReportHtml():
    ProjectFilePath=GetProjectFilePath()
    #创建一个测试用例集
    suit=unittest.TestSuite()
    #获取当前时间并生成相关格式进行赋值操作
    now=time.strftime("%Y-%m-%d_%H_%M_%S",time.localtime())
    #产生报告路径
    Reportfile=ProjectFilePath+"\\EagleProjecttest\Report\\"+now+".html"
    #case 路径
    casefile=ProjectFilePath+"\\EagleProjecttest\case"
    #获取到casefile所有的文件内容
    CaseFileMsg = os.listdir(casefile)
    discover = unittest.defaultTestLoader.discover(casefile, pattern="*.py", top_level_dir=None)
    # 将测试用例方法添加到测试集中
    for test in discover:
        for testcase in test:
            suit.addTest(testcase)
    Re_Openfile=file(Reportfile,"wb")
    #加载HTMLTestRunner方法，将数据流写入对应文件并写明title和descripion
    runner=HTMLTestRunner.HTMLTestRunner(stream=Re_Openfile,title=u'eagle2理论线损测试报告',description=u'eagle2理论线损测试报告')
    #运行测试用例集
    runner.run(suit)
if __name__=="__main__":
    ReportHtml()
