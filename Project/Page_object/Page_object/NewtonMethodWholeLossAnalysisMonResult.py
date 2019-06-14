#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2019/1/3 19:21
当前项目名称  ：AutoTest
功能          ：月分析-输电网综合分析
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
from Public_Theory.GetUsername import GetUsername
import yaml
import time
import Login
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class NewtonMethodWholeLossAnalysisMonResult:
    def __init__(self,driver):
        #驱动
        self.driver=driver
        #获取页面元素
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_file=open(self.ProjectFilePath+"\\Page_object\\Data\\NewtonMethodWholeLossAnalysisMonResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_file)
        self.Data=self.Page_Data['NewtonMethodWholeLossAnalysisMonResult']

        #线损分析ID
        self.LineLossAnalysis_ID = self.Data['LineLossAnalysis_ID']
        #输电网结果分析ID
        self.WholeLossResultAnalysis_xpath = self.Data['WholeLossResultAnalysis_xpath']
        #月分析ID
        self.MonAnalysis_xpath = self.Data['MonAnalysis_xpath']
        #输电网综合分析xpath
        self.WholeLossAnalysis_xpath = self.Data['WholeLossAnalysis_xpath']

        #报表iframe的ID
        self.WholelossAnalysisIframe_ID = self.Data['WholelossAnalysisIframe_ID']
        #激活月份选择框
        self.Month_ID = self.Data['Month_ID']
        #选择2015
        self.Year_xpath = self.Data['Year_xpath']
        #选择八月
        self.Month_xpath = self.Data['Month_xpath']
        #确定按钮
        self.ConfirmButton_ID = self.Data['ConfirmButton_ID']

        #查询按钮ID
        self.Query_ID = self.Data['Query_ID']
        #导出按钮ID
        self.Export_ID = self.Data['Export_ID']

    def NewtonMethodWholeLossAnalysisMonResult_Fun(self,*args):
        #点击线损分析
        self.driver.find_element_by_id(self.LineLossAnalysis_ID).click()
        time.sleep(1)
        #点击输电网结果分析
        self.driver.find_element_by_xpath(self.WholeLossResultAnalysis_xpath).click()
        time.sleep(1)
        #点击月分析
        self.driver.find_element_by_xpath(self.MonAnalysis_xpath).click()
        time.sleep(1)
        #点击输电网综合分析
        self.driver.find_element_by_xpath(self.WholeLossAnalysis_xpath).click()
        time.sleep(5)

        #切换到报表的frame
        self.driver.switch_to_frame(self.WholelossAnalysisIframe_ID)
        time.sleep(1)
        #激活月份选择框
        self.driver.find_element_by_id(self.Month_ID).click()
        time.sleep(1)
        #选择2015
        self.driver.find_element_by_xpath(self.Year_xpath).click()
        time.sleep(1)
        #选择八月
        self.driver.find_element_by_xpath(self.Month_xpath).click()
        time.sleep(1)
        #点击确定
        self.driver.find_element_by_id(self.ConfirmButton_ID).click()
        time.sleep(1)

        #点击查询
        self.driver.find_element_by_id(self.Query_ID).click()
        time.sleep(5)
        #点击导出
        self.driver.find_element_by_id(self.Export_ID).click()
        time.sleep(5)
        self.data_file=u'C:\\Users\\'+GetUsername()+ u'\\Downloads\\输电网综合分析（月）.xls'
        self.sheetname=u'输电网综合分析（月）'
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        #print self.Excel_Data
        return self.Excel_Data

if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('wjtest','000000')
    time.sleep(5)
    Func=NewtonMethodWholeLossAnalysisMonResult(driver)
    TryReturn = Func.NewtonMethodWholeLossAnalysisMonResult_Fun()
    time.sleep(10)
    driver.quit()