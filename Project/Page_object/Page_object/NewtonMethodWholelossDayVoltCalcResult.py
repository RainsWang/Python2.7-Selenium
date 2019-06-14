#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/11/1 14:46
当前项目名称  ：AutoTest
功能          ：日报表-输电网分压线损表
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
class NewtonMethodWholelossDayVoltCalcResult:
    def __init__(self,driver):
        #驱动
        self.driver=driver
        #获取页面元素
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_file=open(self.ProjectFilePath+"\\Page_object\\Data\\NewtonMethodWholelossDayVoltCalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_file)
        self.Data=self.Page_Data['NewtonMethodWholelossDayVoltCalcResult']

        #理论线损ID
        self.LineLoss_ID = self.Data['LineLoss_ID']
        #输电网计算结果查询xpath
        self.WholelossCalcResultQuery_Xpath = self.Data['WholelossCalcResultQuery_Xpath']
        #日报表xpath
        self.DayCalcResult_Xpath = self.Data['DayCalcResult_Xpath']
        #输电网分压线损表xpath
        self.WholelossDayVoltCalcResult_Xpath = self.Data['WholelossDayVoltCalcResult_Xpath']
        #报表iframe的ID
        self.WholelossDayVoltCalcResultIframe_ID = self.Data['WholelossDayVoltCalcResultIframe_ID']
        #导出按钮ID
        self.Export_ID = self.Data['Export_ID']

    def NewtonMethodWholelossDayVoltCalcResult_Fun(self,*args):
        #点击理论线损
        self.driver.find_element_by_id(self.LineLoss_ID).click()
        time.sleep(1)
        #点击输电网计算结果查询
        self.driver.find_element_by_xpath(self.WholelossCalcResultQuery_Xpath).click()
        time.sleep(1)
        #点击日报表
        self.driver.find_element_by_xpath(self.DayCalcResult_Xpath).click()
        time.sleep(1)
        #点击输电网分压线损表
        self.driver.find_element_by_xpath(self.WholelossDayVoltCalcResult_Xpath).click()
        time.sleep(5)

        #切换到报表的frame
        self.driver.switch_to_frame(self.WholelossDayVoltCalcResultIframe_ID)
        #点击导出
        self.driver.find_element_by_id(self.Export_ID).click()
        time.sleep(5)
        self.data_file=u'C:\\Users\\'+GetUsername()+ u'\\Downloads\\全网分电压损耗报表（日）.xls'
        self.sheetname=u'全网分电压损耗报表（日）'
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        #print self.Excel_Data
        return self.Excel_Data

if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('wjtest','000000')
    time.sleep(5)
    Func=NewtonMethodWholelossDayVoltCalcResult(driver)
    TryReturn = Func.NewtonMethodWholelossDayVoltCalcResult_Fun()
    time.sleep(10)
    driver.quit()