#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/12/28 11:19
当前项目名称  ：AutoTest
功能          ：整点报表-分片线损表
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
class NewtonMethodLosspartHourCalcResult:
    def __init__(self,driver):
        #驱动
        self.driver=driver
        #获取页面元素
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_file=open(self.ProjectFilePath+"\\Page_object\\Data\\NewtonMethodLosspartHourCalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_file)
        self.Data=self.Page_Data['NewtonMethodLosspartHourCalcResult']

        #理论线损ID
        self.LineLoss_ID = self.Data['LineLoss_ID']
        #输电网计算结果查询xpath
        self.WholelossCalcResultQuery_Xpath = self.Data['WholelossCalcResultQuery_Xpath']
        #整点报表xpath
        self.HourCalcResult_Xpath = self.Data['HourCalcResult_Xpath']
        #分片线损表-整点xpath
        self.LossPartResult_Xpath = self.Data['LossPartResult_Xpath']
        #报表iframe的ID
        self.LossPartResultIframe_ID = self.Data['LossPartResultIframe_ID']

        #开始时间窗口xpath
        self.StartTimeWindow_xpath = self.Data['StartTimeWindow_xpath']
        #开始时间日期xpath
        self.StartTimeDate_xpath = self.Data['StartTimeDate_xpath']
        #开始时间确定按钮xpath
        self.StartTimeButton_xpath = self.Data['StartTimeButton_xpath']
        self.StartTimeAdjust_ID = self.Data['StartTimeAdjust_ID']
        #查询按钮ID
        self.Query_ID = self.Data['Query_ID']
        #导出按钮ID
        self.Export_ID = self.Data['Export_ID']

    def NewtonMethodLosspartHourCalcResult_Fun(self,*args):
        #点击理论线损
        self.driver.find_element_by_id(self.LineLoss_ID).click()
        time.sleep(1)
        #点击输电网计算结果查询
        self.driver.find_element_by_xpath(self.WholelossCalcResultQuery_Xpath).click()
        time.sleep(1)
        #点击整点报表
        self.driver.find_element_by_xpath(self.HourCalcResult_Xpath).click()
        time.sleep(1)
        #点击导线损耗表
        self.driver.find_element_by_xpath(self.LossPartResult_Xpath).click()
        time.sleep(5)

        #切换到报表的frame
        self.driver.switch_to_frame(self.LossPartResultIframe_ID)
        #选择开始时间
        self.driver.find_element_by_xpath(self.StartTimeWindow_xpath).click()
        time.sleep(0.2)
        self.driver.find_element_by_xpath(self.StartTimeDate_xpath).click()
        time.sleep(0.2)
        for i in range(0,23):
            self.driver.find_element_by_id(self.StartTimeAdjust_ID).click()
            i=i+1
            time.sleep(0.1)
        time.sleep(0.2)
        self.driver.find_element_by_xpath(self.StartTimeButton_xpath).click()

        #点击查询
        self.driver.find_element_by_id(self.Query_ID).click()
        time.sleep(5)
        #点击导出
        self.driver.find_element_by_id(self.Export_ID).click()
        time.sleep(5)
        self.data_file=u'C:\\Users\\'+GetUsername()+ u'\\Downloads\\全网分片线损报表（小时）.xls'
        self.sheetname=u'全网分片线损报表（小时）'
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        #print self.Excel_Data
        return self.Excel_Data

if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('wjtest','000000')
    time.sleep(5)
    Func=NewtonMethodLosspartHourCalcResult(driver)
    TryReturn = Func.NewtonMethodLosspartHourCalcResult_Fun()
    time.sleep(10)
    driver.quit()