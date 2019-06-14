#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/11/1 14:46
当前项目名称  ：AutoTest
功能          ：日报表-输电网线损表
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
class NewtonMethodWholelossDayCalcResult:
    def __init__(self,driver):
        #驱动
        self.driver=driver
        #获取页面元素
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_file=open(self.ProjectFilePath+"\\Page_object\\Data\\NewtonMethodWholelossDayCalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_file)
        self.Data=self.Page_Data['NewtonMethodWholelossDayCalcResult']

        #理论线损ID
        self.LineLoss_ID = self.Data['LineLoss_ID']
        #输电网计算结果查询xpath
        self.WholelossCalcResultQuery_Xpath = self.Data['WholelossCalcResultQuery_Xpath']
        #日报表xpath
        self.DayCalcResult_Xpath = self.Data['DayCalcResult_Xpath']
        #输电网线损表xpath
        self.WholelossCalcResult_Xpath = self.Data['WholelossCalcResult_Xpath']
        #报表iframe的ID
        self.WholelossCalcResultIframe_ID = self.Data['WholelossCalcResultIframe_ID']
        #屏蔽一下代码原因：日期选择框的Iframe的ID和name都是随机的，无法进行定位，使用代表日进行查询和导出
        # #激活开始时间窗口
        # self.StartTimeWindow = self.Data['StartTimeWindow']
        # #开始时间窗口xpath
        # self.StartTimeWindow_ID = self.Data['StartTimeWindow_ID']
        # #开始时间日期xpath
        # self.StartTimeDate_xpath = self.Data['StartTimeDate_xpath']
        # #开始时间确定按钮xpath
        # self.StartTimeButton_xpath = self.Data['StartTimeButton_xpath']
        #
        # #结束时间窗口xpath
        # self.EndTimeWindow_xpath = self.Data['EndTimeWindow_xpath']
        # #结束时间日期xpath
        # self.EndTimeDate_xpath = self.Data['EndTimeDate_xpath']
        # #结束时间确定按钮xpath
        # self.EndTimeButton_xpath = self.Data['EndTimeButton_xpath']
        #查询按钮ID
        self.Query_ID = self.Data['Query_ID']
        #导出按钮ID
        self.Export_ID = self.Data['Export_ID']

    def NewtonMethodWholelossDayCalcResult_Fun(self,*args):
        #点击理论线损
        self.driver.find_element_by_id(self.LineLoss_ID).click()
        time.sleep(1)
        #点击输电网计算结果查询
        self.driver.find_element_by_xpath(self.WholelossCalcResultQuery_Xpath).click()
        time.sleep(1)
        #点击日报表
        self.driver.find_element_by_xpath(self.DayCalcResult_Xpath).click()
        time.sleep(1)
        #点击输电网线损表
        self.driver.find_element_by_xpath(self.WholelossCalcResult_Xpath).click()
        time.sleep(5)

        #切换到报表的frame
        self.driver.switch_to_frame(self.WholelossCalcResultIframe_ID)
        time.sleep(1)
        # #选择开始时间
        # self.driver.find_element_by_id(self.StartTimeWindow).click()
        # time.sleep(0.2)
        # self.driver.find_element_by_id(self.StartTimeWindow_ID).click()
        # self.driver.find_element_by_xpath(self.StartTimeDate_xpath).click()
        # time.sleep(0.2)
        # self.driver.find_element_by_xpath(self.StartTimeButton_xpath).click()
        # time.sleep(0.2)
        #
        # #选择结束时间
        # self.driver.find_element_by_xpath(self.EndTimeWindow_xpath).click()
        # time.sleep(0.2)
        # self.driver.find_element_by_xpath(self.EndTimeDate_xpath).click()
        # time.sleep(0.2)
        # self.driver.find_element_by_xpath(self.EndTimeButton_xpath).click()
        # time.sleep(0.2)
        #点击查询
        self.driver.find_element_by_id(self.Query_ID).click()
        time.sleep(5)
        #点击导出
        self.driver.find_element_by_id(self.Export_ID).click()
        time.sleep(5)
        self.data_file=u'C:\\Users\\'+GetUsername()+ u'\\Downloads\\全网总损耗报表（日）.xls'
        self.sheetname=u'全网总损耗报表（日）'
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        #print self.Excel_Data
        return self.Excel_Data

if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('wjtest','000000')
    time.sleep(5)
    Func=NewtonMethodWholelossDayCalcResult(driver)
    TryReturn = Func.NewtonMethodWholelossDayCalcResult_Fun()
    time.sleep(10)
    driver.quit()