# -*- coding: utf-8 -*-
#!/usr/bin/env python
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：吴鹏飞
IDE           ：PyCharm Community Edition
时间          ：2018/10/22 15:13
当前项目名称  ：eagle2
功能          ：
-------------------------------------漂亮的分割线---------------------------------------------'''
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
class TranotherlosshourCalcResult:
    def __init__(self,driver):
        #驱动
        self.driver=driver
        #获取页面元素
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_file=open(self.ProjectFilePath+"\\Page_object\\Data\\TranotherlosshourCalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_file)
        self.Data=self.Page_Data['TranotherlosshourCalcResult']
        #理论线损id
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #输电网计算结果查询
        self.TranotherlossManualcal_Xpath=self.Data['TranotherlossManualcal_Xpath']
        #整点报表
        self.TranotherlosshourReport_Xpath=self.Data['TranotherlosshourReport_Xpath']
        #其它损耗表
        self.TranotherlosshourCalcResult_Xpath=self.Data['TranotherlosshourCalcResult_Xpath']
        #查询报表frame
        self.TranotherlosshourCalcResult_frame=self.Data['TranotherlosshourCalcResult_frame']
        #打开日期控件
        self.Day_box_xpath=self.Data['Day_box_xpath']
        #选择20140801_xpath
        self.DayResult_xpath=self.Data['DayResult_xpath']
        #输入零点
        self.ZeroHour_ID=self.Data['ZeroHour_ID']
        #点击日期控件中确定
        self.DayQueding_id=self.Data['DayQueding_id']
        #点击查询
        self.Chaxun_id=self.Data['Chaxun_id']
        #点击导出
        self.Daochu_id=self.Data['Daochu_id']
    def TranotherlosshourCalcResult_Fun(self):
        #进入理论线损模块
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #进入输电网计算结果查询
        self.driver.find_element_by_xpath(self.TranotherlossManualcal_Xpath).click()
        time.sleep(2)
        #进入整点报表
        self.driver.find_element_by_xpath(self.TranotherlosshourReport_Xpath).click()
        time.sleep(2)
        #进入其他损耗表
        self.driver.find_element_by_xpath(self.TranotherlosshourCalcResult_Xpath).click()
        time.sleep(5)
        #进入查询frame框
        self.driver.switch_to_frame(self.TranotherlosshourCalcResult_frame)
        time.sleep(3)
        #打开日期控件
        self.driver.find_element_by_id(self.Day_box_xpath).click()
        time.sleep(2)
        #选择20140801
        self.driver.find_element_by_xpath(self.DayResult_xpath).click()
        time.sleep(2)
        #选择零点
        self.driver.find_element_by_id(self.ZeroHour_ID).clear()
        time.sleep(2)
        self.driver.find_element_by_id(self.ZeroHour_ID).send_keys('0')
        time.sleep(2)
        #点击日期控件确定
        self.driver.find_element_by_id(self.DayQueding_id).click()
        time.sleep(2)
        #点击查询
        self.driver.find_element_by_id(self.Chaxun_id).click()
        time.sleep(3)
        #点击导出
        self.driver.find_element_by_id(self.Daochu_id).click()
        time.sleep(10)
        self.data_file=u'C:\\Users\\'+GetUsername()+u'\\Downloads\\输电网其他损耗报表（小时）.xls'
        self.sheetname=u'输电网其他损耗报表（小时）'
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        #print self.Excel_Data
        return self.Excel_Data
if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=TranotherlosshourCalcResult(driver)
    Try1.TranotherlosshourCalcResult_Fun()
    time.sleep(10)
    driver.quit()