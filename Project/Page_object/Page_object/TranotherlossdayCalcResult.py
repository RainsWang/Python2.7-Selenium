# -*- coding: utf-8 -*-
#!/usr/bin/env python
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：吴鹏飞
IDE           ：PyCharm Community Edition
时间          ：2018/10/23 15:58
当前项目名称  ：eagle2
功能          ：
-------------------------------------漂亮的分割线---------------------------------------------'''
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
import yaml
import time
import Login
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class TranotherlossdayCalcResult:
    def __init__(self,driver):
        #驱动
        self.driver=driver
        #获取页面元素
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_file=open(self.ProjectFilePath+"\\Page_object\\Data\\TranotherlossdayCalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_file)
        self.Data=self.Page_Data['TranotherlossdayCalcResult']
        #理论线损id
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #输电网计算结果查询
        self.TranotherlossManualcal_Xpath=self.Data['TranotherlossManualcal_Xpath']
        #日报表
        self.TranotherlossdayReport_Xpath=self.Data['TranotherlossdayReport_Xpath']
        #其它损耗表
        self.TranotherlossdayCalcResult_Xpath=self.Data['TranotherlossdayCalcResult_Xpath']
        #查询报表frame
        self.TranotherlossdayCalcResult_frame=self.Data['TranotherlossdayCalcResult_frame']
        #点击导出
        self.Daochu_id=self.Data['Daochu_id']
    def TranotherlossdayCalcResult_Fun(self):
        #进入理论线损模块
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #进入输电网计算结果查询
        self.driver.find_element_by_xpath(self.TranotherlossManualcal_Xpath).click()
        time.sleep(2)
        #进入日报表
        self.driver.find_element_by_xpath(self.TranotherlossdayReport_Xpath).click()
        time.sleep(2)
        #进入其他损耗表
        self.driver.find_element_by_xpath(self.TranotherlossdayCalcResult_Xpath).click()
        time.sleep(5)
        #进入查询frame框
        self.driver.switch_to_frame(self.TranotherlossdayCalcResult_frame)
        time.sleep(3)
        #点击导出
        self.driver.find_element_by_id(self.Daochu_id).click()
        time.sleep(10)
        self.data_file=u'C:\\Users\\Admin\Downloads\输电网其他损耗报表（日）.xls'
        self.sheetname=u'输电网其他损耗报表（日）'
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        return self.Excel_Data
if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=TranotherlossdayCalcResult(driver)
    Try1.TranotherlossdayCalcResult_Fun()
    time.sleep(10)
    driver.quit()