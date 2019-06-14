#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：王敏
IDE           ：PyCharm Community Edition
时间          ：2018/10/23 10:50
当前项目名称  ：400VCalc
功能          ：月配线结果验证封装
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import time
import yaml
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
from Public_Theory.GetUsername import GetUsername
import Login
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

class DistrictEleQuamoncalcResult:
    def __init__(self,driver):
        '''400V月电量法计算结果验证'''
        #self.driver=webdriver.Chrome()
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistrictEleQuamoncalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_Data['DistrictEleQuamoncalcResult']

        #理论线损id
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #点击配电线路计算结果查询xpath
        self.TheoreticalLineLossManualcal_Xpath=self.Data['TheoreticalLineLossManualcal_Xpath']
        #点击配线月报表
        self.DistributionLineMonthlyReport_Xpath=self.Data['DistributionLineMonthlyReport_Xpath']
        #点击台区线损表
        self.LinelossmeterofDistrict_Xpath=self.Data['LinelossmeterofDistrict_Xpath']
        #切换到结果frame
        self.DistrictCalcResult_frame=self.Data['DistrictCalcResult_frame']
        #点击导出
        self.export_id =self.Data['export_id']
    def DistrictEleQuamoncalcResult_Fun(self):
        '''400V月电量法计算结果验证'''
        #点击理论线损
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #点击配电线路计算结果查询xpath
        self.driver.find_element_by_xpath(self.TheoreticalLineLossManualcal_Xpath).click()
        time.sleep(2)
        #点击配线月报表
        self.driver.find_element_by_xpath(self.DistributionLineMonthlyReport_Xpath).click()
        time.sleep(2)
        #点击台区线路线损表
        self.driver.find_element_by_xpath(self.LinelossmeterofDistrict_Xpath).click()
        time.sleep(2)
        #切换到结果frame
        self.driver.switch_to_frame(self.DistrictCalcResult_frame)
        #self.driver.implicitly_wait(30)
        time.sleep(4)
        # a=self.driver.find_element_by_id(self.export_id).text
        # print  a
        #点击导出按钮
        self.driver.find_element_by_id(self.export_id).click()
        time.sleep(5)
        self.data_path= u'C:\\Users\\'+GetUsername()+u'\\Downloads\\低压台区损耗表报表（月）.xls'
        self.sheetname= u'低压台区损耗表报表（月）'
        self.Excel=ExcelData(self.data_path,self.sheetname)
        self.ExcelDataa=self.Excel.readExcel()
        #return json.dumps(self.ExcelDataa, encoding="UTF-8", ensure_ascii=False)
        return self.ExcelDataa


if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function(u'wm县','000000')
    Try1=DistrictEleQuamoncalcResult(driver)
    Try1.DistrictEleQuamoncalcResult_Fun()
    time.sleep(10)
    driver.quit()

