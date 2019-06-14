#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018-8-27 17:59
当前项目名称  ：Auto
功能          ：电量法日计算结果验证
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

class DistributionlineEleQuadaycalcResult:
    def __init__(self,driver):
        '''620kv日电量法计算结果验证'''
        #self.driver=webdriver.Chrome()
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistributionlineEleQuadaycalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_Data['DistributionlineEleQuadaycalcResult']

        #理论线损id
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #点击配电线路计算结果查询xpath
        self.TheoreticalLineLossManualcal_Xpath=self.Data['TheoreticalLineLossManualcal_Xpath']
        #点击配线日报表
        self.DistributionLineDayReport_Xpath=self.Data['DistributionLineDayReport_Xpath']
        #点击配电线路线损表
        self.Linelossmeterofdistributionline_Xpath=self.Data['Linelossmeterofdistributionline_Xpath']
        #切换到结果frame
        self.DistributionLineCalcResult_frame=self.Data['DistributionLineCalcResult_frame']
        #点击导出
        self.Daochu_id=self.Data['Daochu_id']

    def DistributionlineEleQuadaycalcResult_Fun(self):
        '''620kv日电量法计算结果验证'''
        #点击理论线损
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #点击配电线路计算结果查询xpath
        self.driver.find_element_by_xpath(self.TheoreticalLineLossManualcal_Xpath).click()
        time.sleep(2)
        #点击配线日报表
        self.driver.find_element_by_xpath(self.DistributionLineDayReport_Xpath).click()
        time.sleep(2)
        #点击配电线路线损表
        self.driver.find_element_by_xpath(self.Linelossmeterofdistributionline_Xpath).click()
        time.sleep(2)
        #切换到结果frame
        self.driver.switch_to_frame(self.DistributionLineCalcResult_frame)
        self.driver.implicitly_wait(30)
        #点击导出按钮
        self.driver.find_element_by_id(self.Daochu_id).click()
        time.sleep(10)
        self.data_path= u"C:\\Users\\"+GetUsername()+ u"\\Downloads\\配电线路损耗报表（日）.xls"
        self.sheetname= u'配电线路损耗报表（日）'
        self.Excel=ExcelData(self.data_path,self.sheetname)
        self.ExcelDataa=self.Excel.readExcel()
        #return json.dumps(self.ExcelDataa, encoding="UTF-8", ensure_ascii=False)
        return self.ExcelDataa
if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lyltest','000000')
    Try1=DistributionlineEleQuadaycalcResult(driver)
    Try1.DistributionlineEleQuadaycalcResult_Fun()
    time.sleep(10)
    driver.quit()
