#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2018/12/21 14:43
当前项目名称  ：Auto
功能          ：配电线路潮流精确算法-日报表-配电线路线损表中导线损耗结果核对
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
import Login
import yaml
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
class DistAclinesegeFlowDayCalcResult:
    def __init__(self,driver):
        '''潮流精确算法配线导线损耗结果验证'''
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data=open(self.ProjectFilePath+"\Page_object\Data\DistAclinesegeFlowDayCalcResult.yaml")
        self.Page_data=yaml.load(self.Page_object_data)
        self.Page_object_data.close()
        self.Data=self.Page_data['DistAclinesegeFlowDayCalcResult']

        #点击理论线损
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #点击配电网计算结果查询
        self.TheoreticalLineLosscalresult_Xpath=self.Data['TheoreticalLineLosscalresult_Xpath']
        #点击日报表
        self.DistributionLineDayReport_Xpath=self.Data['DistributionLineDayReport_Xpath']
        #点击配电线路线损表
        self.DistributionlinedaycalcResult_Xpath=self.Data['DistributionlinedaycalcResult_Xpath']
        #切换到配电线路线损表frame中
        self.DistributionlinedaycalcResult_frame=self.Data['DistributionlinedaycalcResult_frame']
        #点击损耗值
        self.DistAclineLoss_Xpath=self.Data['DistAclineLoss_Xpath']
        #切换到配电线路导线损耗表（24点）frame中
        self.DistAclineLossdaycalcResult_frame=self.Data['DistAclineLossdaycalcResult_frame']
        #点击左下角的导出
        self.Export_ID=self.Data['Export_ID']

    def DistAclinesegeFlowDayCalcResult_Fun(self):
        '''潮流精确算法配线导线损耗结果验证'''
        #点击理论线损
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #点击配电网计算结果查询
        self.driver.find_element_by_xpath(self.TheoreticalLineLosscalresult_Xpath).click()
        time.sleep(2)
        #点击日报表
        self.driver.find_element_by_xpath(self.DistributionLineDayReport_Xpath).click()
        time.sleep(2)
        #点击配电线路线损表
        self.driver.find_element_by_xpath(self.DistributionlinedaycalcResult_Xpath).click()
        time.sleep(2)
        #切换到配电线路线损表frame中
        self.driver.switch_to_frame(self.DistributionlinedaycalcResult_frame)
        self.driver.implicitly_wait(15)
        #点击损耗值
        self.driver.find_element_by_xpath(self.DistAclineLoss_Xpath).click()
        time.sleep(2)
        self.driver.switch_to.parent_frame()   ##跳出当前frame
        self.driver.implicitly_wait(10)
        #切换到配电线路导线损耗表（24点）frame中
        self.driver.switch_to_frame(self.DistAclineLossdaycalcResult_frame)
        self.driver.implicitly_wait(15)
        #点击左下角的导出
        self.driver.find_element_by_id(self.Export_ID).click()
        time.sleep(10)
        self.data_file=u"C:\\Users\\Admin\Downloads\配电线路导线损耗报表（24点）.xls"
        self.sheetname=u"配电线路导线损耗报表（24点）"
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_data=self.Excel.readExcel()
        return self.Excel_data

if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=DistAclinesegeFlowDayCalcResult(driver)
    Try1.DistAclinesegeFlowDayCalcResult_Fun()
    time.sleep(10)
    driver.quit()





