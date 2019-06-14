#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2018/12/20 16:45
当前项目名称  ：Auto
功能          ：配电线路潮流精确算法-日报表-配电线路线损表结果核对
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
import yaml
import Login
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
class DistributionlineFlowDayCalcResult:
    def __init__(self,driver):
        '''潮流精确算法配电线路线损表结果验证'''
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistributionlineFlowDayCalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_Data['DistributionlineFlowDayCalcResult']

        #点击理论线损
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #点击配电网计算结果查询
        self.TheoreticalLineLosscalresult_Xpath=self.Data['TheoreticalLineLosscalresult_Xpath']
        #点击日报表
        self.DistributionLineDayReport_Xpath=self.Data['DistributionLineDayReport_Xpath']
        #点击配电线路线损表
        self.DistributionlinedaycalcResult_Xpath=self.Data['DistributionlinedaycalcResult_Xpath']
        #切换到配电线路线损表的frame中
        self.DistributionlinedaycalcResult_frame=self.Data['DistributionlinedaycalcResult_frame']
        #点击导出
        self.Export_ID=self.Data['Export_ID']

    def DistributionlineFlowDayCalcResult_Fun(self):
        '''潮流精确算法配电线路线损表结果验证'''
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
        #切换到配电线路线损表的frame中
        self.driver.switch_to_frame(self.DistributionlinedaycalcResult_frame)
        time.sleep(5)
        #点击导出
        self.driver.find_element_by_id(self.Export_ID).click()
        time.sleep(10)
        self.data_file=u"C:\\Users\\Admin\Downloads\配电线路损耗报表（日）.xls"
        self.sheetname=u"配电线路损耗报表（日）"
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        return self.Excel_Data


if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=DistributionlineFlowDayCalcResult(driver)
    Try1.DistributionlineFlowDayCalcResult_Fun()
    time.sleep(10)
    driver.quit()


