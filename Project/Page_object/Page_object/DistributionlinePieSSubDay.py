#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2019/1/2 17:23
当前项目名称  ：Auto
功能          ：配电线路线损率分段-变电站追溯结果核对
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
import yaml
import Login
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class DistributionlinePieSSubDay:
    def __init__(self,driver):
        '''配电线路线损率分段变电站追溯统计结果核对'''
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistributionlinePieSSubDay.yaml")
        self.Page_data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_data['DistributionlinePieSSubDay']

        #点击线损分析
        self.LineLossAnalysis_ID=self.Data['LineLossAnalysis_ID']
        #点击配电网计算结果分析
        self.TheoreticalLineLosscalresultAnalysis_Xpath=self.Data['TheoreticalLineLosscalresultAnalysis_Xpath']
        #点击日分析
        self.DistributionLineDayAnalysis_Xpath=self.Data['DistributionLineDayAnalysis_Xpath']
        #点击配电线路线损率分段统计
        self.DistributionlineLossRateSegStatistics_Xpath=self.Data['DistributionlineLossRateSegStatistics_Xpath']
        #切换到配电线路线损率分段统计的frame中
        self.DistributionlineLossRateSegStatistics_frame=self.Data['DistributionlineLossRateSegStatistics_frame']
        #切换到配电线路线损率分段统计查询结果的frame中
        self.DistributionlineLossRateSegStatisticsResult_frame=self.Data['DistributionlineLossRateSegStatisticsResult_frame']
        #点击单位追溯
        self.CompanyRetrospect_css=self.Data['CompanyRetrospect_css']
        #切换到配电网线损率分段统计追溯报表（分单位日）frame中
        self.CompanyRetrospect_frame=self.Data['CompanyRetrospect_frame']
        #切换到配电网线损率分段统计追溯报表（分单位日）结果frame中
        self.CompanyRetrospectResult_frame=self.Data['CompanyRetrospectResult_frame']
        #点击变电站名称追溯
        self.SubstationRetrospect_link=self.Data['SubstationRetrospect_link']
        #切换到配电网线损率分段统计追溯报表（分变电站日）frame中
        self.SubstationRetrospectResult_frame=self.Data['SubstationRetrospectResult_frame']
        #点击导出
        self.Export_ID=self.Data['Export_ID']

    def DistributionlinePieSSubDay_Fun(self):
        '''配电线路线损率分段变电站追溯统计结果核对'''
        #点击线损分析
        self.driver.find_element_by_id(self.LineLossAnalysis_ID).click()
        time.sleep(2)
        #点击配电网计算结果分析
        self.driver.find_element_by_xpath(self.TheoreticalLineLosscalresultAnalysis_Xpath).click()
        time.sleep(2)
        #点击日分析
        self.driver.find_element_by_xpath(self.DistributionLineDayAnalysis_Xpath).click()
        time.sleep(2)
        #点击配电线路线损率分段统计
        self.driver.find_element_by_xpath(self.DistributionlineLossRateSegStatistics_Xpath).click()
        time.sleep(2)
        #切换到配电线路线损率分段统计的frame中
        self.driver.switch_to.frame(self.DistributionlineLossRateSegStatistics_frame)
        time.sleep(2)
        self.driver.implicitly_wait(10)
        #切换到配电线路线损率分段统计查询结果的frame中
        self.driver.switch_to.frame(self.DistributionlineLossRateSegStatisticsResult_frame)
        self.driver.implicitly_wait(10)
        #点击单位追溯
        self.driver.find_element_by_css_selector(self.CompanyRetrospect_css).click()
        time.sleep(3)
        self.driver.switch_to.parent_frame()   ##跳出结果框的frame
        time.sleep(5)
        self.driver.switch_to.default_content()    ##跳出最开始的页面frame
        time.sleep(5)
        #切换到配电网线损率分段统计追溯报表（分单位日）frame中
        self.driver.switch_to.frame(self.CompanyRetrospect_frame)
        time.sleep(2)
        self.driver.implicitly_wait(10)
        #切换到配电网线损率分段统计追溯报表（分单位日）结果frame中
        self.driver.switch_to.frame(self.CompanyRetrospectResult_frame)
        self.driver.implicitly_wait(10)
        #点击变电站名称追溯
        self.driver.find_element_by_link_text(self.SubstationRetrospect_link).click()
        time.sleep(3)
        self.driver.switch_to.parent_frame()   ##跳出结果框的frame
        time.sleep(5)
        self.driver.switch_to.parent_frame()    ##跳出分单位的页面frame
        time.sleep(5)
        #切换到配电网线损率分段统计追溯报表（分变电站日）frame中
        self.driver.switch_to.frame(self.SubstationRetrospectResult_frame)
        time.sleep(5)
        #点击导出
        self.driver.find_element_by_id(self.Export_ID).click()
        time.sleep(5)
        self.data_file=u"C:\\Users\\Admin\Downloads\配电线路线损率分段统计表样三(日).xlsx"
        self.sheetname="Sheet1"
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        #print self.Excel_Data
        return self.Excel_Data

if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=DistributionlinePieSSubDay(driver)
    Try1.DistributionlinePieSSubDay_Fun()
    time.sleep(10)
    driver.quit()