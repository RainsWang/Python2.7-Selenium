#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2019/1/2 9:44
当前项目名称  ：Auto
功能          ：配电线路综合分析日报表-配线追溯核对
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
import yaml
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
import Login
class DistributionlineComAAOLLCDay:
    def __init__(self,driver):
        '''配线追溯页面结果验证'''
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistributionlineComAAOLLCDay.yaml")
        self.Page_data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_data['DistributionlineComAAOLLCDay']

        #点击线损分析
        self.LineLossAnalysis_ID=self.Data['LineLossAnalysis_ID']
        #点击配电网计算结果分析
        self.TheoreticalLineLosscalresultAnalysis_Xpath=self.Data['TheoreticalLineLosscalresultAnalysis_Xpath']
        #点击日分析
        self.DistributionLineDayAnalysis_Xpath=self.Data['DistributionLineDayAnalysis_Xpath']
        #点击配电线路综合分析
        self.DistributionlinedaycalcAnalysis_Xpath=self.Data['DistributionlinedaycalcAnalysis_Xpath']
        #切换到配电线路综合分析的frame中
        self.DistributionlinedaycalcAnalysis_frame=self.Data['DistributionlinedaycalcAnalysis_frame']
        #点击查询条件中日期的下拉框
        self.DateDropBox_CSS=self.Data['DateDropBox_CSS']
        self.DateDropBox_CSS1=self.Data['DateDropBox_CSS1']
        #切换到日期选择框中的frame
        self.Date_frame=self.Data['Date_frame']
        #点击日期下拉框年的下拉框
        self.DateYearDropBox_ID=self.Data['DateYearDropBox_ID']
        #选择日期下拉框年中的2015
        self.YearSelect2015_Xpath=self.Data['YearSelect2015_Xpath']
        #点击日期下拉框代表日的下拉框
        self.DateRepresentativeDay_ID=self.Data['DateRepresentativeDay_ID']
        #选择日期下拉框代表日中的2015-08-01
        self.DateRepresentativeDay20150801_Xpath=self.Data['DateRepresentativeDay20150801_Xpath']
        #点击日期下拉框中的确定按钮
        self.DateDropBoxSure_ID=self.Data['DateDropBoxSure_ID']
        #点击页面中的查询按钮
        self.Query_Class=self.Data['Query_Class']
        #点击配电线路追溯
        self.Distributioin_Xpath=self.Data['Distributioin_Xpath']
        #切换到配电线路线损组成分析（日）frame中
        self.DistributioinRetrospect_frame=self.Data['DistributioinRetrospect_frame']
        #点击配线追溯页面的导出
        self.Export_Class=self.Data['Export_Class']

    def DistributionlineComAAOLLCDay_Fun(self):
        '''配线追溯页面结果验证'''
        #点击线损分析
        self.driver.find_element_by_id(self.LineLossAnalysis_ID).click()
        time.sleep(2)
        #点击配电网计算结果分析
        self.driver.find_element_by_xpath(self.TheoreticalLineLosscalresultAnalysis_Xpath).click()
        time.sleep(2)
        #点击日分析
        self.driver.find_element_by_xpath(self.DistributionLineDayAnalysis_Xpath).click()
        time.sleep(2)
        #点击配电线路综合分析
        self.driver.find_element_by_xpath(self.DistributionlinedaycalcAnalysis_Xpath).click()
        time.sleep(2)
        #切换到配电线路综合分析的frame中
        self.driver.switch_to.frame(self.DistributionlinedaycalcAnalysis_frame)
        time.sleep(2)
        self.driver.implicitly_wait(15)
        #点击查询条件中日期的下拉框
        self.driver.find_element_by_css_selector(self.DateDropBox_CSS).click()
        time.sleep(2)
        self.driver.find_element_by_css_selector(self.DateDropBox_CSS1).click()
        time.sleep(2)
        #切换到日期选择框中的frame
        cf=self.driver.find_element_by_css_selector(self.Date_frame)
        self.driver.switch_to.frame(cf)
        time.sleep(5)
        #点击日期下拉框年的下拉框
        self.driver.find_element_by_id(self.DateYearDropBox_ID).click()
        time.sleep(2)
        #选择日期下拉框年中的2015
        self.driver.find_element_by_xpath(self.YearSelect2015_Xpath).click()
        time.sleep(2)
        #点击日期下拉框代表日的下拉框
        self.driver.find_element_by_id(self.DateRepresentativeDay_ID).click()
        time.sleep(2)
        #选择日期下拉框代表日中的2015-08-01
        self.driver.find_element_by_xpath(self.DateRepresentativeDay20150801_Xpath).click()
        time.sleep(2)
        #点击日期下拉框中的确定按钮
        self.driver.find_element_by_id(self.DateDropBoxSure_ID).click()
        time.sleep(2)
        self.driver.switch_to.parent_frame()   ##跳出日期下拉框的frame
        self.driver.implicitly_wait(10)
        #点击页面中的查询按钮
        self.driver.find_element_by_class_name(self.Query_Class).click()
        time.sleep(7)
        #点击配电线路追溯
        self.driver.find_element_by_xpath(self.Distributioin_Xpath).click()
        time.sleep(3)
        #跳出配电线路线损分析的frame
        self.driver.switch_to.default_content()
        self.driver.implicitly_wait(10)
        #切换到配电线路线损组成分析（日）frame中
        self.driver.switch_to.frame(self.DistributioinRetrospect_frame)
        self.driver.implicitly_wait(7)
        #点击配线追溯页面的导出
        self.driver.find_element_by_class_name(self.Export_Class).click()
        time.sleep(5)
        self.data_file=u"C:\\Users\\Administrator\Downloads\配电线路线损组成分析(日).xls"
        self.sheetname=u"配电线路线损组成分析(日)"
        self.sheetname1=u"图表数据"
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel1=ExcelData(self.data_file,self.sheetname1)
        self.Excel_Data=self.Excel.readExcel()
        self.Excel_Data1=self.Excel.readExcel()
        return self.Excel_Data
        return self.Excel_Data1


if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=DistributionlineComAAOLLCDay(driver)
    Try1.DistributionlineComAAOLLCDay_Fun()
    time.sleep(10)
    driver.quit()

