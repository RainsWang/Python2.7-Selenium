#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2018/12/18 16:11
当前项目名称  ：Auto
功能          ：配电线路潮流精确算法日计算
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import yaml
import time
import Login
from Public_Theory.GetProjectFilePath import GetProjectFilePath

class DistributionlineFlowDayCalc:
    def __init__(self,driver):
        '''配电线路潮流精确算法计算'''
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistributionlineFlowDayCalc.yaml")
        self.Page_Data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_Data['DistributionlineFlowDayCalc']

        #理论线损id
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #理论线损手工计算xpath
        self.TheoreticalLineLossManualcal_Xpath=self.Data['TheoreticalLineLossManualcal_Xpath']
        #切换到配电网6-20kV理论线损计算所在的frame并点击
        self.DistributionlineFlowDayCalc_Frame=self.Data['DistributionlineFlowDayCalc_Frame']
        self.DistributionlineFlowDayCalc_id=self.Data['DistributionlineFlowDayCalc_id']
        #选择单位
        self.DistributionlineFlowDayCalcCompany_Xpath=self.Data['DistributionlineFlowDayCalcCompany_Xpath']
        #选择变电站
        self.DistributionlineFlowDayCalcSubstation_Xpath=self.Data['DistributionlineFlowDayCalcSubstation_Xpath']
        #选择配电线路
        self.DistributionlineFlowDayCalcDistributionline_Xpath=self.Data['DistributionlineFlowDayCalcDistributionline_Xpath']
        #点击潮流精确算法
        self.Algorithm_ID=self.Data['Algorithm_ID']
        #点击计算按钮
        self.Calculate_ID=self.Data['Calculate_ID']
        #计算成功后弹出框及点击确定按钮
        self.AlertMsg_ID=self.Data['AlertMsg_ID']
        self.AlertConfirm_ID=self.Data['AlertConfirm_ID']

    def DistributionlineFlowDayCalc_Fun(self):
        '''配电线路潮流精确算法计算'''
        #点击理论线损
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #点击理论线损手工计算
        self.driver.find_element_by_xpath(self.TheoreticalLineLossManualcal_Xpath).click()
        time.sleep(2)
        #切换到配电网6-20kV理论线损计算所在的frame并点击
        self.driver.switch_to_frame(self.DistributionlineFlowDayCalc_Frame)
        time.sleep(2)
        self.driver.find_element_by_id(self.DistributionlineFlowDayCalc_id).click()
        time.sleep(2)
        #选择单位
        self.driver.find_element_by_xpath(self.DistributionlineFlowDayCalcCompany_Xpath).click()
        time.sleep(2)
        #选择变电站
        self.driver.find_element_by_xpath(self.DistributionlineFlowDayCalcSubstation_Xpath).click()
        time.sleep(2)
        #选择配电线路
        self.driver.find_element_by_xpath(self.DistributionlineFlowDayCalcDistributionline_Xpath).click()
        time.sleep(2)
        #点击潮流精确算法
        self.driver.find_element_by_id(self.Algorithm_ID).click()
        time.sleep(2)
        #点击计算
        self.driver.find_element_by_id(self.Calculate_ID).click()
        self.driver.implicitly_wait(45)
        alerttext=self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(2)
        self.driver.find_element_by_id(self.AlertConfirm_ID).click()
        time.sleep(1)
        return alerttext
if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=DistributionlineFlowDayCalc(driver)
    Try1.DistributionlineFlowDayCalc_Fun()
    time.sleep(10)
    driver.quit()

