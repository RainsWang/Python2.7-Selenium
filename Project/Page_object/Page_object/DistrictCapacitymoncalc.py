#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：王敏
IDE           ：PyCharm
时间          ：2019/1/4 13:53
当前项目名称  ：Project
功能          ：400V容量法手工计算
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import yaml
import time
import Login
from Public_Theory.GetProjectFilePath import GetProjectFilePath
import sys

reload(sys)
sys.setdefaultencoding('utf-8')


class DistrictCapacitymoncalc:

    def __init__(self, driver):
        '''配线月电量法计算'''
        # self.driver=webdriver.Chrome()
        self.driver = driver
        self.ProjectFilePath = GetProjectFilePath()
        self.Page_object_data_file = open(self.ProjectFilePath + "\Page_object\Data\DistrictCapacitymoncalc.yaml")
        self.Page_Data = yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data = self.Page_Data['DistrictCapacitymoncalc']

        # 理论线损计算id
        self.TheoreticalLineLoss_ID = self.Data['TheoreticalLineLoss_ID']
        # 理论线损手工计算
        self.TheoreticalLineLossManualcal_Xpath = self.Data['TheoreticalLineLossManualcal_Xpath']
        # 切换到配电网400V理论线损所在frame并点击
        self.TheoreticalLineloss400V_Frame = self.Data['TheoreticalLineloss400V_Frame']
        self.TheoreticalLineloss400V_ID = self.Data['TheoreticalLineloss400V_ID']

        # 直接选择计算的台区
        # 勾选wm县电量法配线月结果验证，总共四层，单位、变电站、配电线路、台区
        # 单位xpath
        self.PowerlawCalculationResultsCom_Xpath = self.Data['PowerlawCalculationResultsCom_Xpath']
        # 变电站
        self.PowerlawCalculationResultsSub_Xpath = self.Data['PowerlawCalculationResultsSub_Xpath']
        # 配电线路
        self.PowerlawCalculationResultsWiring_Xpath = self.Data['PowerlawCalculationResultsWiring_Xpath']
        # 台区
        self.PowerlawCalculationResultsDistrict_Xpath = self.Data['PowerlawCalculationResultsDistrict_Xpath']
        #选择计算方法
        self.CapacityCalaMethod_id = self.Data['CapacityCalaMethod_id']
        # 选择月计算
        self.Districtmoncalc_ID = self.Data['Districtmoncalc_ID']
        # 点击计算按钮
        self.Calculate_ID = self.Data['Calculate_ID']
        # 计算成功后弹出框信息ID以及确定按钮ID
        self.AlertMsg_ID = self.Data['AlertMsg_ID']
        self.AlertConfirm_ID = self.Data['AlertConfirm_ID']

    def DistrictCapacitymoncalc_Fun(self):
        # 台区月电量法计算#
        # 点击理论线损计算#
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(5)
        # 理论线损手工计算
        self.driver.find_element_by_xpath(self.TheoreticalLineLossManualcal_Xpath).click()
        time.sleep(5)
        # 切换到配电网400V理论线损所在frame并点击
        self.driver.switch_to_frame(self.TheoreticalLineloss400V_Frame)
        time.sleep(5)
        self.driver.find_element_by_id(self.TheoreticalLineloss400V_ID).click()
        time.sleep(5)

        # 单位xpath
        self.driver.find_element_by_xpath(self.PowerlawCalculationResultsCom_Xpath).click()
        time.sleep(2)
        # 变电站
        self.driver.find_element_by_xpath(self.PowerlawCalculationResultsSub_Xpath).click()
        time.sleep(2)
        # 配电线路
        self.driver.find_element_by_xpath(self.PowerlawCalculationResultsWiring_Xpath).click()
        time.sleep(2)
        # 台区
        self.driver.find_element_by_xpath(self.PowerlawCalculationResultsDistrict_Xpath).click()
        time.sleep(2)
        self.driver.find_element_by_id(self.CapacityCalaMethod_id).click()
        time.sleep(2)
        self.driver.find_element_by_id(self.Districtmoncalc_ID).click()
        # 点击计算按钮
        self.driver.find_element_by_id(self.Calculate_ID).click()
        self.driver.implicitly_wait(30)
        # 计算成功后弹出框信息ID以及取消按钮ID
        # self.driver.find_element_by_id(self.AlertMsg_ID).click()
        # time.sleep(3)
        alerttext = self.driver.find_element_by_id(self.AlertMsg_ID).text
        '''if u'是否要打开当前台区的计算结果？'==alerttext:
            print "计算成功，跳出页面"
        else:
            print"无法正常计算"'''
        # time.sleep(11)
        self.driver.find_element_by_id(self.AlertConfirm_ID).click()
        time.sleep(5)
        return alerttext


if __name__=="__main__":
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function(u'wm县','000000')
    Try1 = DistrictCapacitymoncalc(driver)
    Try1.DistrictCapacitymoncalc_Fun()
    time.sleep(10)
    driver.quit()