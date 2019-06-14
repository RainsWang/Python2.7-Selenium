#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''---------------------------------------------------------------------------------
脚本功能说明  ：
作者          ：王敏
IDE           ：PyCharm Community Edition
时间          ：2018/10/18 10:00
当前项目名称  ：400Vcala
功能          ：400V运行数据录入
-------------------------------------------------------------------------------------'''
import sys
from selenium import webdriver
import yaml
import time
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from selenium.webdriver.common.keys import Keys
import Login

reload(sys)
sys.setdefaultencoding('utf-8')

class DistrictmoncalcPcsEnter:
    '''台区计算所用到的运行数据录入'''
    def __init__(self, driver):
        #self.driver=webdriver.Chrome()
        self.driver = driver
        self.ProjectFilePath = GetProjectFilePath()

        #获取页面元素
        self.Page_object_data_file = open(self.ProjectFilePath+"\\Page_object\\Data\\DistrictmoncalcPcsEnter.yaml")

        self.Page_data = yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_data['DistrictmoncalcPcsEnter']
        #数据中心id
        self.DataCenter_id = self.Data['DataCenter_id']
        #运营数据管理路径
        self.RunDataManage_Xpath = self.Data['RunDataManage_Xpath']
        #台区线路运行数据录入
        self.DistricDataEntry_Xpath = self.Data['DistricDataEntry_Xpath']
        #台区理论线损运行数据录入
        self.TheoreticalDistricLoss_Xpath = self.Data['TheoreticalDistricLoss_Xpath']


        #理论线损运行数据录入页面frame
        self.TheoreticalDistricLoss_Frame = self.Data['TheoreticalDistricLoss_Frame']

        #查询条件之量测分组选择，先点击下拉框，再选择对应的数据
        self.MeasurementPacketDropdownbox_ID = self.Data['MeasurementPacketDropdownbox_ID']
        self.TheoreticalMounthChoose_Xpath = self.Data['TheoreticalMounthChoose_Xpath']

        #输入台区参数，包含正向有功电量、功率因数、负荷形状系数
        #正向有功电量
        self.ActivePowerActivation_Xpath = self.Data['ActivePowerActivation_Xpath']  #激活输入框
        self.ActivePower_name = self.Data['ActivePower_name']
        #功率因数
        self.PowerFactorActivation_Xpath = self.Data['PowerFactorActivation_Xpath']
        self.PowerFactor_name = self.Data['PowerFactor_name']
        #负荷形状系数
        self.LoadShapeFactorActivation_Xpath = self.Data['LoadShapeFactorActivation_Xpath']#激活输入框
        self.LoadShapeFactor_name = self.Data['LoadShapeFactor_name']
        #输入台区用户数据
        self.DistrictCustomer_ID = self.Data['DistrictCustomer_ID']#切换到台区用户tab页面
        self.DistrictCustomerBefor = self.Data['DistrictCustomerBefor']
        self.DistrictCustomerAfter = self.Data['DistrictCustomerAfter']
        self.DistrictCustomer_name = self.Data['DistrictCustomer_name']

        #数据输入后保存，以及弹出框内容校验
        self.SaveButton_ID = self.Data['SaveButton_ID']
        self.AlertMsg_ID = self.Data['AlertMsg_ID']
        self.AlertButton_ID = self.Data['AlertButton_ID']
    def DistributionlineMeasMonthEnter_Fun(self,ActivePower,PowerFactor,LoadShapeFactor,DistributionTransformData1,DistributionTransformData2,DistributionTransformData3,DistributionTransformData4):
        '''台区月运行数据录入'''
        #ActivePower 正向有功电量
        #PowerFactor 功率因数
        #LoadShapeFactor 负荷形状系数


        #打开左侧的结构层，进入到理论线损数据录入界面
        #数据中心
        '''台区计算所用到的运行数据录入'''
        self.driver.find_element_by_id(self.DataCenter_id).click()
        time.sleep(2)
        #运行数据管理
        self.driver.find_element_by_xpath(self.RunDataManage_Xpath).click()
        time.sleep(3)
        #台区理论运行数据录入
        self.driver.find_element_by_xpath(self.DistricDataEntry_Xpath).click()
        time.sleep(3)
        #理论运行数据录入
        self.driver.find_element_by_xpath(self.TheoreticalDistricLoss_Xpath).click()
        time.sleep(8)


        #切换到理论线损录入界面，筛查出信息
        ##理论线损运行数据录入页面frame
        self.driver.switch_to_frame(self.TheoreticalDistricLoss_Frame)
        #查询条件之量测分组选择，先点击下拉框，再选择对应的数据
        self.driver.find_element_by_id(self.MeasurementPacketDropdownbox_ID).click()
        time.sleep(2)
        #月量测类型分组
        self.driver.find_element_by_xpath(self.TheoreticalMounthChoose_Xpath).click()
        time.sleep(2)

        #输入相关的参数信息
        #输入正向有功电量
        self.driver.find_element_by_xpath(self.ActivePowerActivation_Xpath).click()
        time.sleep(2)
        self.driver.find_element_by_name(self.ActivePower_name).clear()
        self.driver.find_element_by_name(self.ActivePower_name).send_keys(ActivePower)
        time.sleep(2)
        #输入功率因数
        self.driver.find_element_by_xpath(self.PowerFactorActivation_Xpath).click()
        time.sleep(2)
        self.driver.find_element_by_name(self.PowerFactor_name).clear()
        self.driver.find_element_by_name(self.PowerFactor_name).send_keys(PowerFactor)
        time.sleep(2)
        #输入负荷形状系数
        self.driver.find_element_by_xpath(self.LoadShapeFactorActivation_Xpath).click()
        time.sleep(2)
        self.driver.find_element_by_name(self.LoadShapeFactor_name).clear()
        self.driver.find_element_by_name(self.LoadShapeFactor_name).send_keys(LoadShapeFactor)
        self.driver.find_element_by_name(self.LoadShapeFactor_name).send_keys(Keys.TAB)
        time.sleep(3)

        #数据输入完成后
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(2)
        self.AlertMsg = self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(2)
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        time.sleep(3)
        #return AlertMsg

        #切换到配变tab页面
        self.driver.find_element_by_id(self.DistrictCustomer_ID).click()
        time.sleep(5)
        lista = [DistributionTransformData1, DistributionTransformData2, DistributionTransformData3, DistributionTransformData4]
        for i in range(2,5):
            self.driver.find_element_by_xpath(self.DistrictCustomerBefor+str(i)+self.DistrictCustomerAfter).click()
            time.sleep(2)
            self.driver.find_element_by_name(self.DistrictCustomer_name).clear()
            self.driver.find_element_by_name(self.DistrictCustomer_name).send_keys(lista[i-2])
            time.sleep(3)
        self.driver.find_element_by_name(self.DistrictCustomer_name).send_keys(Keys.TAB)
        #数据输入完成后
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(3)
        AlertMsg = self.driver.find_element_by_id(self.AlertMsg_ID).text
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        time.sleep(5)



if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function(u'wm县','000000')
    Try1 = DistrictmoncalcPcsEnter(driver)
    TryReturn = Try1.DistrictmoncalcPcsEnter_Fun(('100', '0.2', '1'), ('40', '10', '30', '20'))
    driver.quit()
