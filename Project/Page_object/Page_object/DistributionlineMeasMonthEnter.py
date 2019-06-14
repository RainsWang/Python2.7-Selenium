#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/8 18:06
当前项目名称  ：620Calc
功能          ：620kv运行数据录入
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import yaml
import time
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from selenium.webdriver.common.keys import Keys
import Login
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class DistributionlineMeasMonthEnter:
    '''配线计算所用到的运行数据录入'''
    def __init__(self,driver):
        #self.driver=webdriver.Chrome()
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()

        #获取页面元素
        self.Page_object_data_file=open(self.ProjectFilePath+"\\Page_object\\Data\\DistributionlineMeasMonthEnter.yaml")
        self.Page_data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_data['DistributionlineMeasMonthEnter']
        #数据中心id
        self.DataCenter_id = self.Data['DataCenter_id']
        #运营数据管理路径
        self.RunDataManage_Xpath = self.Data['RunDataManage_Xpath']
        #配线线路运行数据录入
        self.LineDataEntry_Xpath = self.Data['LineDataEntry_Xpath']
        #配线理论线损运行数据录入
        self.TheoreticalLineLoss_Xpath = self.Data['TheoreticalLineLoss_Xpath']

        #理论线损运行数据录入页面frame
        self.TheoreticalLineLoss_Frame = self.Data['TheoreticalLineLoss_Frame']

        #查询条件之量测分组选择，先点击下拉框，再选择对应的数据
        self.MeasurementPacketDropdownbox_ID = self.Data['MeasurementPacketDropdownbox_ID']
        self.TheoreticalMounthChoose_Xpath = self.Data['TheoreticalMounthChoose_Xpath']

        #输入配线参数，包含正向有功电量、功率因数、负荷形状系数
        #正向有功电量
        self.ActivePowerActivation_Xpath = self.Data['ActivePowerActivation_Xpath']  #激活输入框
        self.ActivePower_name = self.Data['ActivePower_name']
        #功率因数
        self.PowerFactorActivation_Xpath = self.Data['PowerFactorActivation_Xpath']
        self.PowerFactor_name = self.Data['PowerFactor_name']
        #负荷形状系数
        self.LoadShapeFactorActivation_Xpath=self.Data['LoadShapeFactorActivation_Xpath']#激活输入框
        self.LoadShapeFactor_name=self.Data['LoadShapeFactor_name']
        #输入配变数据
        self.DistributionTransform_ID = self.Data['DistributionTransform_ID']#切换到配变tab页面
        self.DistributionTransformActivationBefor=self.Data['DistributionTransformActivationBefor']
        self.DistributionTransformActivationAfter=self.Data['DistributionTransformActivationAfter']
        self.DistributionTransformData_name=self.Data['DistributionTransformData_name']

        #数据输入后保存，以及弹出框内容校验
        self.SaveButton_ID = self.Data['SaveButton_ID']
        self.AlertMsg_ID = self.Data['AlertMsg_ID']
        self.AlertButton_ID = self.Data['AlertButton_ID']
    def DistributionlineMeasMonthEnter_Fun(self,ActivePower,PowerFactor,LoadShapeFactor,DistributionTransformData1,DistributionTransformData2,DistributionTransformData3,DistributionTransformData4):
        '''配线月运行数据录入'''
        #ActivePower 正向有功电量
        #PowerFactor 功率因数
        #LoadShapeFactor 负荷形状系数
        #DistributionTransformData1第一个配变数据
        #DistributionTransformData2第二个配变数据
        #DistributionTransformData3第三个配变数据
        #DistributionTransformData4第四个配变数据

        #打开左侧的树结构，进入到理论线损数据录入界面
        #数据中心
        '''配线计算所用到的运行数据录入'''
        self.driver.find_element_by_id(self.DataCenter_id).click()
        time.sleep(2)
        #运行数据管理
        self.driver.find_element_by_xpath(self.RunDataManage_Xpath).click()
        time.sleep(3)
        #配线理论运行数据录入
        self.driver.find_element_by_xpath(self.LineDataEntry_Xpath).click()
        time.sleep(3)
        #配线理论运行数据录入
        self.driver.find_element_by_xpath(self.TheoreticalLineLoss_Xpath).click()
        time.sleep(3)

        #切换到理论线损录入界面，筛查出信息
        ##理论线损运行数据录入页面frame
        self.driver.switch_to_frame(self.TheoreticalLineLoss_Frame)
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
        self.driver.find_element_by_id(self.DistributionTransform_ID).click()
        time.sleep(5)
        lista=[DistributionTransformData1,DistributionTransformData2,DistributionTransformData3,DistributionTransformData4]
        for i in range(2,6):
            self.driver.find_element_by_xpath(self.DistributionTransformActivationBefor+str(i)+self.DistributionTransformActivationAfter).click()
            time.sleep(2)
            self.driver.find_element_by_name(self.DistributionTransformData_name).clear()
            self.driver.find_element_by_name(self.DistributionTransformData_name).send_keys(lista[i-2])
            time.sleep(3)
        self.driver.find_element_by_name(self.DistributionTransformData_name).send_keys(Keys.TAB)
        #数据输入完成后
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(3)
        AlertMsg = self.driver.find_element_by_id(self.AlertMsg_ID).text
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        time.sleep(5)
        return AlertMsg


if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('Autotest','000000')
    Try1=DistributionlineMeasMonthEnter(driver)
    TryReturn = Try1.DistributionlineMeasMonthEnter_Fun('6666','0.23','1.23','221','222','223','225')
    print TryReturn
    time.sleep(5)
    driver.quit()