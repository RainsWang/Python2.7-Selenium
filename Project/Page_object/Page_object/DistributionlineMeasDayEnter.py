#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018-8-22 18:22
当前项目名称  ：Auto
功能          ：配线日运行数据录入()
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import yaml
import time
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from selenium.webdriver.common.keys import Keys
import Login
import json
import SendKeys
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

class DistributionlineMeasDayEnter:
    '''配线计算所用到的运行数据录入'''
    def __init__(self,driver):
        #self.driver=webdriver.Chrome()
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()

        #获取页面元素
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistributionlineMeasDayEnter.yaml")
        self.Page_data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_data['DistributionlineMeasDayEnter']
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
        #点击量测类型分组
        self.CustomGroupBox_ID=self.Data['CustomGroupBox_ID']
        #切换到24点电流量测类型分组
        self.EleCut24_Xpath=self.Data['EleCut24_Xpath']
        #24点电压激活框
        self.VolBefor=self.Data['VolBefor']
        self.VolAfter=self.Data['VolAfter']
        #24点电压数据录入框
        self.VolDataEnter=self.Data['VolDataEnter']
        #24点电流激活框
        self.EleCutBefor=self.Data['EleCutBefor']
        self.EleCutAfter=self.Data['EleCutAfter']
        #24点电流数据录入框
        self.EleCutDataEnter=self.Data['EleCutDataEnter']
        #切换日量测类型分组
        self.DataCustomGroup_Xpath=self.Data['DataCustomGroup_Xpath']
        #选择日期下拉框
        self.DataDropDownBox_ID = self.Data['DataDropDownBox_ID']
         #切换到日期选择框frame
        self.DateSelectionBox_Frame = self.Data['DateSelectionBox_Frame']
        #选择2014-8-1日
        self.DateSelect_Xpath=self.Data['DateSelect_Xpath']
        #点击日期框中的确定按钮
        self.DateAlerButton_ID=self.Data['DateAlerButton_ID']
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
        #点击“理论线损小时数据”量测类型分组
        self.DistributionTransformHourData_Xpath=self.Data['DistributionTransformHourData_Xpath']
        #有功功率激活框
        self.DistributionTransformActivePowerBefor_Xpath=self.Data['DistributionTransformActivePowerBefor_Xpath']
        self.DistributionTransformActivePowerAfter_Xpath=self.Data['DistributionTransformActivePowerAfter_Xpath']
        #有功功率输入框
        self.DistributionTransformActivePower_name=self.Data['DistributionTransformActivePower_name']
        #无功功率激活框
        self.DistributionTransformReactivePowerBefor_Xpath=self.Data['DistributionTransformReactivePowerBefor_Xpath']
        self.DistributionTransformReactivePowerAfter_Xpath=self.Data['DistributionTransformReactivePowerAfter_Xpath']
        #无功功率输入框
        self.DistributionTransformReactivePower_name=self.Data['DistributionTransformReactivePower_name']
        #电压激活框
        self.DistributionTransformVoltBefor_Xpath=self.Data['DistributionTransformVoltBefor_Xpath']
        self.DistributionTransformVoltrAfter_Xpath=self.Data['DistributionTransformVoltrAfter_Xpath']
        #电压输入框
        self.DistributionTransformVolt_name=self.Data['DistributionTransformVolt_name']
        #点击向后的翻页
        self.Paing_ID=self.Data['Paing_ID']
        #点击向后的翻页(倒数第二页)
        self.PaingTwo_ID=self.Data['PaingTwo_ID']
        #点击日量测类型分组
        self.DistributionTransformDateGroup_Xpath=self.Data['DistributionTransformDateGroup_Xpath']
        #日配变激活框
        self.DistributionTransformActivationBefor=self.Data['DistributionTransformActivationBefor']
        self.DistributionTransformActivationAfter=self.Data['DistributionTransformActivationAfter']
        #日配变输入框
        self.DistributionTransformData_name=self.Data['DistributionTransformData_name']
        #数据输入后保存，以及弹出框内容校验
        self.SaveButton_ID = self.Data['SaveButton_ID']
        self.AlertMsg_ID = self.Data['AlertMsg_ID']
        self.AlertButton_ID = self.Data['AlertButton_ID']
    def DistributionlineMeasDayEnter_Fun(self,DistributionlineDayData,DistributionTransformDayData):
        '''配线和配变日运行数据录入
        DistributionlineDayData :配线日数据
        DistributionTransformDayData：配变日数据
        '''
        self.DistributionlineDayData=DistributionlineDayData
        self.DistributionTransformDayData=DistributionTransformDayData

        self.AlertMsgDay=[]

        #配线日有功电量
        ActivePower = self.DistributionlineDayData[0]
        PowerFactor = self.DistributionlineDayData[1]
        LoadShapeFactor = self.DistributionlineDayData[2]


        #配变日数据
        DistributionTransformDayData=self.DistributionTransformDayData[0:]
        #点击数据中心
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
        #点击f11全屏展示
        SendKeys.SendKeys("{F11}")
        time.sleep(2)
        ##理论线损运行数据录入页面frame
        self.driver.switch_to_frame(self.TheoreticalLineLoss_Frame)
        time.sleep(2)
        #点击量测类型分组框
        self.driver.find_element_by_id(self.CustomGroupBox_ID).click()
        time.sleep(2)
        #切换到日量测类型分组
        self.driver.find_element_by_xpath(self.DataCustomGroup_Xpath).click()
        time.sleep(2)
        #选择日期下拉框
        self.driver.find_element_by_id(self.DataDropDownBox_ID).click()
        time.sleep(2)
        #切换到日期frame
        self.driver.switch_to_frame(self.DateSelectionBox_Frame)
        time.sleep(2)
        #选择2014-8-1日
        self.driver.find_element_by_xpath(self.DateSelect_Xpath).click()
        time.sleep(2)
        #点击日期框中的确定按钮
        self.driver.find_element_by_id(self.DateAlerButton_ID).click()
        time.sleep(3)
        self.driver.switch_to.parent_frame()## 跳出当前iframe
        self.driver.implicitly_wait(10)
        #输入相关的参数信息
        #输入配线正向有功电量
        self.driver.find_element_by_xpath(self.ActivePowerActivation_Xpath).click()
        time.sleep(2)
        self.driver.find_element_by_name(self.ActivePower_name).clear()
        self.driver.find_element_by_name(self.ActivePower_name).send_keys(ActivePower)
        time.sleep(2)
        #输入配线功率因数
        self.driver.find_element_by_xpath(self.PowerFactorActivation_Xpath).click()
        time.sleep(2)
        self.driver.find_element_by_name(self.PowerFactor_name).clear()
        self.driver.find_element_by_name(self.PowerFactor_name).send_keys(PowerFactor)
        time.sleep(2)
        #输入配线负荷形状系数
        self.driver.find_element_by_xpath(self.LoadShapeFactorActivation_Xpath).click()
        time.sleep(2)
        self.driver.find_element_by_name(self.LoadShapeFactor_name).clear()
        self.driver.find_element_by_name(self.LoadShapeFactor_name).send_keys(LoadShapeFactor)
        self.driver.find_element_by_name(self.LoadShapeFactor_name).send_keys(Keys.TAB)
        time.sleep(3)
        #数据输入完成后，保存
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(2)
        self.AlertMsgDay1 = self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(2)

        self.AlertMsgDay.append(self.AlertMsgDay1)
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        time.sleep(3)

        #切换到配变tab页面
        self.driver.find_element_by_id(self.DistributionTransform_ID).click()
        time.sleep(5)

        #输入4个配变数据
        for i in range(2,6):
            self.driver.find_element_by_xpath(self.DistributionTransformActivationBefor+str(i)+self.DistributionTransformActivationAfter).click()
            time.sleep(2)
            self.driver.find_element_by_name(self.DistributionTransformData_name).clear()
            time.sleep(2)
            self.driver.find_element_by_name(self.DistributionTransformData_name).send_keys(DistributionTransformDayData[i-2])
            time.sleep(3)
        self.driver.find_element_by_name(self.DistributionTransformData_name).send_keys(Keys.TAB)
        time.sleep(2)
        #数据输入完成后
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(2)
        self.AlertMsgDay2 = self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(2)
        self.AlertMsgDay.append(self.AlertMsgDay2)
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        return json.dumps(self.AlertMsgDay, encoding="UTF-8", ensure_ascii=False)

    def DistributionlineMeasHourEnter_Fun(self,DistributionlineVoltHourData,DistributionlineElcHourData,DistributionTransformActivePowerHourData,DistributionTransformReactivePowerHourData,DistributionTransformVoltHourData):
        '''
        功能：配线和配变24点运行数据录入
        DistributionlineVoltHourData： 配线24点电压
        DistributionlineElcHourData：配线24点电流
        DistributionTransformActivePowerHourData：配变24点有功功率
        DistributionTransformReactivePowerHourData：配变24点无功功率
        DistributionTransformVoltHourData：配变24点电压
        '''
        self.DistributionlineVoltHourDate=DistributionlineVoltHourData
        self.DistributionlineElcHourData=DistributionlineElcHourData
        self.DistributionTransformActivePowerHourData=DistributionTransformActivePowerHourData
        self.DistributionTransformReactivePowerHourData=DistributionTransformReactivePowerHourData
        self.DistributionTransformVoltHourData=DistributionTransformVoltHourData


        self.AlertMsgHour=[]
        #返回信息集合

        #24点电压
        listvoltage = self.DistributionlineVoltHourDate[0:24]
        #24点电流
        listelec = self.DistributionlineElcHourData[0:24]
        #配变第一页有功功率
        FirstActivePower=self.DistributionTransformActivePowerHourData[0:25]
        #配变第二页有功功率
        SecondActivePower=self.DistributionTransformActivePowerHourData[25:50]
        #配变第三页有功功率
        ThirdActivePower=self.DistributionTransformActivePowerHourData[50:75]
        #配变第四页有功功率
        FourActivePower=self.DistributionTransformActivePowerHourData[75:96]
        #配变第一页无功功率
        FirstReactivePower=self.DistributionTransformReactivePowerHourData[0:25]
        #配变第二页无功功率
        SecondReactivePower=self.DistributionTransformReactivePowerHourData[25:50]
        #配变第三页有功功率
        ThirdReactivePower=self.DistributionTransformReactivePowerHourData[50:75]
        #配变第四页有功功率
        FourReactivePower=self.DistributionTransformReactivePowerHourData[75:96]
        #配变第一页电压
        FirstVolt=self.DistributionTransformVoltHourData[0:25]
        #配变第二页电压
        SecondVolt=self.DistributionTransformVoltHourData[25:50]
        #配变第三页电压
        ThirdVolt=self.DistributionTransformVoltHourData[50:75]
        #配变第四页电压
        FourVolt=self.DistributionTransformVoltHourData[75:96]
        #点击数据中心
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
        #点击f11全屏展示
        SendKeys.SendKeys("{F11}")
        time.sleep(2)
        ##理论线损运行数据录入页面frame
        self.driver.switch_to_frame(self.TheoreticalLineLoss_Frame)
        time.sleep(2)
        #点击量测类型分组
        self.driver.find_element_by_id(self.CustomGroupBox_ID).click()
        time.sleep(5)
        #切换到24点电流量测类型分组
        self.driver.find_element_by_xpath(self.EleCut24_Xpath).click()
        time.sleep(2)
        #输入配线24点电压电流数据
        for i in range(2,26):
            self.driver.find_element_by_xpath(self.VolBefor+str(i)+self.VolAfter).click()
            time.sleep(1)
            self.driver.find_element_by_name(self.VolDataEnter).clear()
            time.sleep(0.3)
            self.driver.find_element_by_name(self.VolDataEnter).send_keys(listvoltage[i-2])
            time.sleep(1)
            self.driver.find_element_by_xpath(self.EleCutBefor+str(i)+self.EleCutAfter).click()
            time.sleep(1)
            self.driver.find_element_by_name(self.EleCutDataEnter).clear()
            time.sleep(0.3)
            self.driver.find_element_by_name(self.EleCutDataEnter).send_keys(listelec[i-2])
            time.sleep(1)
            SendKeys.SendKeys("TAB")
            time.sleep(0.3)
            SendKeys.SendKeys("TAB")
            time.sleep(0.3)
            SendKeys.SendKeys("TAB")
            time.sleep(0.3)
        self.driver.find_element_by_name(self.EleCutDataEnter).send_keys(Keys.TAB)
        time.sleep(2)
        #数据输入完成后
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(2)
        self.AlertMsgHour1 = self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(2)
        self.AlertMsgHour.append(self.AlertMsgHour1)
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        time.sleep(3)

        #切换到配变tab页面
        self.driver.find_element_by_id(self.DistributionTransform_ID).click()
        time.sleep(5)
        #点击量测分组框
        self.driver.find_element_by_id(self.CustomGroupBox_ID).click()
        time.sleep(2)
        #点击“理论线损小时数据”
        self.driver.find_element_by_xpath(self.DistributionTransformHourData_Xpath).click()
        time.sleep(2)
        DistributionTransformActivePowerData=[FirstActivePower,SecondActivePower,ThirdActivePower]
        DistributionTransformReactivePowerData=[FirstReactivePower,SecondReactivePower,ThirdReactivePower]
        DistributionTransformVoltData=[FirstVolt,SecondVolt,ThirdVolt]
        for j in range(0,3):
            for i in range(2,27):
                self.driver.find_element_by_xpath(self.DistributionTransformActivePowerBefor_Xpath+str(i)+self.DistributionTransformActivePowerAfter_Xpath).click()
                time.sleep(1)
                self.driver.find_element_by_name(self.DistributionTransformActivePower_name).clear()
                time.sleep(0.3)
                self.driver.find_element_by_name(self.DistributionTransformActivePower_name).send_keys(DistributionTransformActivePowerData[j][i-2])
                time.sleep(1)
                self.driver.find_element_by_xpath(self.DistributionTransformReactivePowerBefor_Xpath+str(i)+self.DistributionTransformReactivePowerAfter_Xpath).click()
                time.sleep(1)
                self.driver.find_element_by_name(self.DistributionTransformReactivePower_name).clear()
                time.sleep(0.3)
                self.driver.find_element_by_name(self.DistributionTransformReactivePower_name).send_keys(DistributionTransformReactivePowerData[j][i-2])
                time.sleep(0.4)
                self.driver.find_element_by_xpath(self.DistributionTransformVoltBefor_Xpath+str(i)+self.DistributionTransformVoltrAfter_Xpath).click()
                time.sleep(1)
                self.driver.find_element_by_name(self.DistributionTransformVolt_name).clear()
                time.sleep(0.3)
                self.driver.find_element_by_name(self.DistributionTransformVolt_name).send_keys(DistributionTransformVoltData[j][i-2])
                time.sleep(0.4)
                SendKeys.SendKeys("TAB")
                time.sleep(0.3)
                SendKeys.SendKeys("TAB")
                time.sleep(0.3)
            self.driver.find_element_by_name(self.DistributionTransformVolt_name).send_keys(Keys.TAB)
            time.sleep(2)
            #数据输入完成后
            self.driver.find_element_by_id(self.SaveButton_ID).click()
            time.sleep(3)
            self.AlertMsgHour2 = self.driver.find_element_by_id(self.AlertMsg_ID).text
            time.sleep(3)
            self.AlertMsgHour.append(self.AlertMsgHour2)
            self.driver.find_element_by_id(self.AlertButton_ID).click()
            time.sleep(5)
            #点击向后的翻页
            self.driver.find_element_by_id(self.Paing_ID).click()
            time.sleep(5)
        #输入第四页有功功率、无功功率、电压数据
        for i in range(2,23):
            self.driver.find_element_by_xpath(self.DistributionTransformActivePowerBefor_Xpath+str(i)+self.DistributionTransformActivePowerAfter_Xpath).click()
            time.sleep(1)
            self.driver.find_element_by_name(self.DistributionTransformActivePower_name).clear()
            time.sleep(0.3)
            self.driver.find_element_by_name(self.DistributionTransformActivePower_name).send_keys(FourActivePower[i-2])
            time.sleep(1)
            self.driver.find_element_by_xpath(self.DistributionTransformReactivePowerBefor_Xpath+str(i)+self.DistributionTransformReactivePowerAfter_Xpath).click()
            time.sleep(1)
            self.driver.find_element_by_name(self.DistributionTransformReactivePower_name).clear()
            time.sleep(0.3)
            self.driver.find_element_by_name(self.DistributionTransformReactivePower_name).send_keys(FourReactivePower[i-2])
            time.sleep(0.4)
            self.driver.find_element_by_xpath(self.DistributionTransformVoltBefor_Xpath+str(i)+self.DistributionTransformVoltrAfter_Xpath).click()
            time.sleep(1)
            self.driver.find_element_by_name(self.DistributionTransformVolt_name).clear()
            time.sleep(0.3)
            self.driver.find_element_by_name(self.DistributionTransformVolt_name).send_keys(FourVolt[i-2])
            time.sleep(0.4)
            SendKeys.SendKeys("TAB")
            time.sleep(0.3)
            SendKeys.SendKeys("TAB")
            time.sleep(0.3)
        self.driver.find_element_by_name(self.DistributionTransformVolt_name).send_keys(Keys.TAB)
        time.sleep(2)
        #数据输入完成后，点击保存
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(2)
        self.AlertMsgHour3 = self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(2)
        self.AlertMsgHour.append(self.AlertMsgHour3)
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        return json.dumps(self.AlertMsgHour, encoding="UTF-8", ensure_ascii=False)

if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=DistributionlineMeasDayEnter(driver)
    TryReturn = Try1.DistributionlineMeasDayEnter_Fun(('6666','0.23','1.23'),('221','222','223','225'))
    driver.quit()
    driver = webdriver.Chrome()
    Try2=Login.Login(driver)
    Try2.Login_Function('lj','000000')
    Try2=DistributionlineMeasDayEnter(driver)
    TryReturn2 = Try2.DistributionlineMeasHourEnter_Fun(('10.1','10.1','10.1','10.2','10.3','10.1','10.1','10.1','10.1','10.1','10.1','10.5','10.1','10.1','10.7','10.1','10.1','10.1','10.44','10.1','10.1','10.1','10.1','10.1'),('1.25','2.56','3.26','2.26','3.26','4.25','3.25','6.25','3.26','1.25','1.25','4.26','1.25','1.25','6.25','1.25','1.25','8.26','9.26','1.25','1.25','1.25','10.23','1.25'
),('0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.3653','0.3653','0.3653','0.3653','0.3653','0.3653','0.3653','0.3653','0.3653','0.3653','0.3653','0.3653','0.4653','0.4653','0.4653','0.4653','0.4653','0.4653','0.4653','0.4653','0.4653','0.4653','0.4653','0.4653','0.5653','0.5653','0.5653','0.5653','0.5653','0.5653','0.5653','0.5653','0.5653','0.5653','0.5653','0.5653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653','0.2653'),('0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.7954','0.7954','0.7954','0.7954','0.7954','0.7954','0.7954','0.7954','0.7954','0.7954','0.7954','0.7954','0.6954','0.6954','0.6954','0.6954','0.6954','0.6954','0.6954','0.6954','0.6954','0.6954','0.6954','0.6954','0.3954','0.3954','0.3954','0.3954','0.3954','0.3954','0.3954','0.3954','0.3954','0.3954','0.3954','0.3954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954','0.8954'),('0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4','0.4'))
    time.sleep(5)
    driver.quit()