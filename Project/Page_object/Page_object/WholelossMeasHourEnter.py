# -*- coding: utf-8 -*-
#!/usr/bin/env python
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：吴鹏飞
IDE           ：PyCharm Community Edition
时间          ：2018/10/17 14:10
当前项目名称  ：eagle2
功能          ：
-------------------------------------漂亮的分割线---------------------------------------------'''
from selenium import webdriver
import yaml
import time
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from selenium.webdriver.common.keys import Keys
import Login
import json
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class WholelossMeasHourEnter:
    #输电网计算运行数据录入
    def __init__(self,driver):
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        #获取页面元素
        self.Page_object_data_file=open(self.ProjectFilePath+"\\Page_object\\Data\\WholelossMeasHourEnter.yaml")
        self.Page_data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.date=self.Page_data['WholelossMeasHourEnter']
        #数据中心id
        self.DataCenter_id=self.date['DataCenter_id']
        #运行数据管理路径
        self.RunDataManage_Xpath=self.date['RunDataManage_Xpath']
        #输电网运行数据录入
        self.TransmissiongridDataEntry_Xpath=self.date['TransmissiongridDataEntry_Xpath']
        #输电网理论线损运行数据录入
        self.TransmissiongridLossDataEntry_Xpath=self.date['TransmissiongridLossDataEntry_Xpath']
        #输电网理论运行数据录入frame
        self.TransmissiongridLossDataEntry_Frame=self.date['TransmissiongridLossDataEntry_Frame']
        #下拉框
        self.CustomGroupBox_ID=self.date['CustomGroupBox_ID']
        #日量测类型分组
        self.DataCustomGroup_Xpath=self.date['DataCustomGroup_Xpath']
        #日期下拉框
        self.DataDropDownBox_ID=self.date['DataDropDownBox_ID']
        #日期框frame
        self.DateSelectionBox_Frame=self.date['DateSelectionBox_Frame']
        #选择日期20140801
        self.DateSelect_Xpath=self.date['DateSelect_Xpath']
        #点击日期框中确定
        self.DateAlerButton_ID=self.date['DateAlerButton_ID']
        #负荷tab页面
        self.EnergyConsumer_ID=self.date['EnergyConsumer_ID']
        #
        self.EnergyConsumerBefor=self.date['EnergyConsumerBefor']
        self.EnergyConsumerAfter=self.date['EnergyConsumerAfter']
        self.EnergyConsumerQBefor=self.date['EnergyConsumerQBefor']
        self.EnergyConsumerQAfter=self.date['EnergyConsumerQAfter']
        #负荷有功功率
        self.ActivePowerEnter=self.date['ActivePowerEnter']
        #负荷无功功率
        self.ReactivePowerEnter=self.date['ReactivePowerEnter']


        #保存按钮
        self.SaveButton_ID=self.date['SaveButton_ID']
        #保存弹出框信息
        self.AlertMsg_ID=self.date['AlertMsg_ID']
        #确定按钮
        self.AlertButton_ID=self.date['AlertButton_ID']


        #平衡节点tab页面
        self.GeneratingUnit_ID=self.date['GeneratingUnit_ID']
        #平衡节点电压
        self.GeneratingUnitBefor=self.date['GeneratingUnitBefor']
        self.GeneratingUnitAfter=self.date['GeneratingUnitAfter']
        self.VoltageEnter=self.date['VoltageEnter']
        #电容tab页面
        self.ReactorBank_id=self.date['ReactorBank_id']
        self.ReactorBankBefor=self.date['ReactorBankBefor']
        self.ReactorBankAfter=self.date['ReactorBankAfter']
        self.TransportationCapacityEnter=self.date['TransportationCapacityEnter']


    def WholelossMeasHourEnter_Fun(self,*args):
        #数据中心
        '''输电网计算所用到的运行数据录入'''
        self.driver.find_element_by_id(self.DataCenter_id).click()
        time.sleep(2)
        #运行数据管理
        self.driver.find_element_by_xpath(self.RunDataManage_Xpath).click()
        time.sleep(2)
        #输电网运行数据录入
        self.driver.find_element_by_xpath(self.TransmissiongridDataEntry_Xpath).click()
        time.sleep(2)
        #输电网理论运行数据录入
        self.driver.find_element_by_xpath(self.TransmissiongridLossDataEntry_Xpath).click()
        time.sleep(2)

        #输电网运行数据frame
        self.driver.switch_to_frame(self.TransmissiongridLossDataEntry_Frame)
        time.sleep(2)
        #查询条件下拉框
        self.driver.find_element_by_id(self.CustomGroupBox_ID).click()
        time.sleep(2)
        #选择理论日量测类型分组
        self.driver.find_element_by_xpath(self.DataCustomGroup_Xpath).click()
        time.sleep(2)
        #点击日期下拉框
        self.driver.find_element_by_id(self.DataDropDownBox_ID).click()
        time.sleep(2)
        #进入日期frame
        self.driver.switch_to_frame(self.DateSelectionBox_Frame)
        time.sleep(2)
        #选择日期20140801
        self.driver.find_element_by_xpath(self.DateSelect_Xpath).click()
        time.sleep(2)
        #点击日期框确定
        self.driver.find_element_by_id(self.DateAlerButton_ID).click()
        time.sleep(2)
        #跳出日期frame
        self.driver.switch_to.parent_frame()
        #切换到负荷tab页面
        self.driver.find_element_by_id(self.EnergyConsumer_ID).click()
        time.sleep(2)


        #self.args = args
        self.args = args[0].split(",")
        print self.args
        #负荷24点有功功率
        ActivePower=self.args[0:24]
        #负荷24点无功功率
        ReactivePower=self.args[24:48]
        #平衡节点电压
        Voltage=self.args[48:72]
        #电容投运容量
        TransportationCapacity=self.args[72:96]
        self.AlertMsg=[]

        for i in range(2,26):
            self.driver.find_element_by_xpath(self.EnergyConsumerBefor+str(i)+self.EnergyConsumerAfter).click()
            time.sleep(2)
            self.driver.find_element_by_name(self.ActivePowerEnter).clear()
            time.sleep(2)
            self.driver.find_element_by_name(self.ActivePowerEnter).send_keys(ActivePower[i-2])
            time.sleep(2)
            self.driver.find_element_by_xpath(self.EnergyConsumerQBefor+str(i)+self.EnergyConsumerQAfter).click()
            time.sleep(2)
            self.driver.find_element_by_name(self.ReactivePowerEnter).clear()
            time.sleep(2)
            self.driver.find_element_by_name(self.ReactivePowerEnter).send_keys(ReactivePower[i-2])
            time.sleep(2)
            self.driver.find_element_by_name(self.ReactivePowerEnter).send_keys(Keys.TAB)
            time.sleep(2)
        #点击保存，获取提示信息
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(2)
        self.AlertMsg1=self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(3)
        self.AlertMsg.append(self.AlertMsg1)
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        time.sleep(2)


        #切换到平衡节点tab页面
        self.driver.find_element_by_id(self.GeneratingUnit_ID).click()
        time.sleep(2)
        for i in range(2,26):
            self.driver.find_element_by_xpath(self.GeneratingUnitBefor+str(i)+self.GeneratingUnitAfter).click()
            time.sleep(2)
            self.driver.find_element_by_name(self.VoltageEnter).clear()
            time.sleep(2)
            self.driver.find_element_by_name(self.VoltageEnter).send_keys(Voltage[i-2])
            time.sleep(2)
            self.driver.find_element_by_name(self.VoltageEnter).send_keys(Keys.TAB)
            time.sleep(2)
        #点击保存，获取提示信息
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(2)
        self.AlertMsg2=self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(3)
        self.AlertMsg.append(self.AlertMsg2)
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        time.sleep(3)


        #切换到电容tab页面
        self.driver.find_element_by_id(self.ReactorBank_id).click()
        time.sleep(2)
        for i in range(2,26):
            self.driver.find_element_by_xpath(self.ReactorBankBefor+str(i)+self.ReactorBankAfter).click()
            time.sleep(2)
            self.driver.find_element_by_name(self.TransportationCapacityEnter).clear()
            time.sleep(2)
            self.driver.find_element_by_name(self.TransportationCapacityEnter).send_keys(TransportationCapacity[i-2])
            time.sleep(2)
            self.driver.find_element_by_name(self.TransportationCapacityEnter).send_keys(Keys.TAB)
            time.sleep(2)

        #点击保存，获取提示信息
        self.driver.find_element_by_id(self.SaveButton_ID).click()
        time.sleep(2)
        self.AlertMsg3=self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(3)
        self.AlertMsg.append(self.AlertMsg3)
        self.driver.find_element_by_id(self.AlertButton_ID).click()
        time.sleep(2)
        return json.dumps(self.AlertMsg, encoding="UTF-8", ensure_ascii=False)
        #print json.dumps(self.AlertMsg, encoding="UTF-8", ensure_ascii=False)



if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=WholelossMeasHourEnter(driver)
    #TryReturn = Try1.WholelossMeasHourEnter_Fun('11','22')
    #TryReturn = Try1.WholelossMeasHourEnter_Fun('6666','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','6666','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','6666','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','6666','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223','225','0.23','1.23','221','222','223')
    aaa = "0.0694,0.0621,0.0621,0.0658,0.0585,0.084,0.106,0.676,0.7454,1.0634,0.6322,0.1315,0.095,0.6395,0.5372,0.5372,0.5956,0.5152,0.1206,0.2521,0.2631,0.1754,0.0731,0.0658,0.1023,0.106,0.106,0.1096,0.095,0.1206,0.1206,0.983,1.1145,1.1035,0.9757,0.19,0.1571,1.0487,0.9208,1.0305,0.9976,0.9135,0.1937,0.4421,0.4348,0.3143,0.0987,0.1023,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,0.43,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,100"

    TryReturn = Try1.WholelossMeasHourEnter_Fun(aaa)
    print TryReturn
    time.sleep(5)
    driver.quit()





