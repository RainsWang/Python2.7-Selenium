#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/10/23 16:21
当前项目名称  ：AutoTest
功能          ：
-------------------------------------漂亮的分割线----------------------------------------------'''
#!/usr/bin/env python
# -*- coding: utf-8 -*-
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

class NewtonMethodDataEnter:
    #录入主网计算所需要用到的运行数据
    def __init__(self,driver):
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()

        #获取页面元素
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\NewtonMethodDataEnter.yaml")
        self.Page_data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_data['NewtonMethodDataEnter']
        #数据中心id
        self.DataCenter_ID = self.Data['DataCenter_ID']
        #运行数据管理绝对路径
        self.RunDataManage_Xpath = self.Data['RunDataManage_Xpath']
        #输电网运行数据录入绝对路径
        self.TransDataEnter_Xpath = self.Data['TransDataEnter_Xpath']
        #理论线损运行数据录入绝对路径
        self.LineLossDataEnter_Xpath = self.Data['LineLossDataEnter_Xpath']
        #页面收缩按钮id
        self.splitter_ID = self.Data['splitter_ID']

        #理论线损运行数据录入报表查询页面iframe的id
        self.LineLossDataEnterIframe_ID = self.Data['LineLossDataEnterIframe_ID']
        #日期选择框下拉按钮绝对路径
        self.DateDropdownBox_Xpath = self.Data['DateDropdownBox_Xpath']
        #日期选择框iframe的id
        self.Date_ID = self.Data['Date_ID']
        #日期选择框日期1号绝对路径
        self.Date1_Xpath = self.Data['Date1_Xpath']
        #日期选择确定按钮绝对路径
        self.DateDetermineButton_Xpath = self.Data['DateDetermineButton_Xpath']
        #日期选择完成

        #切换到负荷tab
        #负荷tab页绝对路径
        self.LoadTab_Xpath = self.Data['LoadTab_Xpath']

        #有功功率
        self.ActivePower = self.Data['ActivePower']
        #无功功率
        self.ReactivePower = self.Data['ReactivePower']
        #电压
        self.Volt = self.Data['Volt']

        #负荷有功功率运行数据录入绝对路径
        self.LoadPBefor_Xpath = self.Data['LoadPBefor_Xpath']
        self.LoadPAfter_Xpath = self.Data['LoadPAfter_Xpath']

        #负荷无功功率运行数据录入绝对路径
        self.LoadQBefor_Xpath = self.Data['LoadQBefor_Xpath']
        self.LoadQAfter_Xpath = self.Data['LoadQAfter_Xpath']

        #平衡节点电压运行数据录入绝对路径
        self.BalanceNodeVoltBefor_Xpath = self.Data['BalanceNodeVoltBefor_Xpath']
        self.BalanceNodeVoltAfter_Xpath = self.Data['BalanceNodeVoltAfter_Xpath']

        #运行数据保存按钮ID
        self.DataSaveButton_ID = self.Data['DataSaveButton_ID']
        #运行数据保存成功信息提示ID
        self.ConfirmMsg_ID = self.Data['ConfirmMsg_ID']
        #运行数据保存成功确定按钮ID
        self.ConfirmMsgButton_ID = self.Data['ConfirmMsgButton_ID']

        #变电1负荷2页面绝对路径
        self.Sub1Load2_Xpath = self.Data['Sub1Load2_Xpath']
        #变电2负荷1页面绝对路径
        self.Sub2Load1_Xpath = self.Data['Sub2Load1_Xpath']
        #变电2负荷2页面绝对路径
        self.Sub2Load2_Xpath = self.Data['Sub2Load2_Xpath']

        #切换到PQ发电机tab
        #PQ发电机tab页绝对路径
        self.PQTab_Xpath = self.Data['PQTab_Xpath']

        #切换到平衡节点tab
        #平衡节点tab页绝对路径
        self.BalanceNodeTab_Xpath = self.Data['BalanceNodeTab_Xpath']
    def NewtonMethodDataEnter_Fun(self,LoadP,LoadQ,PowerSupplyP,PowerSupplyQ,PowerSupplyV):
        self.LoadP=LoadP
        self.LoadQ=LoadQ
        self.PowerSupplyP=PowerSupplyP
        self.PowerSupplyQ=PowerSupplyQ
        self.PowerSupplyV=PowerSupplyV
        Sub1Load1P=self.LoadP[0:24]   #变电站1负荷1有功功率
        Sub1Load2P=self.LoadP[24:48]   #变电站1负荷2有功功率
        Sub2Load1P=self.LoadP[48:72]   #变电站2负荷1有功功率
        Sub2Load2P=self.LoadP[72:]   #变电站2负荷2有功功率
        Sub1Load1Q=self.LoadQ[0:24]   #变电站1负荷1无功功率
        Sub1Load2Q=self.LoadQ[24:48]   #变电站1负荷2无功功率
        Sub2Load1Q=self.LoadQ[48:72]   #变电站2负荷1无功功率
        Sub2Load2Q=self.LoadQ[72:]   #变电站2负荷2无功功率
        PowerSupplyP=self.PowerSupplyP[0:]   #变电站1电源1有功功率
        PowerSupplyQ=self.PowerSupplyQ[0:]   #变电站1电源1无功功率
        PowerSupplyV=self.PowerSupplyV[0:]   #变电站1电源2电压

        #点击数据中心
        self.driver.find_element_by_id(self.DataCenter_ID).click()
        time.sleep(2)
        #点击运行数据管理
        self.driver.find_element_by_xpath(self.RunDataManage_Xpath).click()
        time.sleep(2)
        #点击输电网运行数据录入
        self.driver.find_element_by_xpath(self.TransDataEnter_Xpath).click()
        time.sleep(2)
        #点击理论线损运行数据录入
        self.driver.find_element_by_xpath(self.LineLossDataEnter_Xpath).click()
        time.sleep(15)
        #已打开数据中心►运行数据管理►输电网运行数据录入►理论线损运行数据录入报表

        #点击页面上方的收缩按钮
        self.driver.find_element_by_id(self.splitter_ID).click()

        #默认量测分组,为理论线损小时数据，不用修改
        #选择日期,为2014-08-01
        self.driver.switch_to_frame(self.LineLossDataEnterIframe_ID)
        self.driver.find_element_by_xpath(self.DateDropdownBox_Xpath).click()
        time.sleep(1)
        self.driver.switch_to_frame(self.Date_ID)
        self.driver.find_element_by_xpath(self.Date1_Xpath).click()
        time.sleep(1)
        self.driver.find_element_by_xpath(self.DateDetermineButton_Xpath).click()
        time.sleep(2)

        #切换到负荷tab页面
        self.driver.switch_to.parent_frame()
        self.driver.find_element_by_xpath(self.LoadTab_Xpath).click()
        time.sleep(3)
        self.ConfirmMsg=[]
        #开始录入变电站1负荷1的运行数据
        #self.activepower=[2.115,2.116,2.117,2.118,2.119,2.12,2.121,2.122,2.123,2.124,2.125,2.126,2.127,2.128,2.129,2.13,2.131,2.132,2.133,2.134,2.135,2.136,2.137,2.138]
        #self.reactivepower=[0.125,0.126,0.127,0.128,0.129,0.13,0.131,0.132,0.133,0.134,0.135,0.136,0.137,0.138,0.139,0.14,0.141,0.142,0.143,0.144,0.145,0.146,0.147,0.148]
        for i in range(2,26):
            self.driver.find_element_by_xpath(self.LoadPBefor_Xpath+str(i)+self.LoadPAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).send_keys(str(Sub1Load1P[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_xpath(self.LoadQBefor_Xpath+str(i)+self.LoadQAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(str(Sub1Load1Q[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(Keys.TAB)
            time.sleep(0.1)
        self.driver.find_element_by_id(self.DataSaveButton_ID).click()
        time.sleep(2)
        self.ConfirmMsg1=self.driver.find_element_by_id(self.ConfirmMsg_ID).text
        self.ConfirmMsg.append('变电站1-负荷1:'+self.ConfirmMsg1)
        self.driver.find_element_by_id(self.ConfirmMsgButton_ID).click()
        time.sleep(1)

        #开始录入变电站1负荷2的运行数据
        self.driver.find_element_by_xpath(self.Sub1Load2_Xpath).click()
        time.sleep(2)
        #self.activepower=[2.215,2.216,2.217,2.218,2.219,2.22,2.221,2.222,2.223,2.224,2.225,2.226,2.227,2.228,2.229,2.23,2.231,2.232,2.233,2.234,2.235,2.236,2.237,2.238]
        #self.reactivepower=[0.225,0.226,0.227,0.228,0.229,0.23,0.231,0.232,0.233,0.234,0.235,0.236,0.237,0.238,0.239,0.24,0.241,0.242,0.243,0.244,0.245,0.246,0.247,0.248]
        for i in range(2,26):
            self.driver.find_element_by_xpath(self.LoadPBefor_Xpath+str(i)+self.LoadPAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).send_keys(str(Sub1Load2P[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_xpath(self.LoadQBefor_Xpath+str(i)+self.LoadQAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(str(Sub1Load2Q[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(Keys.TAB)
            time.sleep(0.1)
        self.driver.find_element_by_id(self.DataSaveButton_ID).click()
        time.sleep(2)
        self.ConfirmMsg2=self.driver.find_element_by_id(self.ConfirmMsg_ID).text
        self.ConfirmMsg.append('变电站1-负荷2:'+self.ConfirmMsg2)
        self.driver.find_element_by_id(self.ConfirmMsgButton_ID).click()
        time.sleep(3)

        #开始录入变电站2负荷1的运行数据
        self.driver.find_element_by_xpath(self.Sub2Load1_Xpath).click()
        time.sleep(2)
        # self.activepower=[1.515,1.516,1.517,1.518,1.519,1.52,1.521,1.522,1.523,1.524,1.525,1.526,1.527,1.528,1.529,1.53,1.531,1.532,1.533,1.534,1.535,1.536,1.537,1.538]
        # self.reactivepower=[1.025,1.026,1.027,1.028,1.029,1.03,1.031,1.032,1.033,1.034,1.035,1.036,1.037,1.038,1.039,1.04,1.041,1.042,1.043,1.044,1.045,1.046,1.047,1.048]
        for i in range(2,26):
            self.driver.find_element_by_xpath(self.LoadPBefor_Xpath+str(i)+self.LoadPAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).send_keys(str(Sub2Load1P[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_xpath(self.LoadQBefor_Xpath+str(i)+self.LoadQAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(str(Sub2Load1Q[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(Keys.TAB)
            time.sleep(0.1)
        self.driver.find_element_by_id(self.DataSaveButton_ID).click()
        time.sleep(2)
        self.ConfirmMsg3=self.driver.find_element_by_id(self.ConfirmMsg_ID).text
        self.ConfirmMsg.append('变电站2-负荷1:'+self.ConfirmMsg3)
        time.sleep(2)
        self.driver.find_element_by_id(self.ConfirmMsgButton_ID).click()
        time.sleep(3)
        #开始录入变电站2负荷2的运行数据
        self.driver.find_element_by_xpath(self.Sub2Load2_Xpath).click()
        time.sleep(2)
        # self.activepower=[1.815,1.816,1.817,1.818,1.819,1.82,1.821,1.822,1.823,1.824,1.825,1.826,1.827,1.828,1.829,1.83,1.831,1.832,1.833,1.834,1.835,1.836,1.837,1.838]
        # self.reactivepower=[0.425,0.426,0.427,0.428,0.429,0.43,0.431,0.432,0.433,0.434,0.435,0.436,0.437,0.438,0.439,0.44,0.441,0.442,0.443,0.444,0.445,0.446,0.447,0.448]
        for i in range(2,26):
            self.driver.find_element_by_xpath(self.LoadPBefor_Xpath+str(i)+self.LoadPAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).send_keys(str(Sub2Load2P[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_xpath(self.LoadQBefor_Xpath+str(i)+self.LoadQAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(str(Sub2Load2Q[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(Keys.TAB)
            time.sleep(0.1)
        self.driver.find_element_by_id(self.DataSaveButton_ID).click()
        time.sleep(2)
        self.ConfirmMsg4=self.driver.find_element_by_id(self.ConfirmMsg_ID).text
        self.ConfirmMsg.append('变电站2-负荷2:'+self.ConfirmMsg4)
        self.driver.find_element_by_id(self.ConfirmMsgButton_ID).click()
        time.sleep(3)

        #切换到PQ节点发电机tab页
        self.driver.find_element_by_xpath(self.PQTab_Xpath).click()
        time.sleep(3)
        #准备主网计算变电站1-电源1的运行数据，并录入保存
        # self.activepower=[4.81,4.82,4.83,4.84,4.85,4.86,4.87,4.88,4.89,4.9,4.91,4.9,4.89,4.88,4.87,4.86,4.85,4.84,4.83,4.82,4.81,4.8,4.79,4.78]
        # self.reactivepower=[1.825,1.826,1.827,1.828,1.829,1.83,1.831,1.832,1.833,1.834,1.835,1.836,1.837,1.838,1.839,1.84,1.841,1.842,1.843,1.844,1.845,1.846,1.847,1.848]
        for i in range(2,26):
            self.driver.find_element_by_xpath(self.LoadPBefor_Xpath+str(i)+self.LoadPAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ActivePower).send_keys(str(PowerSupplyP[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_xpath(self.LoadQBefor_Xpath+str(i)+self.LoadQAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(str(PowerSupplyQ[i-2]))
            time.sleep(0.1)
            self.driver.find_element_by_name(self.ReactivePower).send_keys(Keys.TAB)
            time.sleep(0.1)
        self.driver.find_element_by_id(self.DataSaveButton_ID).click()
        time.sleep(2)
        self.ConfirmMsg5=self.driver.find_element_by_id(self.ConfirmMsg_ID).text
        self.ConfirmMsg.append('变电站1-电源1:'+self.ConfirmMsg5)
        self.driver.find_element_by_id(self.ConfirmMsgButton_ID).click()
        time.sleep(3)
        #切换到平衡节点发电机tab页
        self.driver.find_element_by_xpath(self.BalanceNodeTab_Xpath).click()
        time.sleep(3)
        for i in range(2,26):
            self.driver.find_element_by_xpath(self.BalanceNodeVoltBefor_Xpath+str(i)+self.BalanceNodeVoltAfter_Xpath).click()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.Volt).clear()
            time.sleep(0.1)
            self.driver.find_element_by_name(self.Volt).send_keys(str(PowerSupplyV[i-2]))
            time.sleep(0.2)
        self.driver.find_element_by_name(self.Volt).send_keys(Keys.TAB)
        time.sleep(0.1)
        self.driver.find_element_by_id(self.DataSaveButton_ID).click()
        time.sleep(2)
        self.ConfirmMsg6=self.driver.find_element_by_id(self.ConfirmMsg_ID).text
        self.ConfirmMsg.append('变电站1-电源2:'+self.ConfirmMsg6)
        self.driver.find_element_by_id(self.ConfirmMsgButton_ID).click()
        time.sleep(3)
        return sys.stdout.write('\n'.join(self.ConfirmMsg))

if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('wjtest','000000')
    Func=NewtonMethodDataEnter(driver)
    TryReturn = Func.NewtonMethodDataEnter_Fun(('2.115','2.116','2.117','2.118','2.119','2.12','2.121','2.122','2.123','2.124','2.125','2.126','2.127','2.128','2.129','2.13','2.131','2.132','2.133','2.134','2.135','2.136','2.137','2.138','2.215','2.216','2.217','2.218','2.219','2.22','2.221','2.222','2.223','2.224','2.225','2.226','2.227','2.228','2.229','2.23','2.231','2.232','2.233','2.234','2.235','2.236','2.237','2.238','1.515','1.516','1.517','1.518','1.519','1.52','1.521','1.522','1.523','1.524','1.525','1.526','1.527','1.528','1.529','1.53','1.531','1.532','1.533','1.534','1.535','1.536','1.537','1.538','1.815','1.816','1.817','1.818','1.819','1.82','1.821','1.822','1.823','1.824','1.825','1.826','1.827','1.828','1.829','1.83','1.831','1.832','1.833','1.834','1.835','1.836','1.837','1.838'),('0.125','0.126','0.127','0.128','0.129','0.13','0.131','0.132','0.133','0.134','0.135','0.136','0.137','0.138','0.139','0.14','0.141','0.142','0.143','0.144','0.145','0.146','0.147','0.148','0.225','0.226','0.227','0.228','0.229','0.23','0.231','0.232','0.233','0.234','0.235','0.236','0.237','0.238','0.239','0.24','0.241','0.242','0.243','0.244','0.245','0.246','0.247','0.248','0.225','0.226','0.227','0.228','0.229','0.23','0.231','0.232','0.233','0.234','0.235','0.236','0.237','0.238','0.239','0.24','0.241','0.242','0.243','0.244','0.245','0.246','0.247','0.248','0.425','0.426','0.427','0.428','0.429','0.43','0.431','0.432','0.433','0.434','0.435','0.436','0.437','0.438','0.439','0.44','0.441','0.442','0.443','0.444','0.445','0.446','0.447','0.448'),('4.81','4.82','4.83','4.84','4.85','4.86','4.87','4.88','4.89','4.9','4.91','4.9','4.89','4.88','4.87','4.86','4.85','4.84','4.83','4.82','4.81','4.8','4.79','4.78'),('1.825','1.826','1.827','1.828','1.829','1.83','1.831','1.832','1.833','1.834','1.835','1.836','1.837','1.838','1.839','1.84','1.841','1.842','1.843','1.844','1.845','1.846','1.847','1.848'),('330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2','330.2'))
    time.sleep(1)
    driver.quit()