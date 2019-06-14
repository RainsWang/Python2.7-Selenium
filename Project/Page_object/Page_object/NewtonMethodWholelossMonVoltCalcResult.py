#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/12/25 9:13
当前项目名称  ：AutoTest
功能          ：月报表-输电网分压线损表
-------------------------------------漂亮的分割线----------------------------------------------'''
#!/usr/bin/env python
# -*- coding: utf-8 -*-
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
from Public_Theory.GetUsername import GetUsername
import yaml
import time
import Login
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class NewtonMethodWholelossMonVoltCalcResult:
    def __init__(self,driver):
        #驱动
        self.driver=driver
        #获取页面元素
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_file=open(self.ProjectFilePath+"\\Page_object\\Data\\NewtonMethodWholelossMonVoltCalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_file)
        self.Data=self.Page_Data['NewtonMethodWholelossMonVoltCalcResult']

        #理论线损ID
        self.LineLoss_ID = self.Data['LineLoss_ID']
        #输电网计算结果查询xpath
        self.WholelossCalcResultQuery_Xpath = self.Data['WholelossCalcResultQuery_Xpath']
        #月报表xpath
        self.MonCalcResult_Xpath = self.Data['MonCalcResult_Xpath']
        #输电网分压线损表xpath
        self.WholelossMonVoltCalcResult_Xpath = self.Data['WholelossMonVoltCalcResult_Xpath']
        #报表iframe的ID
        self.WholelossMonVoltCalcResultIframe_ID = self.Data['WholelossMonVoltCalcResultIframe_ID']

        #查询按钮ID
        self.Query_ID = self.Data['Query_ID']
        #导出按钮ID
        self.Export_ID = self.Data['Export_ID']

    def NewtonMethodWholelossMonVoltCalcResult_Fun(self,*args):
        #点击理论线损
        self.driver.find_element_by_id(self.LineLoss_ID).click()
        time.sleep(1)
        #点击输电网计算结果查询
        self.driver.find_element_by_xpath(self.WholelossCalcResultQuery_Xpath).click()
        time.sleep(1)
        #点击月报表
        self.driver.find_element_by_xpath(self.MonCalcResult_Xpath).click()
        time.sleep(1)
        #点击输电网分压线损表
        self.driver.find_element_by_xpath(self.WholelossMonVoltCalcResult_Xpath).click()
        time.sleep(5)

        #切换到报表的frame
        self.driver.switch_to_frame(self.WholelossMonVoltCalcResultIframe_ID)

        #点击查询
        self.driver.find_element_by_id(self.Query_ID).click()
        time.sleep(5)
        #点击导出
        self.driver.find_element_by_id(self.Export_ID).click()
        time.sleep(5)
        self.data_file=u'C:\\Users\\'+GetUsername()+ u'\\Downloads\\全网分电压损耗报表（月）.xls'
        self.sheetname=u'全网分电压损耗报表（月）'
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        #print self.Excel_Data
        return self.Excel_Data

if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('wjtest','000000')
    time.sleep(5)
    Func=NewtonMethodWholelossMonVoltCalcResult(driver)
    TryReturn = Func.NewtonMethodWholelossMonVoltCalcResult()
    time.sleep(10)
    driver.quit()