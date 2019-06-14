# -*- coding: utf-8 -*-
#!/usr/bin/env python
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：吴鹏飞
IDE           ：PyCharm Community Edition
时间          ：2018/10/23 15:58
当前项目名称  ：eagle2
功能          ：
-------------------------------------漂亮的分割线---------------------------------------------'''
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
import yaml
import time
import Login
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class TranotherlossmonCalcResult:
    def __init__(self,driver):
        #驱动
        self.driver=driver
        #获取页面元素
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_file=open(self.ProjectFilePath+"\\Page_object\\Data\\TranotherlossmonCalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_file)
        self.Data=self.Page_Data['TranotherlossmonCalcResult']
        #理论线损id
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #输电网计算结果查询
        self.TranotherlossManualcal_Xpath=self.Data['TranotherlossManualcal_Xpath']
        #月报表
        self.TranotherlossmonReport_Xpath=self.Data['TranotherlossmonReport_Xpath']
        #其它损耗表
        self.TranotherlossmonCalcResult_Xpath=self.Data['TranotherlossmonCalcResult_Xpath']
        #查询报表frame
        self.TranotherlossmonCalcResult_frame=self.Data['TranotherlossmonCalcResult_frame']
        #点击导出
        self.Daochu_id=self.Data['Daochu_id']
    def TranotherlossmonCalcResult_Fun(self):
        #进入理论线损模块
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #进入输电网计算结果查询
        self.driver.find_element_by_xpath(self.TranotherlossManualcal_Xpath).click()
        time.sleep(2)
        #进入月报表
        self.driver.find_element_by_xpath(self.TranotherlossmonReport_Xpath).click()
        time.sleep(2)
        #进入其他损耗表
        self.driver.find_element_by_xpath(self.TranotherlossmonCalcResult_Xpath).click()
        time.sleep(5)
        #进入查询frame框
        self.driver.switch_to_frame(self.TranotherlossmonCalcResult_frame)
        time.sleep(3)
        #点击导出
        self.driver.find_element_by_id(self.Daochu_id).click()
        time.sleep(10)
        self.data_file=u'C:\\Users\\Admin\Downloads\输电网其他损耗报表（月）.xls'
        self.sheetname=u'输电网其他损耗报表（月）'
        self.Excel=ExcelData(self.data_file,self.sheetname)
        self.Excel_Data=self.Excel.readExcel()
        return self.Excel_Data
if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=TranotherlossmonCalcResult(driver)
    Try1.TranotherlossmonCalcResult_Fun()
    time.sleep(10)
    driver.quit()