# -*- coding: utf-8 -*-
#!/usr/bin/env python
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：王敏
IDE           ：PyCharm
时间          ：2018-10-29 16:57
当前项目名称  ：Project
功能          ：台区电表损耗月报表结果
-------------------------------------漂亮的分割线---------------------------------------------'''
from selenium import webdriver
import yaml
import time
import Login
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
from Public_Theory.GetUsername import GetUsername
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

class DistrictEleQuamonMetercalResult:
    def __init__(self,driver):
        self.driver = driver
        #self.driver = webdriver.Chrome()
        self.ProjectFilePath = GetProjectFilePath()
        self.page_object_date_file = open(self.ProjectFilePath+"\Page_object\Data\DistrictEleQuamonMetercalResult.yaml")
        self.page_Data = yaml.load(self.page_object_date_file)
        self.page_object_date_file.close()
        self.Data = self.page_Data['DistrictEleQuamonMetercalResult']

        #点击理论线损
        self.TheoreticalLineLoss_ID = self.Data['TheoreticalLineLoss_ID']
        # 点击配电线路计算结果查询xpath
        self.TheoreticalLineLossManualcal_Xpath = self.Data['TheoreticalLineLossManualcal_Xpath']
        # 点击配线月报表
        self.DistributionLineMonthlyReport_Xpath = self.Data['DistributionLineMonthlyReport_Xpath']
        # 点击台区线损表
        self.LinelossmeterofDistrict_Xpath = self.Data['LinelossmeterofDistrict_Xpath']
        #点击电表损耗frame
        self.TheoreticalmeterResult_frame = self.Data['TheoreticalmeterResult_frame']
        ##点击电表链接
        self.Transformermeter_Xpath = self.Data['Transformermeter_Xpath']
        # 切换到结果frame
        self.DistrictCalcResult_frame = self.Data['DistrictCalcResult_frame']
        # 点击导出
        self.export_id = self.Data['export_id']

    def DistrictEleQuamonMetercalResult_Fun(self):
        # 点击理论线损计算
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        # 点击配电线路计算结果查询
        self.driver.find_element_by_xpath(self.TheoreticalLineLossManualcal_Xpath).click()
        time.sleep(2)
        # 点击月报表
        self.driver.find_element_by_xpath(self.DistributionLineMonthlyReport_Xpath).click()
        time.sleep(2)
        # 点击台区线路线损表
        self.driver.find_element_by_xpath(self.LinelossmeterofDistrict_Xpath).click()
        time.sleep(2)
        #切换到结果frame
        self.driver.switch_to_frame(self.TheoreticalmeterResult_frame)
        time.sleep(5)
        self.driver.implicitly_wait(30)
        ##点击电表损耗链接
        self.driver.find_element_by_xpath(self.Transformermeter_Xpath).click()
        time.sleep(5)
        self.driver.switch_to.parent_frame()
        # 跳出当前iframe
        self.driver.implicitly_wait(10)
        ##切换到电表损耗明细结果frame
        self.driver.switch_to_frame(self.DistrictCalcResult_frame)
        self.driver.implicitly_wait(30)
        # 导出按钮id
        self.driver.find_element_by_id(self.export_id).click()
        time.sleep(10)
        self.data_path = u'C:\\Users\\'+GetUsername()+u'\\Downloads\\低压台区电表损耗报表（月）.xls'
        self.sheetname = u'低压台区电表损耗报表（月）'
        self.Excel = ExcelData(self.data_path, self.sheetname)
        self.ExcelDataa = self.Excel.readExcel()
        return self.ExcelDataa
if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function(u'wm县', '000000')
    Try1 = DistrictEleQuamonMetercalResult(driver)
    Try1.DistrictEleQuamonMetercalResult_Fun()
    time.sleep(10)
    driver.quit()