#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018-8-28 16:32
当前项目名称  ：Auto
功能          ：配线日电量法变压器损耗链接计算结果验证
-------------------------------------漂亮的分割线----------------------------------------------'''
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

class DistTransformerEleQuaDayCalResult:
    def __init__(self,driver):
        '''配线日电量法变压器损耗表链接页面结果验证'''
        #self.driver=webdriver.Chrome()
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistTransformerEleQuaDayCalResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_Data['DistTransformerEleQuaDayCalResult']

        #点击理论线损计算
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #点击配电线路计算结果查询
        self.TheoreticalLineLossManualcalResult_Xpath=self.Data['TheoreticalLineLossManualcalResult_Xpath']
        ##点击日报表
        self.TheoreticalLineloss620kVMonth_Xpath=self.Data['TheoreticalLineloss620kVDay_Xpath']
        #点击配电线路线损表
        self.WiringReport_Xpath=self.Data['WiringReport_Xpath']
        ##切换到结果frame
        self.TheoreticalLinelossResult_frame=self.Data['TheoreticalLinelossResult_frame']
        ##点击变压器损耗链接
        self.Transformerlosslj_Xpath=self.Data['Transformerlosslj_Xpath']
        ##切换到导线结果frame
        self.TransformerlossCalcResult_frame=self.Data['TransformerlossCalcResult_frame']
        #导出按钮id
        self.Daochu_ID=self.Data['Daochu_ID']

    def DistTransformerEleQuaDayCalResult_Fun(self):
        '''配电日电量法变压器损耗表链接页面结果验证'''
        #点击理论线损计算
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #点击配电线路计算结果查询
        self.driver.find_element_by_xpath(self.TheoreticalLineLossManualcalResult_Xpath).click()
        time.sleep(2)
        #点击月报表
        self.driver.find_element_by_xpath(self.TheoreticalLineloss620kVMonth_Xpath).click()
        time.sleep(2)
        #点击配电线路线损表
        self.driver.find_element_by_xpath(self.WiringReport_Xpath).click()
        time.sleep(2)
        ###切换到结果frame
        self.driver.switch_to_frame(self.TheoreticalLinelossResult_frame)
        time.sleep(2)
        self.driver.implicitly_wait(30)
        ##点击变压器损耗链接
        self.driver.find_element_by_xpath(self.Transformerlosslj_Xpath).click()
        time.sleep(3)
        self.driver.switch_to.parent_frame()## 跳出当前iframe
        self.driver.implicitly_wait(10)
        ##切换到变压器损耗明细结果frame
        self.driver.switch_to_frame(self.TransformerlossCalcResult_frame)
        self.driver.implicitly_wait(30)
        #导出按钮id
        self.driver.find_element_by_id(self.Daochu_ID).click()
        time.sleep(10)
        self.data_path= u"C:\\Users\\"+GetUsername()+ u"\\Downloads\\配电变压器损耗表报表（日）.xls"
        self.sheetname= u'配电变压器损耗表报表（日）'
        self.Excel=ExcelData(self.data_path,self.sheetname)
        self.ExcelDataa=self.Excel.readExcel()
        return self.ExcelDataa
if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lyltest','000000')
    Try1=DistTransformerEleQuaDayCalResult(driver)
    aaa =  Try1.DistTransformerEleQuaDayCalResult_Fun()
    data_path= u'D:\\svn\\eagle2\\05Test\\Auto\\Project\\EagleProjecttest\\Data\\eagle620kv\\配电变压器损耗表报表（日）.xls'
    sheetname = u'配电变压器损耗表报表（日）'
    Excel = ExcelData(data_path,sheetname)
    bbb = Excel.readExcel()
    a = cmp(aaa,bbb)
    print a
    time.sleep(10)
    driver.quit()
