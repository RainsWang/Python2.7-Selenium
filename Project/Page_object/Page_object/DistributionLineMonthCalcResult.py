#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/16 15:25
当前项目名称  ：620Calc
功能          ：月配线结果验证封装
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import time
import yaml
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
from Public_Theory.GetUsername import GetUsername
import Login
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

class DistributionlineEleQuamoncalcResult:
    def __init__(self,driver):
        '''620kv月电量法计算结果验证'''
        #self.driver=webdriver.Chrome()
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistributionlineEleQuamoncalcResult.yaml")
        self.Page_Data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_Data['DistributionlineEleQuamoncalcResult']

        #理论线损id
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #点击配电线路计算结果查询xpath
        self.TheoreticalLineLossManualcal_Xpath=self.Data['TheoreticalLineLossManualcal_Xpath']
        #点击配线月报表
        self.DistributionLineMonthlyReport_Xpath=self.Data['DistributionLineMonthlyReport_Xpath']
        #点击配电线路线损表
        self.Linelossmeterofdistributionline_Xpath=self.Data['Linelossmeterofdistributionline_Xpath']
        #切换到结果frame
        self.DistributionLineCalcResult_frame=self.Data['DistributionLineCalcResult_frame']
        #点击导出
        self.Daochu_id=self.Data['Daochu_id']

        '''#获取相关的信息，title和内容信息以前获取，参数变化从1082到1108，range(1082,1109),通过滚动滚动条获取相关信息
        self.BeginNum = self.Data['BeginNum']
        self.EndNum = self.Data['EndNum']
        #title的ID，分为前后两部分，中间有数字参数
        self.TargetResult_id_Before = self.Data['TargetResult_id_Before']
        self.TargetResult_id_After = self.Data['TargetResult_id_After']
        #页面元素的class,结尾有一组数字，定位到哪一个元素
        self.TargetResult_Class_Before = self.Data['TargetResult_Class_Before']
        #定义一个字典，该字典获取到最后的结果值，格式为Title：Result
        self.CalResult = []
        self.Result={}'''

    def DistributionlineEleQuamoncalcResult_Fun(self):
        '''620kv月电量法计算结果验证'''
        #点击理论线损
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #点击配电线路计算结果查询xpath
        self.driver.find_element_by_xpath(self.TheoreticalLineLossManualcal_Xpath).click()
        time.sleep(2)
        #点击配线月报表
        self.driver.find_element_by_xpath(self.DistributionLineMonthlyReport_Xpath).click()
        time.sleep(2)
        #点击配电线路线损表
        self.driver.find_element_by_xpath(self.Linelossmeterofdistributionline_Xpath).click()
        time.sleep(2)
        #切换到结果frame
        self.driver.switch_to_frame(self.DistributionLineCalcResult_frame)
        self.driver.implicitly_wait(30)
        #点击导出按钮
        self.driver.find_element_by_id(self.Daochu_id).click()
        time.sleep(10)
        self.data_path= u'C:\\Users\\'+GetUsername()+ u'\\Downloads\\配电线路损耗报表（月）.xls'
        self.sheetname= u'配电线路损耗报表（月）'
        self.Excel=ExcelData(self.data_path,self.sheetname)
        self.ExcelDataa=self.Excel.readExcel()
        #return json.dumps(self.ExcelDataa, encoding="UTF-8", ensure_ascii=False)
        return self.ExcelDataa
        '''for i in range(self.BeginNum,self.EndNum):
            self.TargetResult_id = self.TargetResult_id_Before+str(i)+self.TargetResult_id_After
            self.TargetResult_Class = self.TargetResult_Class_Before+str(i)
            #print self.TargetResult_id,self.TargetResult_Class
            self.target = self.driver.find_element_by_id(self.TargetResult_id)
            self.driver.execute_script("arguments[0].scrollIntoView();", self.target) #拖动到可见的元素去
            Title = self.driver.find_element_by_id(self.TargetResult_id).text
            PostmortemPara = self.driver.find_element_by_class_name(self.TargetResult_Class).text
            #print Title,   PostmortemPara
            self.Result[Title] = PostmortemPara
            time.sleep(0.3)
        self.CalResult.append(self.Result)
        #return ("%0.2f" % self.CalResult)
        #return self.CalResult
        print json.dumps(self.CalResult, encoding="UTF-8", ensure_ascii=False)'''
if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lyltest','000000')
    Try1=DistributionlineEleQuamoncalcResult(driver)
    Try1.DistributionlineEleQuamoncalcResult_Fun()
    time.sleep(10)
    driver.quit()

