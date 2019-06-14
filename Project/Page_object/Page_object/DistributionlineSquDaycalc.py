#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018-9-28 15:08
当前项目名称  ：Auto
功能          ：配线日均方根电流法计算
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import yaml,re
import time
import Login
from Public_Theory.GetProjectFilePath import GetProjectFilePath
import sys
from selenium.webdriver.common.keys import Keys
reload(sys)
sys.setdefaultencoding('utf-8')
class DistributionlineSquDaycalc:
    def __init__(self,driver):
        '''配线日均方根电流法计算'''
        #self.driver=webdriver.Chrome()
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\DistributionlineSquDaycalc.yaml")
        self.Page_Data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.Page_Data['DistributionlineSquDaycalc']

        #点击理论线损
        self.TheoreticalLineLoss_Xpath = self.Data['TheoreticalLineLoss_Xpath']
        #点击计算任务设置
        self.CalTaskSet_Xpath=self.Data['CalTaskSet_Xpath']
        #点击任务执行情况查询
        self.TaskExecutionQuery_Xpath=self.Data['TaskExecutionQuery_Xpath']
        ##切换到任务设置frame
        self.TaskExecutionQuery_Frame=self.Data['TaskExecutionQuery_Frame']
        #点击所有任务前面的加号
        self.TaskAction_Xpath=self.Data['TaskAction_Xpath']
         #选中配线日均方根电流法任务
        self.DistributionlineRtMeQuaCutMethodTask_Xpath=self.Data['DistributionlineRtMeQuaCutMethodTask_Xpath']
        #点击明细数据
        self.DetailDate_ID=self.Data['DetailDate_ID']
        #切换到日均方根电流法任务执行情况明细frame
        self.DistributionlineRtMeQuaCutMethodTaskDetail_Frame=self.Data['DistributionlineRtMeQuaCutMethodTaskDetail_Frame']
        #点击日期输入框，输入2014-02-01
        self.Datatimeinput_id = self.Data['Datatimeinput_id']
        # #点击日期选择框
        # self.DataSelectBox_Xpath=self.Data['DataSelectBox_Xpath']
        # #选择2014-8-1日
        # #点击日期选择框
        # self.DataSelect_ID=self.Data['DataSelect_ID']
        # #选择2014年
        # self.YearSelect_Xpath=self.Data['YearSelect_Xpath']
        # #选择8月
        # self.MonthSelect_Xpath=self.Data['MonthSelect_Xpath']
        # #点击确定
        # self.Determine_ID=self.Data['Determine_ID']
        # #选择1号
        # self.NumberOne_Xpath=self.Data['NumberOne_Xpath']
        #点击查询
        self.Query_ID=self.Data['Query_ID']
        #选择任务
        self.TaskSelect_Xpath=self.Data['TaskSelect_Xpath']
        #点击补算
        self.Supplement_ID=self.Data['Supplement_ID']
        #切换到补算窗口frame
        #切换弹出框中补算
        self.SupplementButton_ID = self.Data['SupplementButton_ID']
        #计算完成之后验证
        self.AlertConfirm_ID=self.Data['AlertConfirm_ID']


    def DistributionlineSquDaycalc_Fun(self):
        '''配线日均方根电流法计算'''
        #点击理论线损
        self.driver.find_element_by_xpath(self.TheoreticalLineLoss_Xpath ).click()
        time.sleep(2)
         #点击计算任务设置
        self.driver.find_element_by_xpath(self.CalTaskSet_Xpath).click()
        time.sleep(2)
        #点击任务执行情况查询
        self.driver.find_element_by_xpath(self.TaskExecutionQuery_Xpath).click()
        time.sleep(2)
        #切换到任务设置frame
        self.driver.switch_to_frame(self.TaskExecutionQuery_Frame)
        self.driver.implicitly_wait(30)
        #点击所有任务前面的加号
        self.driver.find_element_by_xpath(self.TaskAction_Xpath).click()
        time.sleep(2)
        #选中配线日均方根电流法任务
        self.driver.find_element_by_xpath(self.DistributionlineRtMeQuaCutMethodTask_Xpath).click()
        time.sleep(2)
        #点击明细数据
        self.driver.find_element_by_id(self.DetailDate_ID).click()
        time.sleep(2)
        #跳出当面frame
        self.driver.switch_to.parent_frame()## 跳出当前iframe
        self.driver.implicitly_wait(10)
        #切换到日均方根电流法任务执行情况明细frame
        self.driver.switch_to_frame(self.DistributionlineRtMeQuaCutMethodTaskDetail_Frame)
        self.driver.implicitly_wait(30)
        #点击日期选择框
        self.driver.find_element_by_id(self.Datatimeinput_id).click()
        self.driver.find_element_by_id(self.Datatimeinput_id).send_keys(Keys.CONTROL,'a')
        self.driver.find_element_by_id(self.Datatimeinput_id).send_keys('2014-08-01')

        # self.driver.find_element_by_xpath(self.DataSelectBox_Xpath).click()
        # time.sleep(2)
        # # #选择2014-8-1日
        # #点击日期选择框
        # self.driver.find_element_by_id(self.DataSelect_ID).click()
        # time.sleep(2)
        # ##选择2014年
        #
        # self.driver.find_element_by_xpath(self.YearSelect_Xpath).click()
        # time.sleep(2)
        # ##选择8月
        # self.driver.find_element_by_xpath(self.MonthSelect_Xpath).click()
        # time.sleep(2)
        # ##点击确定
        # self.driver.find_element_by_id(self.Determine_ID).click()
        # time.sleep(2)
        # #选择1号
        # self.driver.find_element_by_xpath(self.NumberOne_Xpath).click()
        # time.sleep(2)
        #点击查询
        self.driver.find_element_by_id(self.Query_ID).click()
        time.sleep(2)
        #选择任务
        self.driver.find_element_by_xpath(self.TaskSelect_Xpath).click()
        time.sleep(2)
        #点击补算
        self.driver.find_element_by_id(self.Supplement_ID).click()
        time.sleep(2)
        #切换到补算窗口frame
        self.DownLoadPage = self.driver.page_source
        self.FrameIDDynamicPart = re.findall("iframe id=\"eagleWindow(.*?)frame",self.DownLoadPage)
        #print self.FrameIDDynamicPart
        self.SupplementFrame_id = "eagleWindow"+self.FrameIDDynamicPart[0]+"frame"
        self.driver.switch_to.frame(self.SupplementFrame_id)
        time.sleep(2)
        #切换弹出框中补算
        self.driver.find_element_by_id(self.SupplementButton_ID).click()
        time.sleep(30)
        #计算完成之后验证
        alerttext=self.driver.find_element_by_id(self.AlertConfirm_ID).text
        #建议添加隐式等待
        # print alerttext
        # if alerttext == u'补算完成':
        #     print "计算成功！"
        # else:
        #     print "无法正常计算"
        # time.sleep(2)
        return alerttext
if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lyltest','000000')
    Try1=DistributionlineSquDaycalc(driver)
    Try1.DistributionlineSquDaycalc_Fun()
    time.sleep(10)
    driver.quit()