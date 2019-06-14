# -*- coding: utf-8 -*-
#!/usr/bin/env python
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：吴鹏飞
IDE           ：PyCharm Community Edition
时间          ：2018/10/19 15:58
当前项目名称  ：eagle2
功能          ：
-------------------------------------漂亮的分割线---------------------------------------------'''
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
import yaml
import time
import Login
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class WholelossNiuhourcal:
    def __init__(self,driver):
        self.driver=driver
        #获取页面元素
        self.ProjectFilePath=GetProjectFilePath()
        self.ProjectFilePath_file=open(self.ProjectFilePath+"\\Page_object\\Data\\WholelossNiuhourcalc.yaml")
        self.PageData=yaml.load(self.ProjectFilePath_file)
        self.ProjectFilePath_file.close()
        self.Data=self.PageData['WholelossNiuhourcal']

        #理论线损id
        self.TheoreticalLineLoss_ID=self.Data['TheoreticalLineLoss_ID']
        #理论线损手工计算Xpath
        self.TheoreticalLineLossManualcal_Xpath=self.Data['TheoreticalLineLossManualcal_Xpath']
        #手工计算frame
        self.WholelossNiuhourcal_Frame=self.Data['WholelossNiuhourcal_Frame']
        #勾选计算范围id
        self.WholelossNiuhourcalCompany_Xpath=self.Data['WholelossNiuhourcalCompany_Xpath']
        #计算按钮id
        self.Calculate_ID=self.Data['Calculate_ID']
        #计算完成弹出框信息ID
        self.AlertMsg_ID=self.Data['AlertMsg_ID']
        #点击确定提示框
        self.AlertConfirm_ID=self.Data['AlertConfirm_ID']
    def WholelossNiuhourcal_Fun(self):
        #进入理论线损模块
        self.driver.find_element_by_id(self.TheoreticalLineLoss_ID).click()
        time.sleep(2)
        #进入理论线损计算页面
        self.driver.find_element_by_xpath(self.TheoreticalLineLossManualcal_Xpath).click()
        time.sleep(2)
        self.driver.switch_to_frame(self.WholelossNiuhourcal_Frame)
        time.sleep(4)
        #勾选单位
        self.driver.find_element_by_xpath(self.WholelossNiuhourcalCompany_Xpath).click()
        time.sleep(2)
        #点击计算
        self.driver.find_element_by_id(self.Calculate_ID).click()
        self.driver.implicitly_wait(60)
        Alerttext=self.driver.find_element_by_id(self.AlertMsg_ID).text
        time.sleep(2)
        self.driver.find_element_by_id(self.AlertConfirm_ID).click()
        time.sleep(2)
        #print Alerttext
        return Alerttext
if __name__=="__main__":
    driver=webdriver.Chrome()
    Try=Login.Login(driver)
    Try.Login_Function('lj','000000')
    Try1=WholelossNiuhourcal(driver)
    Try1.WholelossNiuhourcal_Fun()
    time.sleep(10)
    driver.quit()
