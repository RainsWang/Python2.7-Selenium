#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/10/31 16:26
当前项目名称  ：AutoTest
功能          ：
-------------------------------------漂亮的分割线----------------------------------------------'''
#!/usr/bin/env python
# -*- coding: utf-8 -*-
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
import yaml
import time
import Login
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class NewtonMethodCalc:
    #输电网手工计算页面牛顿拉夫逊算法计算
    def __init__(self,driver):
        self.driver=driver
        self.ProjectFilePath=GetProjectFilePath()
        #获取页面元素
        self.Page_object_data_file=open(self.ProjectFilePath+"\\Page_object\\Data\\NewtonMethodCalc.yaml")
        self.PageData=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        self.Data=self.PageData['NewtonMethodCalc']

        #理论线损ID
        self.LineLoss_ID = self.Data['LineLoss_ID']
        #理论线损手工计算绝对路径
        self.LineLossManualCalc = self.Data['LineLossManualCalc']
        #计算窗口iframe的ID
        self.CalcWindowIframe_ID = self.Data['CalcWindowIframe_ID']
        #选择单位绝对路径
        self.Company_Xpath = self.Data['Company_Xpath']
        #计算按钮CSS路径
        self.CalcButton_CSS = self.Data['CalcButton_CSS']
        #计算完成提示信息ID
        self.CalcCompletion_ID = self.Data['CalcCompletion_ID']
    def NewtonMethodCalc_Fun(self):
        #点击理论线损
        self.driver.find_element_by_id(self.LineLoss_ID).click()
        time.sleep(2)
        #点击理论线损手工计算
        self.driver.find_element_by_xpath(self.LineLossManualCalc).click()
        time.sleep(5)
        #切换iframe
        self.driver.switch_to_frame(self.CalcWindowIframe_ID)
        #默认进入输电网理论线损计算页面，选择单位
        self.driver.find_element_by_xpath(self.Company_Xpath).click()
        time.sleep(5)
        #点击计算
        self.driver.find_element_by_css_selector(self.CalcButton_CSS).click()
        self.driver.implicitly_wait(90)
        time.sleep(3)
        alerttext=self.driver.find_element_by_id(self.CalcCompletion_ID).text
        time.sleep(1)
        #print alerttext
        return alerttext
if __name__ == '__main__':
    driver = webdriver.Chrome()
    Try = Login.Login(driver)
    Try.Login_Function('wjtest','000000')
    Try1=NewtonMethodCalc(driver)
    Try1.NewtonMethodCalc_Fun()
    time.sleep(10)
    driver.quit()