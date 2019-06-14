# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/9/27 0027 下午 4:28
当前项目名称  ：620计算
功能          ：月容量法理论线损计算界面封装
-------------------------------------漂亮的分割线----------------------------------------------
'''
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import yaml,Login,time
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath

class DistributionlineVoluMonthCalc:
    '''620kv月容量法计算页面封装'''
    def __init__(self,driver):
        #self.driver = webdriver.Chrome()
        self.driver = driver
        self.datafile = GetProjectFilePath() +u"\Page_object\Data\DistributionlineVoluMonthCalc.yaml"
        with open(self.datafile) as datafile:
            self.data = yaml.load(datafile)
        self.datamsg = self.data['DistributionlineVoluMonthCalc']

        #树结构中理论线损的ID
        self.DistributionlineLossTree_id = self.datamsg["DistributionlineLossTree_id"]
        # 树结构中理论线损手工计算的xpath定位路径
        self.DistributionlineLossManualCal_xpath = self.datamsg["DistributionlineLossManualCal_xpath"]
        # 理论线损手工计算frame信息
        self.TheoreticalLineloss620kV_Frame_id = self.datamsg["TheoreticalLineloss620kV_Frame_id"]
        # 配电网6-20kV理论线损计算tab路径信息
        self.TheoreticalLineloss620kVTab_css = self.datamsg["TheoreticalLineloss620kVTab_css"]
        # 选择计算线路定位信息
        self.TheroeticalLineloss620kVAeraLine_xpath = self.datamsg["TheroeticalLineloss620kVAeraLine_xpath"]
        # 计算方法选择为容量法定位信息
        self.TheroeticalLineLossCalMethod_xpath =self.datamsg["TheroeticalLineLossCalMethod_xpath"]
        # 计算按钮定位新
        self.Calculate_css = self.datamsg["Calculate_css"]
        # 计算成功后弹出框信息ID以及取消按钮ID
        self.AlertMsg_ID =self.datamsg["AlertMsg_ID"]
        self.AlertCancel_ID=self.datamsg["AlertCancel_ID"]


    def DistributionlineVoluMonthCalc_Fun(self):
        '''访问配电线路620kv计算界面，选择计算线路以及计算方法并执行计算，如果成功返回True，失败False'''
        #点击树结构中的理论计算
        self.driver.find_element_by_id(self.DistributionlineLossTree_id).click()
        #点击树结构中的理论线损手工计算
        time.sleep(1)
        self.driver.find_element_by_xpath(self.DistributionlineLossManualCal_xpath).click()
        #切换到计算界面的frame
        time.sleep(1)
        self.driver.switch_to.frame(self.TheoreticalLineloss620kV_Frame_id)
        #点击配电线路620kv计算的tab
        time.sleep(1)
        self.driver.find_element_by_css_selector(self.TheoreticalLineloss620kVTab_css).click()
        #选择计算的线路
        time.sleep(1)
        self.driver.find_element_by_xpath(self.TheroeticalLineloss620kVAeraLine_xpath).click()
        #设置计算的方法为容量法
        time.sleep(1)
        self.driver.find_element_by_xpath(self.TheroeticalLineLossCalMethod_xpath).click()
        #点击计算按钮
        time.sleep(1)
        self.driver.find_element_by_css_selector(self.Calculate_css).click()
        self.driver.implicitly_wait(120)
        #获取到计算结束后的弹出信息
        self.alerttext = self.driver.find_element_by_id(self.AlertMsg_ID).text
        #取消跳转
        time.sleep(1)
        self.driver.find_element_by_id(self.AlertCancel_ID).click()
        if u'是否要打开当前配线的计算结果？' == self.alerttext:
            return True
        else:
            return False
if __name__ == '__main__':
    driver = webdriver.Chrome()
    test = Login.Login(driver)
    test.Login_Function('Autotest','000000')
    test1 = DistributionlineVoluMonthCalc(driver)
    test1.DistributionlineVoluMonthCalc_Fun()