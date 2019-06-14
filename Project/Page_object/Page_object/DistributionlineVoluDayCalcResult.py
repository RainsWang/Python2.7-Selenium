# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/10/19 0019 上午 11:20
当前项目名称  ：Auto
功能          ：配电线路日报表计算结果导出以及核对
-------------------------------------漂亮的分割线----------------------------------------------
'''

import yaml,time
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
from Public_Theory.GetUsername import GetUsername
from Page_object.Page_object.Login import Login
class DistributionlineVoluDayCalcResult():
    def __init__(self,driver):
        #self.driver = webdriver.Chrome()
        self.driver = driver
        self.datafile = GetProjectFilePath() + u"\Page_object\Data\DistributionlineVoluDayCalcResult.yaml"
        with open(self.datafile) as datafile:
            self.data = yaml.load(datafile)
        self.datamsg = self.data['DistributionlineVoluDayCalcResult']

        # 树结构中理论线损的ID
        self.DistributionlineLossTree_id = self.datamsg["DistributionlineLossTree_id"]
        # 树结构中配电网计算结果查询的xpath
        self.TheoreticalLineLossCalResult_Xpath = self.datamsg["TheoreticalLineLossCalResult_Xpath"]
        # 树结构中日报表的xpath
        self.DistributionLineDaylyReport_Xpath = self.datamsg["DistributionLineDaylyReport_Xpath"]
        # 配电线路线损表
        self.LinelossmeterDistributionline_Xpath = self.datamsg["LinelossmeterDistributionline_Xpath"]
        # 切换到frame
        self.DistributionLineCalcResult_frame = self.datamsg["DistributionLineCalcResult_frame"]
        # 导出按钮
        self.Export_id = self.datamsg["Export_id"]
    def DistributionlineVoluDayCalcResult_Fun(self):
        # 点击理论线损树
        time.sleep(1)
        self.driver.find_element_by_id(self.DistributionlineLossTree_id).click()
        # 点击配电网计算结果查询
        time.sleep(1)
        self.driver.find_element_by_xpath(self.TheoreticalLineLossCalResult_Xpath).click()
        time.sleep(1)
        # 点击月报表
        self.driver.find_element_by_xpath(self.DistributionLineDaylyReport_Xpath).click()
        time.sleep(1)
        # 点击配电线路线损表
        self.driver.find_element_by_xpath(self.LinelossmeterDistributionline_Xpath).click()
        time.sleep(1)
        # 切换frame到配电线路线损表frame
        self.driver.switch_to.frame(self.DistributionLineCalcResult_frame)
        time.sleep(1)
        self.driver.implicitly_wait(30)
        # 点击左下方的导出按钮
        self.driver.find_element_by_id(self.Export_id).click()
        time.sleep(10)
        #读取导出文件execl内容，生成列表返回
        self.ExportExecl = ExcelData(u"C:\\Users\\"+GetUsername()+u"\\Downloads\\配电线路损耗报表（日）.xls", u"配电线路损耗报表（日）")
        return self.ExportExecl.readExcel()

if __name__ == "__main__":
    try:
        driver = webdriver.Chrome()
        LoginRun = Login(driver)
        LoginRun.Login_Function("Autotest", "000000")
        Try = DistributionlineVoluDayCalcResult(driver)
        result = Try.DistributionlineVoluDayCalcResult_Fun()
        print result

    finally:
        driver.quit()