# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/10/23 0023 下午 3:26
当前项目名称  ：Auto
功能          ：理论线损容量法日计算变压器损耗结果验证
-------------------------------------漂亮的分割线----------------------------------------------
'''


import yaml,time
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from Public_Theory.ExcelToDict import ExcelData
from Public_Theory.GetUsername import GetUsername
from Page_object.Page_object.Login import Login
class DistributionlineVoluDayCalcTransLoss():
    def __init__(self,driver):
        #self.driver = webdriver.Chrome()
        self.driver = driver
        self.datafile = GetProjectFilePath() + u"\Page_object\Data\DistributionlineVoluDayCalcTransLoss.yaml"
        with open(self.datafile) as datafile:
            self.data = yaml.load(datafile)
        self.datamsg = self.data['DistributionlineVoluDayCalcTransLoss']

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
        #点击变压器损耗的连接
        self.TransLossLink_Path = self.datamsg["TransLossLink_Path"]
        #切换到变压器损耗的frame中
        self.TransLossFrame_id = self.datamsg["TransLossFrame_id"]
        # 导出按钮
        self.Export_id = self.datamsg["Export_id"]
    def DistributionlineVoluDayCalcTransLoss_Fun(self):
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
        self.driver.implicitly_wait(30)
        #点击变压器损耗的连接
        self.driver.find_element_by_xpath(self.TransLossLink_Path).click()
        time.sleep(1)
        #切换到变压器损耗的frame中
        self.driver.switch_to.parent_frame()
        time.sleep(1)
        self.driver.switch_to.frame(self.TransLossFrame_id)
        time.sleep(1)
        # 点击左下方的导出按钮
        self.driver.find_element_by_id(self.Export_id).click()
        time.sleep(10)
        #读取导出文件execl内容，生成列表返回
        self.ExportExecl = ExcelData(u"C:\\Users\\"+GetUsername()+u"\\Downloads\\配电变压器损耗表报表（日）.xls", u"配电变压器损耗表报表（日）")
        return self.ExportExecl.readExcel()

if __name__ == "__main__":
    try:
        driver = webdriver.Chrome()
        LoginRun = Login(driver)
        LoginRun.Login_Function("Autotest", "000000")
        Try = DistributionlineVoluDayCalcTransLoss(driver)
        result = Try.DistributionlineVoluDayCalcTransLoss_Fun()
        print result

    finally:
        driver.quit()