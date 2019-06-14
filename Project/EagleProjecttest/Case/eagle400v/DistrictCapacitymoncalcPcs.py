#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：王敏
IDE           ：PyCharm
时间          ：2019/1/4 13:36
当前项目名称  ：Project
功能          ：400V容量法月计算
-------------------------------------漂亮的分割线----------------------------------------------'''
import yaml
import time
from Public_Theory import Log, Deletefiles, GetProjectFilePath, ExcelToDict, GetCapture
from Page_object.Page_object.Login import Login
from Page_object.Page_object.DistrictmoncalcPcsEnter import DistrictmoncalcPcsEnter
from Page_object.Page_object.DistrictCapacitymoncalc import DistrictCapacitymoncalc
from Page_object.Page_object.DistrictCapacitymoncalcResult import DistrictCapacitymoncalcResult
from Page_object.Page_object.DistrictEleQuamonWirecalcResult import DistrictEleQuamonWirecalcResult
from Page_object.Page_object.DistrictEleQuamonMetercalResult import DistrictEleQuamonMetercalResult
from Page_object.Page_object import BrowserEngine
import unittest
class DistrictCapacitymoncalcPcs(unittest.TestCase):
    '''400V理论手工计算容量法月计算流程：'''
    def setUp(self):
        Deletefiles.del_file()
        self.ProjectFilePath = GetProjectFilePath.GetProjectFilePath()
        self.filepath1 = self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle400v\\低压台区损耗表报表（容）.xls'
        self.filepath2 = self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle400v\\低压台区导线损耗报表（容）.xls'
        self.filepath3 = self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle400v\\低压台区电表损耗报表（容）.xls'

        #print self.filepath1
        #初始化日志对象
        self.ReportLog = Log.LogOutput('理论线损数据输入以及计算')
        #self.driver=webdriver.Chrome()
        self.browser = BrowserEngine.BrowserEngine()
        self.driver = self.browser.GetDriver()
        self.data_file = open(self.ProjectFilePath+'\\EagleProjecttest\\Data\\eagle400v\\DistrictEleQuamonEnterdata.yaml')
        self.data = yaml.load(self.data_file)
        self.data_file.close()
        self.data1 = self.data['DistrictEleQuamonEnterdata']
        self.LoginMsg = self.data1['Login']
        self.DataEnter = self.data1['DistrictDataEntry']

    def test_A1_DistrictCapacitymoncalc(self):
        '''台区月容量法损耗计算过程'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if u'wm县' == ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            time.sleep(5)
            self.RunFun = DistrictCapacitymoncalc(self.driver)
            self.ReturnMsg2 = self.RunFun.DistrictCapacitymoncalc_Fun()
            time.sleep(5)
            if u'是否要打开当前台区的计算结果？'== self.ReturnMsg2:
                self.ReportLog.info_log("计算成功")
        except Exception, e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("计算失败，请检查")
            raise e
    def test_A2_DistrictCapacitymoncalcResult(self):
        '''台区月容量法损耗计算结果验证'''
        try:
            Loginrun = Login(self.driver)
            self.ReturnMsg= Loginrun.Login_Function(self.LoginMsg['username'], self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if u'wm县' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            self.data_path = self.filepath1
            self.sheetname = u'月低压台区损耗报表（月）'
            self.Excel = ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatab = self.Excel.readExcel()
            #print self.ExcelDatab
            self.WiringCalResultrun =DistrictCapacitymoncalcResult(self.driver)
            self.ReturnMsg3=self.WiringCalResultrun.DistrictCapacitymoncalcResult_Fun()
            time.sleep(15)
            #print self.ReturnMsg3
            self.assertEqual(0, cmp(self.ReturnMsg3, self.ExcelDatab))
            self.ReportLog.info_log("返回的计算结果%s"%self.ExcelDatab)
        except Exception, e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def test_A3_DistrictCapacitymonWirecalcResult(self):
        '''台区月容量法导线损耗计算结果验证'''
        try:
            Loginrun = Login(self.driver)
            self.ReturnMsg = Loginrun.Login_Function(self.LoginMsg['username'], self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if u'wm县' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            self.data_path = self.filepath2
            self.sheetname = u'低压台区导线损耗报表（月）'
            self.Excel = ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatab2 = self.Excel.readExcel()
            self.WirecalcResulttrun = DistrictEleQuamonWirecalcResult(self.driver)
            self.ReturnMsg4=self.WirecalcResulttrun.DistrictEleQuamonWirecalcResult_Fun()
            time.sleep(15)
            self.assertEqual(0, cmp(self.ReturnMsg4,self.ExcelDatab2))
            self.ReportLog.info_log("返回的计算结果%s"%self.ExcelDatab2)
        except Exception, e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def test_A4_DistrictCapacitymonMetercalResult(self):
        '''台区月容量法电表损耗计算结果验证'''
        try:
            Loginrun = Login(self.driver)
            self.ReturnMsg = Loginrun.Login_Function(self.LoginMsg['username'], self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log(
                "登录用户名:%s 密码:%s返回结果:%s" % (self.LoginMsg['username'], self.LoginMsg['password'], self.ReturnMsg))
            if u'wm县' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            self.data_path = self.filepath3
            self.sheetname = u'低压台区电表损耗报表（月）'
            self.Excel = ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatab3 = self.Excel.readExcel()
            self.metercalcResulttrun = DistrictEleQuamonMetercalResult(self.driver)
            self.ReturnMsg5 = self.metercalcResulttrun.DistrictEleQuamonMetercalResult_Fun()
            time.sleep(15)
            self.assertEqual(0, cmp(self.ReturnMsg5, self.ExcelDatab3))
            self.ReportLog.info_log("返回的计算结果%s" %self.ExcelDatab3)
        except Exception, e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e

    def tearDown(self):
        time.sleep(10)
        self.driver.quit()
if __name__ == '__main__':
    unittest.main()