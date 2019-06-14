#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：王敏
IDE           ：PyCharm Community Edition
时间          ：2018/10/24 11:08
当前项目名称  ：400VCalc
功能          ：登录系统，访问数据中心理论线损运行数据录入界面输入数据信息
访问理论线损手动计算界面计算结果，并跳转到相关界面核对结果信息
-------------------------------------漂亮的分割线----------------------------------------------'''
import yaml
import time
from Public_Theory import Log, Deletefiles, GetProjectFilePath, ExcelToDict, GetCapture
from Page_object.Page_object.Login import Login
from Page_object.Page_object.DistrictmoncalcPcsEnter import DistrictmoncalcPcsEnter
from Page_object.Page_object.DistrictEleQuamoncalc import DistrictEleQuamoncalc
from Page_object.Page_object.DistrictEleQuamoncalcResult import DistrictEleQuamoncalcResult
from Page_object.Page_object.DistrictEleQuamonWirecalcResult import DistrictEleQuamonWirecalcResult
from Page_object.Page_object.DistrictEleQuamonMetercalResult import DistrictEleQuamonMetercalResult
from Page_object.Page_object import BrowserEngine
import unittest

class DistrictEleQuamoncalcPcs(unittest.TestCase):
    '''400V理论计算电量法月计算流程：'''
    def setUp(self):
        Deletefiles.del_file()
        self.ProjectFilePath = GetProjectFilePath.GetProjectFilePath()
        self.filepath1 = self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle400v\\低压台区损耗表报表（月）.xls'
        self.filepath2 = self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle400v\\低压台区导线损耗报表（月）.xls'
        self.filepath3 = self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle400v\\低压台区电表损耗报表（月）.xls'

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

    def test_A1_DistrictmoncalcPcsEnter(self):
        '''台区月理论运行数据录入'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if u'wm县' == ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            # 访问数据录入界面，进行数据的录入
            DataEnterrun = DistrictmoncalcPcsEnter(self.driver)
            self.ReturnMsg1 = DataEnterrun.DistributionlineMeasMonthEnter_Fun(self.DataEnter['ActivePower'],self.DataEnter['PowerFactor'],\
                                                                            self.DataEnter['LoadShapeFactor'],self.DataEnter['DistributionTransformData1'],\
                                                                            self.DataEnter['DistributionTransformData2'],self.DataEnter['DistributionTransformData3'],self.DataEnter['DistributionTransformData4'])
            time.sleep(5)
            #print self.ReturnMsg1
            if '没有修改的数据要保存！' == self.ReturnMsg1:
                self.ReportLog.info_log("未修改输入数据，请确认数据是否为正确的数据")
            else:
                self.ReportLog.info_log("录入数据成功")
        except Exception, e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("数据无法正确保存")
            raise e

    def test_A2_DistrictEleQuamoncalc(self):
        '''台区月电量法理论手工计算'''
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
            self.RunFun = DistrictEleQuamoncalc(self.driver)
            self.ReturnMsg2 = self.RunFun.DistrictEleQuamoncalc_Fun()
            time.sleep(5)
            if u'是否要打开当前台区的计算结果？'== self.ReturnMsg2:
                self.ReportLog.info_log("计算成功")
        except Exception, e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("计算失败，请检查")
            raise e

    def test_A3_DistrictEleQuamoncalcResult(self):
        '''台区月电量法损耗计算结果验证'''
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
            self.sheetname = u'低压台区损耗表报表（月）'
            self.Excel = ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatab = self.Excel.readExcel()
            #print self.ExcelDatab
            self.WiringCalResultrun =DistrictEleQuamoncalcResult(self.driver)
            self.ReturnMsg3=self.WiringCalResultrun.DistrictEleQuamoncalcResult_Fun()
            time.sleep(15)
            #print self.ReturnMsg3
            self.assertEqual(0, cmp(self.ReturnMsg3, self.ExcelDatab))
            self.ReportLog.info_log("返回的计算结果%s"%self.ExcelDatab)
        except Exception, e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def test_A4_DistrictEleQuamonWirecalcResult(self):
        '''台区月电量法导线损耗计算结果验证'''
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
    def test_A5_DistrictEleQuamonMetercalResult(self):
        '''台区月电量法电表损耗计算结果验证'''
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