#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/12 17:56
当前项目名称  ：620Calc
功能          ：登录系统，访问数据中心理论线损运行数据录入界面输入数据信息
访问理论线损手动计算界面计算结果，并跳转到相关界面核对结果信息
-------------------------------------漂亮的分割线----------------------------------------------'''
import yaml
import time
from Public_Theory import Log, Deletefiles, GetProjectFilePath, ExcelToDict, GetCapture
from Page_object.Page_object.Login import Login
from Page_object.Page_object.DistributionlineMeasMonthEnter import DistributionlineMeasMonthEnter
from Page_object.Page_object.DistributionlineEleQuamoncalc import DistributionlineEleQuamoncalc
from Page_object.Page_object.DistributionLineMonthCalcResult import DistributionlineEleQuamoncalcResult
from Page_object.Page_object.DistAclinesegeEleQuaMonCalResult import DistAclinesegeEleQuaMonCalResult
from Page_object.Page_object.DistTransformerEleQuaMonCalResult import DistTransformerEleQuaMonCalResult
import unittest
from Page_object.Page_object import BrowserEngine
class DistributionlineEleQuamoncalcPcs(unittest.TestCase):
    '''配线运行数据录入以及结果核对'''
    def setUp(self):
        Deletefiles.del_file()
        self.ProjectFilePath= GetProjectFilePath.GetProjectFilePath()
        self.filepath1=self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle620kv\\配电线路损耗报表（月）.xls'
        self.filepath2=self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle620kv\\配电线路导线损耗报表（月）.xls'
        self.filepath3=self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle620kv\\配电变压器损耗表报表（月）.xls'
        #print self.filepath1
        #初始化日志对象
        self.ReportLog= Log.LogOutput('理论线损数据输入以及计算')
        #self.driver=webdriver.Chrome()
        self.browser = BrowserEngine.BrowserEngine()
        self.driver = self.browser.GetDriver()
        self.data_file=open(self.ProjectFilePath+'\\EagleProjecttest\\Data\\eagle620kv\\DistributionlineMeasMonthEnterAndCalc.yaml')
        self.data=yaml.load(self.data_file)
        self.data_file.close()
        self.data1=self.data['DistributionlineMeasMonthEnterAndCalc']
        self.LoginMsg=self.data1['Login']
        self.DataEnter=self.data1['DistributionlineDataEntry']
        #self.data1=os.path.abspath(u'配电线路损耗报表（月）.xls')
    def test_A1_DistributionlineMeasMonthEnter(self):
        '''单线图月电量法理论运行数据录入'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lyltest' == ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #访问数据录入界面，进行数据的录入
            DataEnterrun=DistributionlineMeasMonthEnter(self.driver)
            self.ReturnMsg1=DataEnterrun.DistributionlineMeasMonthEnter_Fun(self.DataEnter['ActivePower'],self.DataEnter['PowerFactor'],\
                                                                            self.DataEnter['LoadShapeFactor'],self.DataEnter['DistributionTransformData1'],\
                                                                            self.DataEnter['DistributionTransformData2'],self.DataEnter['DistributionTransformData3'],self.DataEnter['DistributionTransformData4'])
            time.sleep(5)
            #print self.ReturnMsg1
            if '没有修改的数据要保存！' == self.ReturnMsg1:
                self.ReportLog.info_log("未修改输入数据，请确认数据是否为正确的数据")
            else:
                self.ReportLog.info_log("录入数据成功")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("数据无法正确保存")
            raise e
    def test_A2_DistributionlineEleQuamoncalc(self):
        '''6-20kv月电量法计算'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lyltest' == ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            time.sleep(5)
            self.DistributionlineEleQuamoncalc=DistributionlineEleQuamoncalc(self.driver)
            self.ReturnMsg2=self.DistributionlineEleQuamoncalc.DistributionlineEleQuamoncalc_Fun()
            time.sleep(5)
            if u'是否要打开当前配线的计算结果？'==self.ReturnMsg2:
                self.ReportLog.info_log("计算成功")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("计算失败，请检查")
            raise e
    def test_A3_DistributionlineEleQuamoncalcResult(self):
        '''单线图月电量法整体计算结果验证'''
        try:
            Loginrun=Login(self.driver)
            self.ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'lyltest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            self.data_path=self.filepath1
            #self.data_path= u'D:\\620Calc\\Project\\eagle620kv\\EagleProjecttest\\Data\\eagle620kv\\配电线路损耗报表（月）.xls'
            self.sheetname= u'配电线路损耗报表（月）'
            self.Excel= ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatab=self.Excel.readExcel()
            #print self.ExcelDatab
            self.WiringCalResultrun =DistributionlineEleQuamoncalcResult(self.driver)
            self.ReturnMsg3=self.WiringCalResultrun.DistributionlineEleQuamoncalcResult_Fun()
            time.sleep(5)
            #print self.ReturnMsg3
            self.assertEqual(0,cmp(self.ReturnMsg3,self.ExcelDatab))
            self.ReportLog.info_log("返回的计算结果%s"%self.ExcelDatab)
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def test_A4_DistAclinesegeEleQuaMonCalResult(self):
        '''单线图月电量法导线损耗明细结果验证'''
        try:
            Loginrun=Login(self.driver)
            self.ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'lyltest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #self.data_path= u'D:\\620Calc\\Project\\eagle620kv\\EagleProjecttest\\Data\\eagle620kv\\配电线路导线损耗报表（月）.xls'
            self.data_path=self.filepath2
            self.sheetname= u'配电线路导线损耗报表（月）'
            self.Excel= ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatac=self.Excel.readExcel()
            #print self.ExcelDatab
            self.TraverselossmeterCalcResultrun=DistAclinesegeEleQuaMonCalResult(self.driver)
            self.ReturnMsg4=self.TraverselossmeterCalcResultrun.DistAclinesegeEleQuaMonCalResult_Fun()
            time.sleep(5)
            self.assertEqual(0,cmp(self.ReturnMsg4,self.ExcelDatac))
            self.ReportLog.info_log("返回的计算结果%s"%self.ExcelDatac)
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def test_A5_DistTransformerEleQuaMonCalResult(self):
        '''单线图月电量法变压器损耗明细结果验证'''
        try:
            Loginrun=Login(self.driver)
            self.ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'lyltest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #self.data_path= u'D:\\620Calc\\Project\\eagle620kv\\EagleProjecttest\\Data\\eagle620kv\\配电变压器损耗表报表（月）.xls'
            self.data_path=self.filepath3
            self.sheetname= u'配电变压器损耗表报表（月）'
            self.Excel= ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatad=self.Excel.readExcel()
            self.Transformerlossmeterresultrun=DistTransformerEleQuaMonCalResult(self.driver)
            self.ReturnMsg5=self.Transformerlossmeterresultrun.DistTransformerEleQuaMonCalResult_Fun()
            time.sleep(5)
            self.assertEqual(0,cmp(self.ReturnMsg5,self.ExcelDatad))
            self.ReportLog.info_log("返回的计算结果%s"%self.ExcelDatad)
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def tearDown(self):
        time.sleep(10)
        self.driver.quit()

