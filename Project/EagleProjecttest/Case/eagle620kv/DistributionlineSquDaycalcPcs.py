#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018-9-28 16:19
当前项目名称  ：Auto
功能          ：配线日均方根电流法整体计算结果验证
-------------------------------------漂亮的分割线----------------------------------------------'''
import yaml
import time
from Public_Theory import Log, Deletefiles, GetProjectFilePath, ExcelToDict, GetCapture
from Page_object.Page_object.Login import Login
from Page_object.Page_object.DistributionlineMeasDayEnter import DistributionlineMeasDayEnter
from Page_object.Page_object.DistTransformerEleQuaDayCalResult import DistTransformerEleQuaDayCalResult
import unittest
import sys
from Page_object.Page_object import BrowserEngine
from Page_object.Page_object.DistributionlineSquDaycalc import DistributionlineSquDaycalc
from Page_object.Page_object.DistributionlineSquDaycalcResult import DistributionlineSquDaycalcResult
from Page_object.Page_object.DistAclinesegeSquDaycalcResult import DistAclinesegeSquDaycalcResult
reload(sys)
sys.setdefaultencoding('utf-8')


class DistributionlineSquDaycalcPcs(unittest.TestCase):
    '''配线运行数据录入以及结果核对'''
    def setUp(self):
        #path=os.path.join('C:/Users/',getpass.getuser(),'Downloads')
        Deletefiles.del_file()
        self.ProjectFilePath= GetProjectFilePath.GetProjectFilePath()
        self.filepath1=self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle620kv\\配电线路均方根电流法损耗报表（日）.xls'
        self.filepath2=self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle620kv\\配电线路均方根电流法导线损耗报表（日）.xls'
        self.filepath3=self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle620kv\\配电线路均方根电流法变压器损耗表报表（日）.xls'
        #print self.filepath1
        #初始化日志对象
        self.ReportLog= Log.LogOutput('理论线损数据输入以及计算')
        #self.driver=webdriver.Chrome()
        self.browser = BrowserEngine.BrowserEngine()
        self.driver = self.browser.GetDriver()
        self.data_file=open(self.ProjectFilePath+'\\EagleProjecttest\\Data\\eagle620kv\\DistributionlineMeasDayEnterAndCalc.yaml')
        self.data=yaml.load(self.data_file)
        self.data_file.close()
        self.data1=self.data['DistributionlineMeasDayEnterAndCalc']
        self.LoginMsg=self.data1['Login']
        self.DataEnter=self.data1['DistributionlineDataEntry']
    def test_A1_DistributionlineSquMeasDay(self):
        '''单线图日均方根电流法理论运行数据录入'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lyltest' == ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #访问数据录入界面，进行日运行数据的录入
            DataEnterrun=DistributionlineMeasDayEnter(self.driver)
            self.ReturnMsg1=DataEnterrun.DistributionlineMeasDayEnter_Fun(self.DataEnter['DistributionlineDayData'].split(","),self.DataEnter['DistributionTransformDayData'].split(","))
            time.sleep(5)
            if '没有修改的数据要保存！' == self.ReturnMsg1:
                self.ReportLog.info_log("未修改输入数据，请确认数据是否为正确的数据")
            else:
                self.ReportLog.info_log("录入数据成功")
            time.sleep(2)
            #访问数据录入界面，进行小时运行数据的录入
            self.driver.quit()
            self.browser = BrowserEngine.BrowserEngine()
            self.driver = self.browser.GetDriver()
            Loginrun=Login(self.driver)
            ReturnMsg2=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lyltest' == ReturnMsg2:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            DataEnterrun2=DistributionlineMeasDayEnter(self.driver)
            self.ReturnMsg2=DataEnterrun2.DistributionlineMeasHourEnter_Fun(self.DataEnter['DistributionlineVoltHourData'].split(","),self.DataEnter['DistributionlineElcHourData'].split(","),self.DataEnter['DistributionTransformActivePowerHourData'].split(","),self.DataEnter['DistributionTransformReactivePowerHourData'].split(","),self.DataEnter['DistributionTransformVoltHourData'].split(","))
            time.sleep(5)
            if '没有修改的数据要保存！' == self.ReturnMsg2:
                self.ReportLog.info_log("未修改输入数据，请确认数据是否为正确的数据")
            else:
                self.ReportLog.info_log("录入数据成功")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("数据无法正确保存")
            raise e
    def test_A2_DistributionlineSquDaycalc(self):
        '''6-20kv日均方根电流法计算'''
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
            self.DistributionlineSquDaycalc=DistributionlineSquDaycalc(self.driver)
            self.ReturnMsg2=self.DistributionlineSquDaycalc.DistributionlineSquDaycalc_Fun()
            time.sleep(5)
            if u'补算完成'==self.ReturnMsg2:
                self.ReportLog.info_log("计算成功")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("计算失败，请检查")
            raise e
    def test_A3_DistributionlineEleQuadaycalcResult(self):
        '''单线图日均方根电流法整体计算结果验证'''
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
            self.sheetname= u'配电线路损耗报表（日）'
            self.Excel= ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatab=self.Excel.readExcel()
            #print self.ExcelDatab
            self.DistributionlineSquDaycalcResult =DistributionlineSquDaycalcResult(self.driver)
            self.ReturnMsg3=self.DistributionlineSquDaycalcResult.DistributionlineSquDaycalcResult_Fun()
            time.sleep(5)
            #print self.ReturnMsg3
            self.assertEqual(0,cmp(self.ReturnMsg3,self.ExcelDatab))
            self.ReportLog.info_log("返回的计算结果%s"%self.ExcelDatab)
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def test_A4_DistAclinesegeSquDaycalcResult(self):
        '''单线图日电量法导线损耗明细结果验证'''
        try:
            #清空文件导出路径下的文件
            Loginrun=Login(self.driver)
            self.ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'lyltest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            self.data_path=self.filepath2
            self.sheetname= u'配电线路导线损耗报表（日）'
            self.Excel= ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatac=self.Excel.readExcel()
            #print self.ExcelDatab
            self.DistAclinesegeSquDaycalcResult=DistAclinesegeSquDaycalcResult(self.driver)
            self.ReturnMsg4=self.DistAclinesegeSquDaycalcResult.DistAclinesegeSquDaycalcResult_Fun()
            time.sleep(5)
            self.assertEqual(0,cmp(self.ReturnMsg4,self.ExcelDatac))
            self.ReportLog.info_log("返回的计算结果%s"%self.ReturnMsg4)
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def test_A5_DistTransformerSquDaycalcResult(self):
        '''单线图日电量法变压器损耗明细结果验证'''
        try:
            #清空文件导出路径下的文件
            Loginrun=Login(self.driver)
            self.ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(5)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'lyltest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            self.data_path=self.filepath3
            self.sheetname= u'配电变压器损耗表报表（日）'
            self.Excel= ExcelToDict.ExcelData(self.data_path, self.sheetname)
            self.ExcelDatad=self.Excel.readExcel()
            self.Transformerlossmeterresultrun=DistTransformerEleQuaDayCalResult(self.driver)
            self.ReturnMsg5=self.Transformerlossmeterresultrun.DistTransformerEleQuaDayCalResult_Fun()
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
