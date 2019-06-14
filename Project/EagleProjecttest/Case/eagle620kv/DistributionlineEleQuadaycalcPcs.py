#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018-8-28 16:45
当前项目名称  ：Auto
功能          ：登录系统，访问数据中心配线理论线损运行数据录入界面输入日运行数据信息
访问理论线损手动计算界面计算结果，点击导线、变压器损耗链接，并跳转到相关界面核对结果信息
-------------------------------------漂亮的分割线----------------------------------------------'''
import yaml
import time
from Page_object.Page_object.Login import Login
from Page_object.Page_object.DistributionlineMeasDayEnter import DistributionlineMeasDayEnter
import unittest
import sys
from Page_object.Page_object import BrowserEngine
from Page_object.Page_object.DistributionlineEleQuadaycalc import DistributionlineEleQuadaycalc
from Page_object.Page_object.DistributionlineEleQuadaycalcResult import DistributionlineEleQuadaycalcResult
from Page_object.Page_object.DistAclinesegeEleQuadayCalResult import DistAclinesegeEleQuadayCalResult
from Page_object.Page_object.DistTransformerEleQuaDayCalResult import DistTransformerEleQuaDayCalResult
from Public_Theory import Log, Deletefiles, GetProjectFilePath, ExcelToDict, GetCapture
reload(sys)
sys.setdefaultencoding('utf-8')


class DistributionlineEleQuadaycalcPcs(unittest.TestCase):
    '''配线运行数据录入以及结果核对'''
    def setUp(self):
        #path=os.path.join('C:/Users/',getpass.getuser(),'Downloads')
        Deletefiles.del_file()
        self.ProjectFilePath= GetProjectFilePath.GetProjectFilePath()
        self.filepath1=self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle620kv\\配电线路损耗报表（日）.xls'
        self.filepath2=self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle620kv\\配电线路导线损耗报表（日）.xls'
        self.filepath3=self.ProjectFilePath+u'\\EagleProjecttest\\Data\\eagle620kv\\配电变压器损耗表报表（日）.xls'
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
    def test_A1_DistributionlineMeasDayEnter(self):
        '''单线图日电量法理论运行数据录入'''
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
            DataEnterrun=DistributionlineMeasDayEnter(self.driver)
            self.ReturnMsg1=DataEnterrun.DistributionlineMeasDayEnter_Fun(self.DataEnter['DistributionlineDayData'].split(","),self.DataEnter['DistributionTransformDayData'].split(","))
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
    def test_A2_DistributionlineEleQuadaycalc(self):
        '''6-20kv日电量法计算'''
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
            self.DistributionlineEleQuadaycalc=DistributionlineEleQuadaycalc(self.driver)
            self.ReturnMsg2=self.DistributionlineEleQuadaycalc.DistributionlineEleQuadaycalc_Fun()
            time.sleep(5)
            if u'是否要打开当前配线的计算结果？'==self.ReturnMsg2:
                self.ReportLog.info_log("计算成功")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("计算失败，请检查")
            raise e
    def test_A3_DistributionlineEleQuadaycalcResult(self):
        '''单线图日电量法整体计算结果验证'''
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
            self.Excel=ExcelToDict.ExcelData(self.data_path,self.sheetname)
            self.ExcelDatab=self.Excel.readExcel()
            self.DistributionlineEleQuadaycalcResult =DistributionlineEleQuadaycalcResult(self.driver)
            self.ReturnMsg3=self.DistributionlineEleQuadaycalcResult.DistributionlineEleQuadaycalcResult_Fun()
            time.sleep(5)
            #print self.ReturnMsg3
            self.assertEqual(0,cmp(self.ReturnMsg3,self.ExcelDatab))
            self.ReportLog.info_log("返回的计算结果%s"%self.ExcelDatab)
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def test_A4_DistAclinesegeEleQuaDayCalResult(self):
        '''单线图日电量法导线损耗明细结果验证'''
        try:
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
            self.Excel=ExcelToDict.ExcelData(self.data_path,self.sheetname)
            self.ExcelDatac=self.Excel.readExcel()
            self.DistAclinesegeEleQuadayCalResult=DistAclinesegeEleQuadayCalResult(self.driver)
            self.ReturnMsg4=self.DistAclinesegeEleQuadayCalResult.DistAclinesegeEleQuadayCalResult_Fun()
            time.sleep(5)
            self.assertEqual(0,cmp(self.ReturnMsg4,self.ExcelDatac))
            self.ReportLog.info_log("返回的计算结果%s"%self.ReturnMsg4)
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("返回的结果不正确")
            raise e
    def test_A5_DistTransformerEleQuaDayCalResult(self):
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
            self.Excel=ExcelToDict.ExcelData(self.data_path,self.sheetname)
            self.ExcelDatad=self.Excel.readExcel()
            self.DistTransformerEleQuaDayCalResult=DistTransformerEleQuaDayCalResult(self.driver)
            self.ReturnMsg5=self.DistTransformerEleQuaDayCalResult.DistTransformerEleQuaDayCalResult_Fun()
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
