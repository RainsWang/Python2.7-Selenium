#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/11/1 16:58
当前项目名称  ：AutoTest
功能          ：
-------------------------------------漂亮的分割线----------------------------------------------'''
import yaml
import time
from Public_Theory import Deletefiles,ExcelToDict,GetCapture,GetProjectFilePath,GetUsername,Log
from Page_object.Page_object.Login import Login
from Page_object.Page_object.BrowserEngine import BrowserEngine
from Page_object.Page_object.NewtonMethodDataEnter import NewtonMethodDataEnter
from Page_object.Page_object.NewtonMethodCalc import NewtonMethodCalc
from Page_object.Page_object.NewtonMethodWholelossHourCalcResult import NewtonMethodWholelossHourCalcResult
from Page_object.Page_object.NewtonMethodWholelossDayCalcResult import NewtonMethodWholelossDayCalcResult
from Page_object.Page_object.NewtonMethodWholelossMonCalcResult import NewtonMethodWholelossMonCalcResult
from Page_object.Page_object.NewtonMethodWholelossMonVoltCalcResult import NewtonMethodWholelossMonVoltCalcResult
from Page_object.Page_object.NewtonMethodWholelossDayVoltCalcResult import NewtonMethodWholelossDayVoltCalcResult
from Page_object.Page_object.NewtonMethodWholelossHourVoltCalcResult import NewtonMethodWholelossHourVoltCalcResult
from Page_object.Page_object.NewtonMethodSubstationMonCalcResult import NewtonMethodSubstationMonCalcResult
from Page_object.Page_object.NewtonMethodSubstationDayCalcResult import NewtonMethodSubstationDayCalcResult
from Page_object.Page_object.NewtonMethodSubstationHourCalcResult import NewtonMethodSubstationHourCalcResult
from Page_object.Page_object.NewtonMethodTransformerMonCalcResult import NewtonMethodTransformerMonCalcResult
from Page_object.Page_object.NewtonMethodTransformerDayCalcResult import NewtonMethodTransformerDayCalcResult
from Page_object.Page_object.NewtonMethodTransformerHourCalcResult import NewtonMethodTransformerHourCalcResult
from Page_object.Page_object.NewtonMethodOtherLossMonResult import NewtonMethodOtherLossMonResult
from Page_object.Page_object.NewtonMethodOtherLossDayResult import NewtonMethodOtherLossDayResult
from Page_object.Page_object.NewtonMethodOtherLossHourResult import NewtonMethodOtherLossHourResult
from Page_object.Page_object.NewtonMethodAclineMonCalcResult import NewtonMethodAclineMonCalcResult
from Page_object.Page_object.NewtonMethodAclineDayCalcResult import NewtonMethodAclineDayCalcResult
from Page_object.Page_object.NewtonMethodAclineHourCalcResult import NewtonMethodAclineHourCalcResult
from Page_object.Page_object.NewtonMethodLosspartHourCalcResult import NewtonMethodLosspartHourCalcResult
from Page_object.Page_object.NewtonMethodLosspartVoltHourCalcResult import NewtonMethodLosspartVoltHourCalcResult
from Page_object.Page_object.NewtonMethodNodeVoltPowerHourCalcResult import NewtonMethodNodeVoltPowerHourCalcResult
from Page_object.Page_object.NewtonMethodWholeLossAnalysisMonResult import NewtonMethodWholeLossAnalysisMonResult
from Page_object.Page_object.NewtonMethodWholeLossAnalysisMonAcline import NewtonMethodWholeLossAnalysisMonAcline
from Page_object.Page_object.NewtonMethodWholeLossAnalysisMonTrans import NewtonMethodWholeLossAnalysisMonTrans
from Page_object.Page_object.NewtonMethodWholeLossAnalysisDayResult import NewtonMethodWholeLossAnalysisDayResult
from Page_object.Page_object.NewtonMethodWholeLossAnalysisDayAcline import NewtonMethodWholeLossAnalysisDayAcline
from Page_object.Page_object.NewtonMethodWholeLossAnalysisDayTrans import NewtonMethodWholeLossAnalysisDayTrans
from Page_object.Page_object.NewtonMethodTransAnalysisMonResult import NewtonMethodTransAnalysisMonResult
from Page_object.Page_object.NewtonMethodTransAnalysisMonRetrospect import NewtonMethodTransAnalysisMonRetrospect
from Page_object.Page_object.NewtonMethodTransAnalysisDayResult import NewtonMethodTransAnalysisDayResult
from Page_object.Page_object.NewtonMethodTransAnalysisDayRetrospect import NewtonMethodTransAnalysisDayRetrospect
import unittest
import json
import os
import getpass
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class NewtonMethodCalcPcs(unittest.TestCase):
    '''24点运行数据录入及核对。'''
    def setUp(self):
        #path=os.path.join(u'C:\\Users\\',getpass.getuser(),'Downloads')
        Deletefiles.del_file()
        self.ProjectFilePath=GetProjectFilePath.GetProjectFilePath()
        self.filepath1=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\全网总损耗报表（小时）.xls'
        self.filepath2=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\全网总损耗报表（月）.xls'
        self.filepath3=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\全网总损耗报表（日）.xls'
        self.filepath4=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\全网分电压损耗报表（月）.xls'
        self.filepath5=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\全网分电压损耗报表（日）.xls'
        self.filepath6=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\全网分电压损耗报表（小时）.xls'
        self.filepath7=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\变电站损耗报表（月）.xls'
        self.filepath8=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\变电站损耗报表（日）.xls'
        self.filepath9=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\变电站损耗报表（小时）.xls'
        self.filepath10=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\变压器损耗报表（月）.xls'
        self.filepath11=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\变压器损耗报表（日）.xls'
        self.filepath12=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\变压器损耗报表（小时）.xls'
        self.filepath13=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\导线损耗月报表.xls'
        self.filepath14=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\导线损耗日报表.xls'
        self.filepath15=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\导线损耗报表（小时）.xls'
        self.filepath16=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网其他损耗报表（月）.xls'
        self.filepath17=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网其他损耗报表（日）.xls'
        self.filepath18=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网其他损耗报表（小时）.xls'
        self.filepath19=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\全网分片线损报表（小时）.xls'
        self.filepath20=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\全网分片分压线损报表（小时）.xls'
        self.filepath21=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\节点电压功率报表（小时）.xls'
        self.filepath22=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网综合分析（月）.xls'
        self.filepath23=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网高损线路分析（月）.xls'
        self.filepath24=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网高损变压器分析（月）.xls'
        self.filepath25=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网综合分析（日）.xls'
        self.filepath26=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网高损线路分析(日).xls'
        self.filepath27=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网高损变压器分析（日）.xls'
        self.filepath28=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网变压器经济运行（月）.xls'
        self.filepath29=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网变压器损耗明细（月）-重载变.xls'
        self.filepath30=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网变压器经济运行（日）.xls'
        self.filepath31=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网变压器损耗明细（日）-重载变.xls'
        #日志初始化
        #self.driver=webdriver.Chrome()
        self.ReportLog= Log.LogOutput('牛顿拉夫逊算法主网理论线损数据输入以及计算')
        self.browser=BrowserEngine()
        self.driver = self.browser.GetDriver()
        self.data_file = open(self.ProjectFilePath+'\\EagleProjecttest\\Data\\eagleWhole\\NewtonMethodCalcPcs.yaml')
        self.data = yaml.load(self.data_file)
        self.data_file.close()
        self.data1 = self.data['NewtonMethodCalcPcs']
        self.LoginMsg = self.data1['Login']
        self.DataEnter = self.data1['NewtonMethodDataEnter']
        self.EnterMsg=['没有修改的数据要保存！','没有修改的数据要保存！','没有修改的数据要保存！']
        self.EnterMsgT=json.dumps(self.EnterMsg, encoding="UTF-8", ensure_ascii=False)
    def test_A01_NewtonMethodDataEnter(self):
        '''登录并录入运行数据'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名：%s 密码：%s 返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log('登录成功')
            else:
                self.ReportLog.info_log('登录失败')
            NewtonMethodEnterRun=NewtonMethodDataEnter(self.driver)
            self.EnterReturnMsg=NewtonMethodEnterRun.NewtonMethodDataEnter_Fun(self.DataEnter['LoadP'].split(","),self.DataEnter['LoadQ'].split(","),\
                                                                               self.DataEnter['PowerSupplyP'].split(","),self.DataEnter['PowerSupplyQ'].split(","),\
                                                                               self.DataEnter['PowerSupplyV'].split(","))
            time.sleep(5)
            if self.EnterMsgT==self.EnterReturnMsg:
                self.ReportLog.info_log('未修改输入数据，请确认数据是否为正确的数据')
            else:
                self.ReportLog.info_log('录入数据成功')
        except Exception,e:
            self.ReportLog.info_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("数据无法保存！")
            raise e

    def test_A02_NewtonMethodCalc(self):
        '''登录并计算'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名：%s 密码：%s 返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log('登录成功')
            else:
                self.ReportLog.info_log('登录失败')
            Calc1=NewtonMethodCalc(self.driver)
            self.CalcMsg=Calc1.NewtonMethodCalc_Fun()
            if u'是否要打开当前全网总损耗？'==self.CalcMsg:
                self.ReportLog.info_log('计算成功！')
            else:
                self.ReportLog.info_log('计算失败!')
        except Exception,e:
            self.ReportLog.info_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("计算失败，请检查！")
            raise e
    def test_A03_NewtonMethodWholelossHourCalcResult(self):
        '''整点-输电网线损表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholelossHourCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholelossHourCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'全网总损耗报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath1,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A04_NewtonMethodWholelossMonCalcResult(self):
        '''月-输电网线损表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholelossMonCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholelossMonCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'全网总损耗报表（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath2,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A05_NewtonMethodWholelossDayCalcResult(self):
        '''日-输电网线损表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholelossDayCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholelossDayCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'全网总损耗报表（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath3,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A06_NewtonMethodWholelossMonVoltCalcResult(self):
        '''月-输电网分压线损表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholelossMonVoltCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholelossMonVoltCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'全网分电压损耗报表（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath4,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A07_NewtonMethodWholelossDayVoltCalcResult(self):
        '''日-输电网分压线损表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholelossDayVoltCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholelossDayVoltCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'全网分电压损耗报表（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath5,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A08_NewtonMethodWholelossHourVoltCalcResult(self):
        '''整点-输电网分压线损表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholelossHourVoltCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholelossHourVoltCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'全网分电压损耗报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath6,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A09_NewtonMethodSubstationMonCalcResult(self):
        '''月-变电站损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodSubstationMonCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodSubstationMonCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'变电站损耗报表（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath7,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise
    def test_A10_NewtonMethodSubstationDayCalcResult(self):
        '''日-变电站损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodSubstationDayCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodSubstationDayCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'变电站损耗报表（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath8,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A11_NewtonMethodSubstationHourCalcResult(self):
        '''整点-变电站损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodSubstationHourCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodSubstationHourCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'变电站损耗报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath9,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A12_NewtonMethodTransformerMonCalcResult(self):
        '''月-变压器损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodTransformerMonCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodTransformerMonCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'变压器损耗报表（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath10,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A13_NewtonMethodTransformerDayCalcResult(self):
        '''日-变压器损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodTransformerDayCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodTransformerDayCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'变压器损耗报表（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath11,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A14_NewtonMethodTransformerHourCalcResult(self):
        '''整点-变压器损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodTransformerHourCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodTransformerHourCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'变压器损耗报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath12,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A15_NewtonMethodAclineMonCalcResult(self):
        '''月-导线损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodAclineMonCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodAclineMonCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'导线损耗月报表'
            self.Excel=ExcelToDict.ExcelData(self.filepath13,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A16_NewtonMethodAclineDayCalcResult(self):
        '''日-导线损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodAclineDayCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodAclineDayCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'导线损耗日报表'
            self.Excel=ExcelToDict.ExcelData(self.filepath14,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def     test_A17_NewtonMethodAclineHourCalcResult(self):
        '''整点-导线损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodAclineHourCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodAclineHourCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'导线损耗报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath15,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A18_NewtonMethodOtherLossMonResult(self):
        '''月-其他损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodOtherLossMonResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodOtherLossMonResult_Fun()
            time.sleep(5)
            self.sheetname=u'输电网其他损耗报表（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath16,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A19_NewtonMethodOtherLossDayResult(self):
        '''日-其他损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodOtherLossDayResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodOtherLossDayResult_Fun()
            time.sleep(5)
            self.sheetname=u'输电网其他损耗报表（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath17,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A20_NewtonMethodOtherLossHourResult(self):
        '''整点-其他损耗表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodOtherLossHourResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodOtherLossHourResult_Fun()
            time.sleep(5)
            self.sheetname=u'输电网其他损耗报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath18,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A21_NewtonMethodLosspartHourCalcResult(self):
        '''整点-分片线损表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodLosspartHourCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodLosspartHourCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'全网分片线损报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath19,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A22_NewtonMethodLosspartVoltHourCalcResult(self):
        '''整点-分片分压线损表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodLosspartVoltHourCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodLosspartVoltHourCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'全网分片分压线损报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath20,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A23_NewtonMethodNodeVoltPowerHourCalcResult(self):
        '''整点-节点电压功率表'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodNodeVoltPowerHourCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodNodeVoltPowerHourCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'节点电压功率报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath21,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A24_NewtonMethodWholeLossAnalysisMonResult(self):
        '''月-输电网综合分析'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholeLossAnalysisMonResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholeLossAnalysisMonResult_Fun()
            time.sleep(5)
            self.sheetname=u'输电网综合分析（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath22,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A25_NewtonMethodWholeLossAnalysisMonAcline(self):
        '''月-输电网综合分析-导线追溯'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholeLossAnalysisMonAcline(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholeLossAnalysisMonAcline_Fun()
            time.sleep(5)
            self.sheetname=u'输电网高损线路分析（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath23,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A26_NewtonMethodWholeLossAnalysisMonTrans(self):
        '''月-输电网综合分析-变压器追溯'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholeLossAnalysisMonTrans(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholeLossAnalysisMonTrans_Fun()
            time.sleep(5)
            self.sheetname=u'输电网高损变压器分析（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath24,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A27_NewtonMethodWholeLossAnalysisDayResult(self):
        '''日-输电网综合分析'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholeLossAnalysisDayResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholeLossAnalysisDayResult_Fun()
            time.sleep(5)
            self.sheetname=u'输电网综合分析（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath25,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A28_NewtonMethodWholeLossAnalysisDayAcline(self):
        '''日-输电网综合分析-导线追溯'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholeLossAnalysisDayAcline(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholeLossAnalysisDayAcline_Fun()
            time.sleep(5)
            self.sheetname=u'输电网高损线路分析(日)'
            self.Excel=ExcelToDict.ExcelData(self.filepath26,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A29_NewtonMethodWholeLossAnalysisDayTrans(self):
        '''日-输电网综合分析-变压器追溯'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodWholeLossAnalysisDayTrans(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodWholeLossAnalysisDayTrans_Fun()
            time.sleep(5)
            self.sheetname=u'输电网高损变压器分析（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath27,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A30_NewtonMethodTransAnalysisMonResult(self):
        '''月-变压器经济运行分析'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodTransAnalysisMonResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodTransAnalysisMonResult_Fun()
            time.sleep(5)
            self.sheetname=u'输电网变压器经济运行（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath28,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A31_NewtonMethodTransAnalysisMonRetrospect(self):
        '''月-变压器经济运行分析-重载变追溯'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodTransAnalysisMonRetrospect(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodTransAnalysisMonRetrospect_Fun()
            time.sleep(5)
            self.sheetname=u'输电网变压器损耗明细（月）-重载变'
            self.Excel=ExcelToDict.ExcelData(self.filepath29,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A32_NewtonMethodTransAnalysisDayResult(self):
        '''日-变压器经济运行分析'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodTransAnalysisDayResult(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodTransAnalysisDayResult_Fun()
            time.sleep(5)
            self.sheetname=u'输电网变压器经济运行（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath30,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A33_NewtonMethodTransAnalysisDayRetrospect(self):
        '''日-变压器经济运行分析-重载变追溯'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'wjtest'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=NewtonMethodTransAnalysisDayRetrospect(self.driver)
            self.ReturnMsg=FlowCalcResult.NewtonMethodTransAnalysisDayRetrospect_Fun()
            time.sleep(5)
            self.sheetname=u'输电网变压器损耗明细（日）-重载变'
            self.Excel=ExcelToDict.ExcelData(self.filepath31,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e

    def tearDown(self):
        time.sleep(5)
        self.driver.quit()
