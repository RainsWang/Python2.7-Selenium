# -*- coding: utf-8 -*-
#!/usr/bin/env python
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2018/12/17 15:20
当前项目名称  ：eagle2
功能          ：6-20kV潮流精确算法日计算整体流程
-------------------------------------漂亮的分割线---------------------------------------------'''
import yaml
import time
import unittest
import json
from Public_Theory import Deletefiles,GetProjectFilePath,Log,GetCapture,ExcelToDict
from Page_object.Page_object.Login import Login
from Page_object.Page_object.DistributionlineMeasDayEnter import DistributionlineMeasDayEnter
from Page_object.Page_object.DistributionlineFlowDayCalc import DistributionlineFlowDayCalc
from Page_object.Page_object.DistributionlineFlowDayCalcResult import DistributionlineFlowDayCalcResult
from Page_object.Page_object.DistAclinesegeFlowDayCalcResult import DistAclinesegeFlowDayCalcResult
from Page_object.Page_object.DistTransformerFlowDayCalcResult import DistTransformerFlowDayCalcResult
from Page_object.Page_object.DistributionlineComAStDay import DistributionlineComAStDay
from Page_object.Page_object.DistributionlineComAAOLLCDay import DistributionlineComAAOLLCDay
from Page_object.Page_object.DistributionlineComAPowFADay import DistributionlineComAPowFADay
from Page_object.Page_object.DistributionlineComAHlAclDay import DistributionlineComAHlAclDay
from Page_object.Page_object.DistributionlineComAHlTraDay import DistributionlineComAHlTraDay
from Page_object.Page_object.DistributionlineComAEconTranDay import DistributionlineComAEconTranDay
from Page_object.Page_object.DistributionlineOverloadTranDay import  DistributionlineOverloadTranDay
from Page_object.Page_object.DistributionlineEconomicTranDay import DistributionlineEconomicTranDay
from Page_object.Page_object.DistributionlineLightloadTranDay import DistributionlineLightloadTranDay
from Page_object.Page_object.DistributionlinePieSStDay import DistributionlinePieSStDay
from Page_object.Page_object.DistributionlinePieSComDay import DistributionlinePieSComDay
from Page_object.Page_object.DistributionlinePieSSubDay import DistributionlinePieSSubDay
from Page_object.Page_object import BrowserEngine
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

class DistributionlineFlowDayCalcProc(unittest.TestCase):
    '''6-20kV潮流精确算法日计算整体流程。'''
    def setUp(self):
        Deletefiles.del_file()
        self.ProjectFilePath=GetProjectFilePath.GetProjectFilePath()
        self.filepath1=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电线路损耗报表潮流（日）.xls'
        self.filepath2=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电线路导线损耗报表潮流（24点）.xls'
        self.filepath3=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电变压器损耗报表潮流（24点）.xls'
        self.filepath4=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电线路线损综合分析报表.xls'
        self.filepath5=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电线路线损组成分析(日).xls'
        self.filepath6=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电线路功率因数分析报表(日).xls'
        self.filepath7=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电线路高损导线分析报表.xls'
        self.filepath8=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电网高损配变分析报表(日).xls'
        self.filepath9=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配变运行综合分析（日）.xls'
        self.filepath10=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配变损耗组成分析(日)-重载变.xls'
        self.filepath11=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配变损耗组成分析(日)-经济变.xls'
        self.filepath12=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配变损耗组成分析(日)-轻载变.xls'
        self.filepath13=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电线路线损率分段统计表样一(日).xlsx'
        self.filepath14=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电线路线损率分段统计表样二(日).xlsx'
        self.filepath15=self.ProjectFilePath+u'\EagleProjecttest\Data\eagle620kv\配电线路线损率分段统计表样三(日).xlsx'



        #日志初始化
        self.ReportLog= Log.LogOutput('理论线损数据输入以及计算')
        self.Browser=BrowserEngine.BrowserEngine()
        self.driver=self.Browser.GetDriver()
        self.data_file=open(self.ProjectFilePath+'\\EagleProjecttest\\Data\\eagle620kv\\DistributionlineFlowDayCalcProc.yaml')
        self.data=yaml.load(self.data_file)
        self.data_file.close()
        self.data1=self.data['DistributionlineFlowDayCalcProc']
        self.LoginMsg=self.data1['Login']
        self.DataEnter=self.data1['DistributionlineDataEntry']
        self.EnterMsgDay=['没有修改的数据要保存！','没有修改的数据要保存！']
        self.EnterMsgDayT=json.dumps(self.EnterMsgDay, encoding="UTF-8", ensure_ascii=False)
        self.EnterMsgHour=['没有修改的数据要保存！','没有修改的数据要保存！','没有修改的数据要保存！','没有修改的数据要保存！','没有修改的数据要保存！']
        self.EnterMsgHourT=json.dumps(self.EnterMsgHour, encoding="UTF-8", ensure_ascii=False)
    def test_A01_DistributionlineMeasDayEnter(self):
        '''登录并录入配线、配变日运行数据'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            self.DistributionlineDayData=self.DataEnter['DistributionlineDayData'].split(",")
            self.DistributionTransformDayData=self.DataEnter['DistributionTransformDayData'].split(",")
            DataEnterrun=DistributionlineMeasDayEnter(self.driver)
            self.ReturnMsg1=DataEnterrun.DistributionlineMeasDayEnter_Fun(self.DistributionlineDayData,self.DistributionTransformDayData)
            if self.EnterMsgDayT==self.ReturnMsg1:
                self.ReportLog.info_log(u'未修改输入数据，请确认数据是否为正确的数据')
            else:
                self.ReportLog.info_log(u'录入数据成功')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据无法正确保存")
            raise e
    def test_A02_DistributionlineMeasDayEnter(self):
        '''登录并录入配变24点功率、电压'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            self.DistributionlineVoltHourData=self.DataEnter['DistributionlineVoltHourData'].split(",")
            self.DistributionlineElcHourData=self.DataEnter['DistributionlineElcHourData'].split(",")
            self.DistributionTransformActivePowerHourData=self.DataEnter['DistributionTransformActivePowerHourData'].split(",")
            self.DistributionTransformReactivePowerHourData=self.DataEnter['DistributionTransformReactivePowerHourData'].split(",")
            self.DistributionTransformVoltHourData=self.DataEnter['DistributionTransformVoltHourData'].split(",")
            DataEnterrun=DistributionlineMeasDayEnter(self.driver)
            self.ReturnMsg2=DataEnterrun.DistributionlineMeasHourEnter_Fun(self.DistributionlineVoltHourData,self.DistributionlineElcHourData,self.DistributionTransformActivePowerHourData,self.DistributionTransformReactivePowerHourData,self.DistributionTransformVoltHourData)
            if self.EnterMsgHourT==self.ReturnMsg2:
                self.ReportLog.info_log(u'未修改输入数据，请确认数据是否为正确的数据')
            else:
                self.ReportLog.info_log(u'录入数据成功')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据无法正确保存")
            raise e
    def test_A03_DistributionlineFlowDayCalc(self):
        '''配电线路潮流精确算法计算'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalc=DistributionlineFlowDayCalc(self.driver)
            self.ReturnMsg=FlowCalc.DistributionlineFlowDayCalc_Fun()
            if u'是否要打开当前配线的计算结果？'==self.ReturnMsg:
                self.ReportLog.info_log(u'计算成功！')
            else:
                self.ReportLog.info_log(u'计算失败！')
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"计算失败！")
            raise e
    def test_A04_DistributionlineFlowDayCalcResult(self):
        '''潮流精确算法配电线路线损表结果验证'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlineFlowDayCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlineFlowDayCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'配电线路损耗报表（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath1,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A05_DistAclinesegeFlowDayCalcResult(self):
        '''潮流精确算法配线导线损耗结果验证'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistAclinesegeFlowDayCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.DistAclinesegeFlowDayCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'配电线路导线损耗报表（24点）'
            self.Excel=ExcelToDict.ExcelData(self.filepath2,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A06_DistTransformerFlowDayCalcResult(self):
        '''潮流精确算法配电变压器损耗结果验证'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistTransformerFlowDayCalcResult(self.driver)
            self.ReturnMsg=FlowCalcResult.DistTransformerFlowDayCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'配电变压器损耗报表（24点）'
            self.Excel=ExcelToDict.ExcelData(self.filepath3,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A07_DistributionlineComAStDay(self):
        '''潮流精确算法-配电线路综合分析日报表结果核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlineComAStDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlineComAStDay_Fun()
            time.sleep(5)
            self.sheetname=u'配电线路线损综合分析报表'
            self.Excel=ExcelToDict.ExcelData(self.filepath4,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    # def test_A08_DistributionlineComAAOLLCDay(self):
    #     '''配线综合分析-配线追溯页面结果验证'''
    #     try:
    #         Loginrun=Login(self.driver)
    #         ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
    #         time.sleep(3)
    #         self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
    #         if 'lj'==ReturnMsg:
    #             self.ReportLog.info_log(u'登录成功！')
    #         else:
    #             self.ReportLog.info_log(u'登录失败！')
    #         FlowCalcResult=DistributionlineComAAOLLCDay(self.driver)
    #         self.ReturnMsg=FlowCalcResult.DistributionlineComAAOLLCDay_Fun()
    #         time.sleep(5)
    #         self.sheetname1=u'配电线路线损组成分析(日)'
    #         self.sheetname2=u'图表数据'
    #         self.Excel1=ExcelToDict.ExcelData(self.filepath5,self.sheetname1)
    #         self.Excel2=ExcelToDict.ExcelData(self.filepath5,self.sheetname2)
    #         self.Excel_Data1=self.Excel1.readExcel()
    #         self.Excel_Data2=self.Excel2.readExcel()
    #         self.assertEqual(self.Excel_Data1,self.ReturnMsg)
    #         self.ReportLog.info_log(u'数据校验成功！')
    #         #print self.Excel_Data
    #         #print self.ReturnMsg
    #     except Exception,e:
    #         self.ReportLog.error_log(e)
    #         GetCapture.GetCapture(self.driver)
    #         self.ReportLog.error_log(u"数据校验错误，请检查！")
    #         raise e
    def test_A09_DistributionlineComAPowFADay(self):
        '''配线综合分析-功率因数追溯结果报表核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlineComAPowFADay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlineComAPowFADay_Fun()
            time.sleep(5)
            self.sheetname=u'Sheet1'
            self.Excel=ExcelToDict.ExcelData(self.filepath6,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A10_DistributionlineComAHlAclDay(self):
        '''配线综合分析-导线损耗电量追溯结果报表核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlineComAHlAclDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlineComAHlAclDay_Fun()
            time.sleep(5)
            self.sheetname=u'配电线路高损导线分析报表'
            self.Excel=ExcelToDict.ExcelData(self.filepath7,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A11_DistributionlineComAHlTraDay(self):
        '''配线综合分析-变压器损耗电量追溯结果报表核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlineComAHlTraDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlineComAHlTraDay_Fun()
            time.sleep(6)
            self.sheetname=u'配电网高损配变分析报表(日)'
            self.Excel=ExcelToDict.ExcelData(self.filepath8,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A12_DistributionlineComAEconTranDay(self):
        '''配线综合分析-变压器台数追溯结果报表核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlineComAEconTranDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlineComAEconTranDay_Fun()
            time.sleep(5)
            self.sheetname=u'配变运行综合分析（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath9,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A13_DistributionlineOverloadTranDay(self):
        '''配线综合分析-重载变追溯结果核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlineOverloadTranDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlineOverloadTranDay_Fun()
            time.sleep(5)
            self.sheetname=u'配变损耗组成分析(日)-重载变'
            self.Excel=ExcelToDict.ExcelData(self.filepath10,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A14_DistributionlineEconomicTranDay(self):
        '''配线综合分析-经济变追溯结果核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlineEconomicTranDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlineEconomicTranDay_Fun()
            time.sleep(5)
            self.sheetname=u'配变损耗组成分析(日)-经济变'
            self.Excel=ExcelToDict.ExcelData(self.filepath11,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A15_DistributionlineLightloadTranDay(self):
        '''配线综合分析-轻载变追溯结果核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlineLightloadTranDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlineLightloadTranDay_Fun()
            time.sleep(6)
            self.sheetname=u'配变损耗组成分析(日)-轻载变'
            self.Excel=ExcelToDict.ExcelData(self.filepath12,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A16_DistributionlinePieSStDay(self):
        '''配电线路线损率分段统计结果核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlinePieSStDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlinePieSStDay_Fun()
            time.sleep(5)
            self.sheetname=u'Sheet1'
            self.Excel=ExcelToDict.ExcelData(self.filepath13,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A17_DistributionlinePieSComDay(self):
        '''配电线路线损率分段单位追溯统计结果核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlinePieSComDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlinePieSComDay_Fun()
            time.sleep(5)
            self.sheetname=u'Sheet1'
            self.Excel=ExcelToDict.ExcelData(self.filepath14,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e
    def test_A18_DistributionlinePieSSubDay(self):
        '''配电线路线损率分段变电站追溯统计结果核对'''
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log(u'登录成功！')
            else:
                self.ReportLog.info_log(u'登录失败！')
            FlowCalcResult=DistributionlinePieSSubDay(self.driver)
            self.ReturnMsg=FlowCalcResult.DistributionlinePieSSubDay_Fun()
            time.sleep(5)
            self.sheetname=u'Sheet1'
            self.Excel=ExcelToDict.ExcelData(self.filepath15,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.ReturnMsg)
            self.ReportLog.info_log(u'数据校验成功！')
            #print self.Excel_Data
            #print self.ReturnMsg
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log(u"数据校验错误，请检查！")
            raise e





    def tearDown(self):
        time.sleep(10)
        self.driver.quit()





