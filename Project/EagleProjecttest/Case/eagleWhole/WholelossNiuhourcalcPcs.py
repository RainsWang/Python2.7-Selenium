# -*- coding: utf-8 -*-
#!/usr/bin/env python
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：吴鹏飞
IDE           ：PyCharm Community Edition
时间          ：2018/10/17 10:03
当前项目名称  ：eagle2
功能          ：输电网潮流计算
-------------------------------------漂亮的分割线---------------------------------------------'''
import yaml
import time
from Public_Theory import Deletefiles,ExcelToDict,GetCapture,GetProjectFilePath,GetUsername,Log
from Page_object.Page_object.Login import Login
from Page_object.Page_object.BrowserEngine import BrowserEngine
from Page_object.Page_object.WholelossMeasHourEnter import WholelossMeasHourEnter
from Page_object.Page_object.WholelossNiuhourcalc import WholelossNiuhourcal
from Page_object.Page_object.TranotherlosshourCalcResult import TranotherlosshourCalcResult
from Page_object.Page_object.TranotherlossdayCalcResult import TranotherlossdayCalcResult
from Page_object.Page_object.TranotherlossmonCalcResult import TranotherlossmonCalcResult
import unittest
import json
import os
import getpass
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
class WholelossNiuhourcalcPcs(unittest.TestCase):
    '''24点运行数据录入及核对。'''
    def setUp(self):
        #path=os.path.join(u'C:\\Users\\',getpass.getuser(),'Downloads')
        Deletefiles.del_file()
        self.ProjectFilePath=GetProjectFilePath.GetProjectFilePath()
        self.filepath1=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网其他损耗报表（小时）.xls'
        self.filepath2=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网其他损耗报表（日）.xls'
        self.filepath3=self.ProjectFilePath+u'\EagleProjecttest\Data\eagleWhole\输电网其他损耗报表（月）.xls'
        #日志初始化
        self.ReportLog=Log.LogOutput('理论线损数据输入以及计算')
        self.browser=BrowserEngine()
        self.driver=self.browser.GetDriver()
        self.data_file=open(self.ProjectFilePath+'\\EagleProjecttest\\Data\\eagleWhole\\WholelossNiuhourcalc.yaml')
        self.date=yaml.load(self.data_file)
        self.data_file.close()
        self.date1=self.date['WholelossNiuhourcalc']
        self.LoginMsg=self.date1['Login']
        self.DataEnter=self.date1['WholelossNiuhourcalcDataEntry']
        self.EnterMsg=['没有修改的数据要保存！','没有修改的数据要保存！','没有修改的数据要保存！']
        self.EnterMsgT=json.dumps(self.EnterMsg, encoding="UTF-8", ensure_ascii=False)
        #print self.DataEnter
        #print self.DataEnter['Args']
    def test_A1_WholelossMeasHourEnter(self):
        #登录并录入运行数据
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名：%s 密码：%s 返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log('登录成功')
            else:
                self.ReportLog.info_log('登录失败')
            WholelEnterRun=WholelossMeasHourEnter(self.driver)
            self.EnterReturnMsg=WholelEnterRun.WholelossMeasHourEnter_Fun(self.DataEnter['Args'])
            print self.EnterMsgT
            print self.EnterReturnMsg
            if self.EnterMsgT==self.EnterReturnMsg:
                self.ReportLog.info_log('未修改输入数据，请确认数据是否为正确的数据')
            else:
                self.ReportLog.info_log('录入数据成功')
        except Exception,e:
            self.ReportLog.info_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("数据无法保存！")
            raise e
    def test_A2_WholelossNiuhourcalc(self):
        #登录并计算
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名：%s 密码：%s 返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log('登录成功')
            else:
                self.ReportLog.info_log('登录失败')
            WholeCalc=WholelossNiuhourcal(self.driver)
            self.WholeCalcMsg=WholeCalc.WholelossNiuhourcal_Fun()
            if u'是否要打开当前全网总损耗？'==self.WholeCalcMsg:
                self.ReportLog.info_log('计算成功！')
            else:
                self.ReportLog.info_log('计算失败!')
        except Exception,e:
            self.ReportLog.info_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("计算失败，请检查！")
            raise e
    def test_A3_TranotherlosshourcalcResult(self):
        #登录并查看整点其他损耗表
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名：%s 密码：%s 返回结果：%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log('登录成功！')
            else:
                self.ReportLog.info_log('登录失败！')
            OtherResuslt=TranotherlosshourCalcResult(self.driver)
            self.OtherResusltMsg=OtherResuslt.TranotherlosshourCalcResult_Fun()
            time.sleep(7)
            self.sheetname=u'输电网其他损耗报表（小时）'
            self.Excel=ExcelToDict.ExcelData(self.filepath1,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            # print self.Excel_Data
            # print self.OtherResusltMsg

            self.assertEqual(self.Excel_Data,self.OtherResusltMsg)
            #self.ReportLog.info_log("返回的计算结果%s"%self.OtherResusltMsg)
            self.ReportLog.info_log('数据校验成功！')
        except Exception,e:
            self.ReportLog.info_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("数据校验错误，请检查！")
            raise e
    def test_A4_TranotherlossdayCalcResult(self):
        #登录并查看日其他损耗表
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名：%s 密码：%s 返回结果：%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log('登录成功！')
            else:
                self.ReportLog.info_log('登录失败！')
            OtherResuslt=TranotherlossdayCalcResult(self.driver)
            self.OtherResusltMsg=OtherResuslt.TranotherlossdayCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'输电网其他损耗报表（日）'
            self.Excel=ExcelToDict.ExcelData(self.filepath2,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.OtherResusltMsg)
            #self.ReportLog.info_log("返回的计算结果%s"%self.OtherResusltMsg)
            self.ReportLog.info_log('数据校验成功！')
        except Exception,e:
            self.ReportLog.info_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("数据校验错误，请检查！")
            raise e
    def test_A5_TranotherlossmonCalcResult(self):
        #登录并查看月其他损耗表
        try:
            Loginrun=Login(self.driver)
            ReturnMsg=Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名：%s 密码：%s 返回结果：%s"%(self.LoginMsg['username'],self.LoginMsg['password'],ReturnMsg))
            if 'lj'==ReturnMsg:
                self.ReportLog.info_log('登录成功！')
            else:
                self.ReportLog.info_log('登录失败！')
            OtherResuslt=TranotherlossmonCalcResult(self.driver)
            self.OtherResusltMsg=OtherResuslt.TranotherlossmonCalcResult_Fun()
            time.sleep(5)
            self.sheetname=u'输电网其他损耗报表（月）'
            self.Excel=ExcelToDict.ExcelData(self.filepath3,self.sheetname)
            self.Excel_Data=self.Excel.readExcel()
            self.assertEqual(self.Excel_Data,self.OtherResusltMsg)
            #self.ReportLog.info_log("返回的计算结果%s"%self.OtherResusltMsg)
            self.ReportLog.info_log('数据校验成功！')
        except Exception,e:
            self.ReportLog.info_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("数据校验错误，请检查！")
            raise e

    def tearDown(self):
        time.sleep(10)
        self.driver.quit()










