# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/9/25 0025 下午 1:51
当前项目名称  ：Auto
功能          ：登录系统，访问数据中心理论线损运行数据录入界面输入数据信息
访问理论线损手动计算界面进行电容法月计算，并跳转到相关界面核对计算结果信息
-------------------------------------漂亮的分割线----------------------------------------------
'''

import yaml, unittest,time

from Public_Theory import Log, Deletefiles, GetProjectFilePath, GetCapture,ExcelToDict

from Page_object.Page_object.BrowserEngine import BrowserEngine
from Page_object.Page_object.Login import Login
from Page_object.Page_object.DistributionlineMeasMonthEnter import DistributionlineMeasMonthEnter
from Page_object.Page_object.DistributionlineVoluMonthCalc import DistributionlineVoluMonthCalc
from Page_object.Page_object.DistributionlineVoluMonthCalcResult import DistributionlineVoluMonthCalcResult
from Page_object.Page_object.DistributionlineVoluMonthCalcTransLoss import DistributionlineVoluMonthCalcTransLoss
from Page_object.Page_object.DistributionlineVoluMonthCalcLineLoss import DistributionlineVoluMonthCalcLineLoss

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

class DistributionlineVoluMonthCalcProc(unittest.TestCase):
    '''理论线损月计算容量法计算以及相关结果核对'''
    def setUp(self):
        #初始化日志对象
        self.ReportLog= Log.LogOutput('理论线损月数据输入以及月容量法计算')
        #通过浏览器引擎公共方法来启动浏览器
        self.browser = BrowserEngine()
        self.driver = self.browser.GetDriver()
        #清理相关目录下的导出的文件内容
        Deletefiles.del_file()
        #获取到当前文件的目录路径
        self.ProjectFilePath= GetProjectFilePath.GetProjectFilePath()
        #读取yaml中数据信息
        self.data_file = open(self.ProjectFilePath + u'\\EagleProjecttest\\Data\\eagle620kv\\DistributionlineVoluMonthCalcProc.yaml')
        self.data = yaml.load(self.data_file)
        self.data_file.close()
        self.LoginMsg = self.data['DistributionlineVoluMethodCalcProc']['Login']
        self.DataEntryMsg = self.data['DistributionlineVoluMethodCalcProc']['DistributionlineDataEntry']

    def test_A1_DistributionlineMeasMonthEnter(self):
        '''理论线损月数据录入，与其它配电线路月数据输入一样'''
        try:
            self.Loginrun=Login(self.driver)
            self.ReturnMsg=self.Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            #print ReturnMsg
            time.sleep(3)
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'Autotest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #访问数据录入界面，进行数据的录入

            self.DataEnterRun=DistributionlineMeasMonthEnter(self.driver)
            self.DataEnterReturn = self.DataEnterRun.DistributionlineMeasMonthEnter_Fun(self.DataEntryMsg['ActivePower'],self.DataEntryMsg['PowerFactor'],self.DataEntryMsg['LoadShapeFactor'],self.DataEntryMsg['DistributionTransformData1'],self.DataEntryMsg['DistributionTransformData2'],self.DataEntryMsg['DistributionTransformData3'],self.DataEntryMsg['DistributionTransformData4'])
            time.sleep(5)
            #print self.ReturnMsg1
            if '没有修改的数据要保存！' == self.DataEnterReturn:
                self.ReportLog.info_log("未修改输入数据，请确认数据是否为正确的数据")
            else:
                self.ReportLog.info_log("录入数据成功")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("数据无法正确保存")
            raise e

    def test_A2_DistributionlineVoluMonthCalc(self):
        '''配电网理论线损月容量法计算'''
        try:
            self.Loginrun = Login(self.driver)
            self.ReturnMsg = self.Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            #print ReturnMsg
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'Autotest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #访问理论线损手工计算界面并进行容量法月计算，如计算完成反正True，失败返回False
            self.ValuMonthCalc= DistributionlineVoluMonthCalc(self.driver)
            self.ValuMonthCalcRe = self.ValuMonthCalc.DistributionlineVoluMonthCalc_Fun()
            if self.ValuMonthCalcRe:
                self.ReportLog.info_log("计算成功，请对结果进行核对")
            else:
                self.ReportLog.error_log("计算失败，请核对失败内容")

        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("执行计算过程中出现问题，请查看截图")
            raise e

    def test_A3_DistributionlineVoluMonthCalcResult(self):
        '''配电线路月容量法计算结果导出并与预期结果进行核对'''
        try:
            self.Loginrun = Login(self.driver)
            self.ReturnMsg = self.Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            #print ReturnMsg
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'Autotest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #进入到理论线损容量法月计算结果查询界面
            self.VoluMonthCalcRun= DistributionlineVoluMonthCalcResult(self.driver)
            self.VoluMonthCalcResult = self.VoluMonthCalcRun.DistributionlineVoluMonthCalcResult_Fun()
            #获取到预期的结果数据信息
            self.TargetExecl = GetProjectFilePath.GetProjectFilePath() + u"\\EagleProjecttest\\Data\\eagle620kv\\配电线路容量法损耗报表（月）.xls"
            self.TargetExeclObject = ExcelToDict.ExcelData(self.TargetExecl, u"配电线路损耗报表（月）")
            self.TargetValue = self.TargetExeclObject.readExcel()
            #将导出结果与预期结果进行比对
            self.assertEqual(self.VoluMonthCalcResult,self.TargetValue)
            self.ReportLog.info_log("计算结果与预期结果匹配")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("导出配电线路月容量法计算结果异常")
            raise e
    def test_A4_DistributionlineVoluMonthCalcTransLoss(self):
        '''容量法月计算变压器损耗结果导出和核对'''
        try:
            self.Loginrun = Login(self.driver)
            self.ReturnMsg = self.Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'Autotest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #进入到理论线损容量法月计算变压器损耗显示界面
            self.VoluMonthCalcRun= DistributionlineVoluMonthCalcTransLoss(self.driver)
            self.VoluMonthCalcTransLossResult = self.VoluMonthCalcRun.DistributionlineVoluMonthCalcTransLoss_Fun()
            #获取到预期的容量法月计算变压器损耗显示界面
            self.TargetExecl = GetProjectFilePath.GetProjectFilePath() + u"\\EagleProjecttest\\Data\\eagle620kv\\配电线路容量法变压器损耗表报表（月）.xls"
            self.TargetExeclObject = ExcelToDict.ExcelData(self.TargetExecl, u"配电变压器损耗表报表（月）")
            self.TargetValue = self.TargetExeclObject.readExcel()
            #将导出结果与预期结果进行比对
            self.assertEqual(self.VoluMonthCalcTransLossResult,self.TargetValue)
            self.ReportLog.info_log("计算结果与预期结果匹配")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("导出配电线路月容量法计算结果异常")
            raise e

    def test_A5_DistributionlineVoluMonthCalcLineLoss(self):
        '''容量法月计算变压器损耗结果导出和核对'''
        try:
            self.Loginrun = Login(self.driver)
            self.ReturnMsg = self.Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'Autotest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #进入到理论线损容量法月计算变压器损耗显示界面
            self.VoluMonthCalcRun= DistributionlineVoluMonthCalcLineLoss(self.driver)
            self.VoluMonthCalcLineLossResult = self.VoluMonthCalcRun.DistributionlineVoluMonthCalcLineLoss_Fun()
            #获取到预期的容量法月计算变压器损耗显示界面
            self.TargetExecl = GetProjectFilePath.GetProjectFilePath() + u"\\EagleProjecttest\\Data\\eagle620kv\\配电线路容量法导线损耗报表（月）.xls"
            self.TargetExeclObject = ExcelToDict.ExcelData(self.TargetExecl, u"配电线路导线损耗报表（月）")
            self.TargetValue = self.TargetExeclObject.readExcel()
            #将导出结果与预期结果进行比对
            self.assertEqual(self.VoluMonthCalcLineLossResult,self.TargetValue)
            self.ReportLog.info_log("计算结果与预期结果匹配")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("导出配电线路月容量法计算结果异常")
            raise e


    def tearDown(self):
        time.sleep(10)
        self.driver.quit()

if __name__ == '__main__':
    unittest.main()

