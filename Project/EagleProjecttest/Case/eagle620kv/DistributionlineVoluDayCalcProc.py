# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/10/18 0018 上午 10:50
当前项目名称  ：Auto
功能          ：登录系统，访问数据中心理论线损运行数据录入界面输入日数据信息
访问理论线损手动计算界面进行电容法日计算，并跳转到相关界面核对计算结果信息
-------------------------------------漂亮的分割线----------------------------------------------
'''
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import yaml, unittest,time

from Public_Theory import Log, Deletefiles, GetProjectFilePath, GetCapture,ExcelToDict

from Page_object.Page_object.BrowserEngine import BrowserEngine
from Page_object.Page_object.Login import Login
from Page_object.Page_object.DistributionlineMeasDayEnter import DistributionlineMeasDayEnter
from Page_object.Page_object.DistributionlineVoluDayCalc import DistributionlineVoluDayCalc
from Page_object.Page_object.DistributionlineVoluDayCalcResult import DistributionlineVoluDayCalcResult
from Page_object.Page_object.DistributionlineVoluDayCalcTransLoss import DistributionlineVoluDayCalcTransLoss
from Page_object.Page_object.DistributionlineVoluDayCalcLineLoss import DistributionlineVoluDayCalcLineLoss
class DistributionlineVoluDayCalcProc(unittest.TestCase):
    '''理论线损日容量法计算以及相关结果核对'''
    def setUp(self):
        #初始化日志对象
        self.ReportLog= Log.LogOutput('理论线损容量法日数据输入以及计算')
        #通过浏览器引擎公共方法来启动浏览器
        self.browser = BrowserEngine()
        self.driver = self.browser.GetDriver()
        #清理相关目录下的导出的文件内容
        Deletefiles.del_file()
        #获取到当前文件的目录路径
        self.ProjectFilePath= GetProjectFilePath.GetProjectFilePath()
        #读取yaml中数据信息
        self.data_file = open(self.ProjectFilePath + u'\\EagleProjecttest\\Data\\eagle620kv\\DistributionlineVoluDayCalcProc.yaml')
        self.data = yaml.load(self.data_file)
        self.data_file.close()
        self.LoginMsg = self.data['DistributionlineVoluDayCalcProc']['Login']
        self.DataEntryMsg = self.data['DistributionlineVoluDayCalcProc']['DistributionlineDataEntry']

    def test_A1_DistributionlineMeasDayEnter(self):
        '''理论线损日数据录入，与其它配电线路日数据输入一样,输入2014-8-1日数据信息'''
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
            self.DataEnterRun=DistributionlineMeasDayEnter(self.driver)
            self.DataEnterReturn = self.DataEnterRun.DistributionlineMeasDayEnter_Fun(self.DataEntryMsg['Firstargs'].split(","),self.DataEntryMsg['Secondargs'].split(","))
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

    def test_A2_DistributionlineVoluDayCalc(self):
        '''配电网理论线损日容量法计算'''
        try:
            self.Loginrun = Login(self.driver)
            self.ReturnMsg = self.Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            #print ReturnMsg
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'Autotest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #访问理论线损手工计算界面并进行容量法日计算，如计算完成反正True，失败返回False
            self.ValuDayCalc= DistributionlineVoluDayCalc(self.driver)
            self.ValuDayCalcRe = self.ValuDayCalc.DistributionlineVoluDayCalc_Fun()
            if self.ValuDayCalc:
                self.ReportLog.info_log("计算成功，请对结果进行核对")
            else:
                self.ReportLog.error_log("计算失败，请核对失败内容")

        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("执行计算过程中出现问题，请查看截图")
            raise e

    def test_A3_DistributionlineVoluDayCalcResult(self):
        '''配电线路日容量法计算结果导出并与预期结果进行核对'''
        try:
            self.Loginrun = Login(self.driver)
            self.ReturnMsg = self.Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            #print ReturnMsg
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'Autotest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #进入到理论线损容量法日计算结果查询界面
            self.VoluDayCalcRun= DistributionlineVoluDayCalcResult(self.driver)
            self.VoluDayCalcResult = self.VoluDayCalcRun.DistributionlineVoluDayCalcResult_Fun()
            #获取到预期的结果数据信息
            self.TargetExecl = GetProjectFilePath.GetProjectFilePath() + u"\\EagleProjecttest\\Data\\eagle620kv\\配电线路容量法损耗报表（日）.xls"
            self.TargetExeclObject = ExcelToDict.ExcelData(self.TargetExecl, u"配电线路损耗报表（日）")
            self.TargetValue = self.TargetExeclObject.readExcel()
            #将导出结果与预期结果进行比对
            self.assertEqual(self.VoluDayCalcResult,self.TargetValue)
            self.ReportLog.info_log("计算结果与预期结果匹配")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("导出配电线路日容量法计算结果异常")
            raise e
    def test_A4_DistributionlineVoluDayCalcTransLoss(self):
        '''容量法日计算变压器损耗结果导出和核对'''
        try:
            self.Loginrun = Login(self.driver)
            self.ReturnMsg = self.Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'Autotest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #进入到理论线损容量法日计算变压器损耗显示界面
            self.VoluDayCalcRun= DistributionlineVoluDayCalcTransLoss(self.driver)
            self.VoluDayCalcTransLossResult = self.VoluDayCalcRun.DistributionlineVoluDayCalcTransLoss_Fun()
            #获取到预期的容量法日计算变压器损耗显示界面
            self.TargetExecl = GetProjectFilePath.GetProjectFilePath() + u"\\EagleProjecttest\\Data\\eagle620kv\\配电线路容量法变压器损耗表报表（日）.xls"
            self.TargetExeclObject = ExcelToDict.ExcelData(self.TargetExecl, u"配电变压器损耗表报表（日）")
            self.TargetValue = self.TargetExeclObject.readExcel()
            #将导出结果与预期结果进行比对
            self.assertEqual(self.VoluDayCalcTransLossResult,self.TargetValue)
            self.ReportLog.info_log("计算结果与预期结果匹配")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("配电线路日容量法计算结果变压器损耗异常")
            raise e

    def test_A5_DistributionlineVoluDayCalcLineLoss(self):
        '''容量法日计算导线损耗结果导出和核对'''
        try:
            self.Loginrun = Login(self.driver)
            self.ReturnMsg = self.Loginrun.Login_Function(self.LoginMsg['username'],self.LoginMsg['password'])
            self.ReportLog.info_log("登录用户名:%s 密码:%s返回结果:%s"%(self.LoginMsg['username'],self.LoginMsg['password'],self.ReturnMsg))
            if 'Autotest' == self.ReturnMsg:
                self.ReportLog.info_log("登录成功！")
            else:
                self.ReportLog.error_log("登录失败！")
            #进入到理论线损容量法日计算变压器损耗显示界面
            self.VoluDayCalcRun= DistributionlineVoluDayCalcLineLoss(self.driver)
            self.VoluDayCalcLineLossResult = self.VoluDayCalcRun.DistributionlineVoluDayCalcLineLoss_Fun()
            #获取到预期的容量法日计算变压器损耗显示界面
            self.TargetExecl = GetProjectFilePath.GetProjectFilePath() + u"\\EagleProjecttest\\Data\\eagle620kv\\配电线路容量法导线损耗报表（日）.xls"
            self.TargetExeclObject = ExcelToDict.ExcelData(self.TargetExecl, u"配电线路导线损耗报表（日）")
            self.TargetValue = self.TargetExeclObject.readExcel()
            #将导出结果与预期结果进行比对
            self.assertEqual(self.VoluDayCalcLineLossResult,self.TargetValue)
            self.ReportLog.info_log("计算结果与预期结果匹配")
        except Exception,e:
            self.ReportLog.error_log(e)
            GetCapture.GetCapture(self.driver)
            self.ReportLog.error_log("配电线路日容量法导线损耗结果异常")
            raise e


    def tearDown(self):
        time.sleep(10)
        self.driver.quit()

if __name__ == '__main__':
    unittest.main()

