# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/9/28 0028 下午 2:52
当前项目名称  ：Auto
功能          ：月容量法计算结果核对
-------------------------------------漂亮的分割线----------------------------------------------
'''

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from Public_Theory import Deletefiles, GetProjectFilePath
from Public_Theory import ExcelToDict
import time
try:
    Deletefiles.del_file()
    driver = webdriver.Chrome()
    driver.get('http://192.168.0.141:3030/eagle2/coc/application/controller/frame/LoginUI.action')
    driver.maximize_window()
    driver.find_element_by_id('username').clear()
    driver.find_element_by_id('username').send_keys('Autotest')
    driver.find_element_by_id('username').send_keys(Keys.TAB)
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("password").send_keys('000000')
    driver.find_element_by_class_name('loginbutton').click()
    driver.implicitly_wait(10)
    # 点击理论线损树
    driver.find_element_by_id('Lltmenu-panel_header_hd').click()
    #点击配电网计算结果查询
    driver.find_element_by_xpath("//*[@id = 'Lltmenu-panel-body']/div/table/tbody/tr[5]/td/div/img[1]").click()
    time.sleep(1)
    #点击月报表
    driver.find_element_by_xpath("//*[@id = 'Lltmenu-panel-body']/div/table/tbody/tr[6]/td/div/img[2]").click()
    time.sleep(1)
    #点击配电线路线损表
    driver.find_element_by_xpath("//*[@id = 'Lltmenu-panel-body']/div/table/tbody/tr[7]/td/div").click()
    time.sleep(1)
    #切换frame到配电线路线损表frame
    driver.switch_to.frame("iframe_297efe173ab9e771013aba8af15d000a")
    time.sleep(1)
    #点击左下方的导出按钮
    driver.find_element_by_id("button-1083-btnIconEl").click()
    time.sleep(10)
    ExportExecl = ExcelToDict.ExcelData(u"C:\\Users\\Administrator\\Downloads\\配电线路损耗报表（月）.xls",u"配电线路损耗报表（月）")
    TargetExeclPath = GetProjectFilePath.GetProjectFilePath() + u"\\EagleProjecttest\\Data\\配电线路容量法损耗报表（月）.xls"
    TargetExecl = ExcelToDict.ExcelData(TargetExeclPath,u"配电线路损耗报表（月）")
    CmpResult = cmp(TargetExecl.readExcel(),ExportExecl.readExcel())
    print CmpResult

except Exception,e:
    print e
finally:
    time.sleep(10)
    driver.quit()
