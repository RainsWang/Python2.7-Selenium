# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/10/19 0019 上午 11:19
当前项目名称  ：Auto
功能          ：日配电线路,需要设置代表日为2014-08-01，代表日需要与手工计算的日期一致
-------------------------------------漂亮的分割线----------------------------------------------
'''


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from Public_Theory import Deletefiles, GetProjectFilePath
from Public_Theory import ExcelToDict
import time,re
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
    time.sleep(1)
    #点击配电网计算结果查询
    driver.find_element_by_xpath("//*[@id = 'Lltmenu-panel-body']/div/table/tbody/tr[5]/td/div/img[1]").click()
    time.sleep(1)
    #点击日报表
    driver.find_element_by_xpath("//*[@id = 'Lltmenu-panel-body']/div/table/tbody/tr[7]/td/div/img[2]").click()
    time.sleep(1)
    #点击配电线路线损表
    driver.find_element_by_xpath("//*[@id = 'Lltmenu-panel-body']/div/table/tbody/tr[8]/td/div").click()
    time.sleep(10)
    #切换frame到配电线路线损表frame
    # a = driver.page_source
    # framemsg = re.findall("iframe name=(.*?) width=",a)
    # print framemsg
    # b = driver.find_element_by_tag_name("iframe")
    # print b
    driver.switch_to.frame("iframe_297efe173af22531013af33990d80000")
    time.sleep(1)
    #点击日期输入框
    # driver.find_element_by_css_selector("#DATE-inputEl").click()
    # time.sleep(1)
    # driver.switch_to.frame("easyXDM_default7523_provider")
    # time.sleep(1)
    # driver.find_element_by_xpath("//div[2]/div/div/table/tbody/tr/td[6]/a").click()
    # time.sleep(1)
    # driver.find_element_by_id("button-1023-btnEl").click()
    # driver.switch_to.parent_frame()
    # driver.find_element_by_xpath("//div/div/div[2]/div/div/div[2]/div/div/div/div[3]/em/button").click()
    # time.sleep(10)



    #点击左下方的导出按钮
    # driver.find_element_by_id("button-1083-btnIconEl").click()
    # time.sleep(10)
    # ExportExecl = ExcelToDict.ExcelData(u"C:\\Users\\Administrator\\Downloads\\配电线路损耗报表（月）.xls",u"配电线路损耗报表（月）")
    # TargetExeclPath = GetProjectFilePath.GetProjectFilePath() + u"\\EagleProjecttest\\Data\\配电线路容量法损耗报表（月）.xls"
    # TargetExecl = ExcelToDict.ExcelData(TargetExeclPath,u"配电线路损耗报表（月）")
    # CmpResult = cmp(TargetExecl.readExcel(),ExportExecl.readExcel())
    # print CmpResult

except Exception,e:
    print e
finally:
    time.sleep(60)
    driver.quit()