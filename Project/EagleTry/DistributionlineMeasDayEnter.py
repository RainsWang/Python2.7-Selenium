#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018-8-8 10:50
当前项目名称  ：620Calc
功能          ：配线日运行数据录入
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import  SendKeys
try:
    #开始登录操作TABTABTAB
    driver = webdriver.Chrome()
    driver.get('http://192.168.0.141:3030/eagle2/coc/application/controller/frame/LoginUI.action')
    driver.maximize_window()
    driver.find_element_by_id('username').clear()
    driver.find_element_by_id('username').send_keys('lyltest')
    driver.find_element_by_id('username').send_keys(Keys.TAB)
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("password").send_keys('000000')
    driver.find_element_by_class_name('loginbutton').click()
    driver.implicitly_wait(10)
    #点击数据中心
    driver.find_element_by_id('Dacmenu-panel_header_hd-textEl').click()
    time.sleep(2)
    #点击运行数据管理
    driver.find_element_by_xpath("//*[@id='Dacmenu-panel-body']/div/table/tbody/tr[4]/td/div/img").click()
    time.sleep(2)
    #点击配电网运行数据录入
    driver.find_element_by_xpath("//div[@id='Dacmenu-panel-body']/div/table/tbody/tr[6]/td/div/img[2]").click()
    time.sleep(2)
    #点击配电网理论运行数据录入
    driver.find_element_by_xpath("//div[@id='Dacmenu-panel-body']/div/table/tbody/tr[8]/td/div/img[4]").click()
    time.sleep(2)
    #点击f11全屏展示
    SendKeys.SendKeys("{F11}")
    time.sleep(2)
    #切换到运行数据录入页面iframe
    driver.switch_to_frame('iframe_120001140402')
    time.sleep(3)
    #点击量测分组框
    driver.find_element_by_id("customGroupBox-inputEl").click()
    time.sleep(5)
    #切换到24点电流
    driver.find_element_by_xpath("//div[10]/div/ul/li[2]").click()
    time.sleep(3)
    #输入配线24点电压数据
    listvoltage=[10.1,10.1,10.1,10.2,10.3,10.1,10.1,10.1,10.1,10.1,10.1,10.5,10.1,10.1,10.7,10.1,10.1,10.1,10.44,10.1,10.1,10.1,10.1,10.1]
    listelec=[1.25,2.56,3.26,2.26,3.26,4.25,3.25,6.25,3.26,1.25,1.25,4.26,1.25,1.25,6.25,1.25,1.25,8.26,9.26,1.25,1.25,1.25,10.23,1.25]
    for i in range(2,26):
        #target = "//[@id = 'measGridLLS-targetEl']/div/div[2]/div/table/tbody/tr["+str(i)+"]"
        target1 = '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[63]/div'
        #driver.execute_script("arguments[0].scrollIntoView();", target)
        #time.sleep(2)
        driver.find_element_by_xpath(target1).click()
        time.sleep(1)
        driver.find_element_by_name("V").clear()
        time.sleep(0.3)
        driver.find_element_by_name("V").send_keys(str(listvoltage[i-2]))
        time.sleep(1)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[68]/div').click()
        time.sleep(0.4)
        driver.find_element_by_name("I").clear()
        time.sleep(0.3)
        driver.find_element_by_name("I").send_keys(str( listelec[i-2]))
        time.sleep(0.4)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
    driver.find_element_by_name("I").send_keys(Keys.TAB)
    time.sleep(5)
    driver.find_element_by_id("save_meas_button-btnInnerEl").click()
    time.sleep(3)
    ConfirmMsg=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    time.sleep(2)
    if u'修改数据保存成功！'==ConfirmMsg:
        driver.find_element_by_id('button-1005-btnEl').click()
        print "配线数据保存成功，请进行理论线损计算"
    else:
        driver.find_element_by_id("button-1005-btnEl").click()
        print "没有要保存的数据！"
    time.sleep(2)
    #点击量测类型分组框
    driver.find_element_by_id("customGroupBox-inputEl").click()
    time.sleep(2)
    #切换到日量测类型分组
    driver.find_element_by_xpath("//div[10]/div/ul/li[3]").click()
    time.sleep(2)
    #录入配线运行数据信息
    driver.find_element_by_xpath("//div[2]/div[2]/div/table/tbody/tr[2]/td[3]/div").click()
    driver.find_element_by_name("PQ_P_A").clear()
    #正向有功电量录入
    driver.find_element_by_name("PQ_P_A").send_keys('6666')
    time.sleep(3)
    #功率因数录入
    driver.find_element_by_xpath("//div[2]/div[2]/div/table/tbody/tr[2]/td[59]/div").click()
    time.sleep(1)
    driver.find_element_by_name("F").clear()
    time.sleep(1)
    driver.find_element_by_name("F").send_keys('0.23')
    time.sleep(3)
    #负荷形状系数录入
    driver.find_element_by_xpath("//div[2]/div[2]/div/table/tbody/tr[2]/td[72]/div").click()
    time.sleep(1)
    driver.find_element_by_name("K").clear()
    time.sleep(1)
    driver.find_element_by_name("K").send_keys('1.23')
    time.sleep(3)
    driver.find_element_by_name("K").send_keys(Keys.TAB)
    time.sleep(1)
    #输入数据后进行保存
    driver.find_element_by_id("save_meas_button-btnEl").click()
    time.sleep(6)
    ConfirmMsg=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    if u'修改数据保存成功！'==ConfirmMsg:
        driver.find_element_by_id('button-1005-btnEl').click()
        print "配线数据保存成功，请进行理论线损计算"
    else:
        driver.find_element_by_id("button-1005-btnEl").click()
        print "没有要保存的数据！"
    time.sleep(4)
    #输入配变数据
    #点击配变tab页面
    driver.find_element_by_id("tab-1162-btnEl").click()
    time.sleep(2)
    #点击量测分组框
    driver.find_element_by_id("customGroupBox-inputEl").click()
    time.sleep(3)
    #点击“理论线损小时数据”
    driver.find_element_by_xpath("//div[10]/div/ul/li[2]").click()
    time.sleep(2)
    #输入配变24点有功功率、无功功率、电压数据
    #输入前25个有功功率、无功功率、电压数据
    #输入第一页数据
    listActivePower=[0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.3653,0.3653,0.3653,0.3653,0.3653,0.3653,0.3653,0.3653,0.3653,0.3653,0.3653,0.3653,0.4653,0.4653,0.4653,0.4653,0.4653,0.4653,0.4653,0.4653,0.4653,0.4653,0.4653,0.4653,0.5653,0.5653,0.5653,0.5653,0.5653,0.5653,0.5653,0.5653,0.5653,0.5653,0.5653,0.5653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653,0.2653]
    listReactivePower=[0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.7954,0.7954,0.7954,0.7954,0.7954,0.7954,0.7954,0.7954,0.7954,0.7954,0.7954,0.7954,0.6954,0.6954,0.6954,0.6954,0.6954,0.6954,0.6954,0.6954,0.6954,0.6954,0.6954,0.6954,0.3954,0.3954,0.3954,0.3954,0.3954,0.3954,0.3954,0.3954,0.3954,0.3954,0.3954,0.3954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954,0.8954]
    listDistributionVolt=[0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4]
    #输入前25个有功功率、无功功率、电压数据
    for i in range(2,27):
        driver.find_element_by_xpath( '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[51]/div').click()
        time.sleep(1)
        driver.find_element_by_name("P").clear()
        time.sleep(0.3)
        driver.find_element_by_name("P").send_keys(str(listActivePower[i-2]))
        time.sleep(1)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[52]/div').click()
        time.sleep(0.4)
        driver.find_element_by_name("Q").clear()
        time.sleep(0.3)
        driver.find_element_by_name("Q").send_keys(str( listReactivePower[i-2]))
        time.sleep(0.4)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[63]/div').click()
        time.sleep(0.4)
        driver.find_element_by_name("V").clear()
        time.sleep(0.3)
        driver.find_element_by_name("V").send_keys(str( listDistributionVolt[i-2]))
        time.sleep(0.4)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
    driver.find_element_by_name("V").send_keys(Keys.TAB)
    time.sleep(2)
        #输入数据后进行保存
    driver.find_element_by_id("save_meas_button-btnEl").click()
    time.sleep(5)
    ConfirmMsg=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    if u'修改数据保存成功！'==ConfirmMsg:
        driver.find_element_by_id('button-1005-btnEl').click()
        print "配变数据保存成功，请进行理论线损计算"
    else:
        driver.find_element_by_id("button-1005-btnEl").click()
        print "没有要保存的数据！"
    #点击向后的翻页
    driver.find_element_by_id("button-1046-btnEl").click()
    time.sleep(2)
    #输入第二页数据
    for i in range(2,27):
        driver.find_element_by_xpath( '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[51]/div').click()
        time.sleep(1)
        driver.find_element_by_name("P").clear()
        time.sleep(0.3)
        driver.find_element_by_name("P").send_keys(str(listActivePower[i+23]))
        time.sleep(1)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[52]/div').click()
        time.sleep(0.4)
        driver.find_element_by_name("Q").clear()
        time.sleep(0.3)
        driver.find_element_by_name("Q").send_keys(str( listReactivePower[i+23]))
        time.sleep(0.4)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[63]/div').click()
        time.sleep(0.4)
        driver.find_element_by_name("V").clear()
        time.sleep(0.3)
        driver.find_element_by_name("V").send_keys(str( listDistributionVolt[i+23]))
        time.sleep(0.4)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
    driver.find_element_by_name("V").send_keys(Keys.TAB)
    time.sleep(2)
    #输入数据后进行保存
    driver.find_element_by_id("save_meas_button-btnEl").click()
    time.sleep(5)
    ConfirmMsg=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    if u'修改数据保存成功！'==ConfirmMsg:
        driver.find_element_by_id('button-1005-btnEl').click()
        print "配变数据保存成功，请进行理论线损计算"
    else:
        driver.find_element_by_id("button-1005-btnEl").click()
        print "没有要保存的数据！"
    time.sleep(5)
    #点击向后的翻页
    driver.find_element_by_id("button-1046-btnEl").click()
    time.sleep(2)
    #输入第三页数据
    for i in range(2,27):
        driver.find_element_by_xpath( '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[51]/div').click()
        time.sleep(1)
        driver.find_element_by_name("P").clear()
        time.sleep(0.3)
        driver.find_element_by_name("P").send_keys(str(listActivePower[i+48]))
        time.sleep(1)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[52]/div').click()
        time.sleep(0.4)
        driver.find_element_by_name("Q").clear()
        time.sleep(0.3)
        driver.find_element_by_name("Q").send_keys(str( listReactivePower[i+48]))
        time.sleep(0.4)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[63]/div').click()
        time.sleep(0.4)
        driver.find_element_by_name("V").clear()
        time.sleep(0.3)
        driver.find_element_by_name("V").send_keys(str( listDistributionVolt[i+48]))
        time.sleep(0.4)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
    driver.find_element_by_name("V").send_keys(Keys.TAB)
    time.sleep(2)
    #输入数据后进行保存
    driver.find_element_by_id("save_meas_button-btnEl").click()
    time.sleep(5)
    ConfirmMsg=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    if u'修改数据保存成功！'==ConfirmMsg:
        driver.find_element_by_id('button-1005-btnEl').click()
        print "配变数据保存成功，请进行理论线损计算"
    else:
        driver.find_element_by_id("button-1005-btnEl").click()
        print "没有要保存的数据！"
    time.sleep(5)
    #点击向后的翻页
    driver.find_element_by_id("button-1047-btnEl").click()
    time.sleep(2)
     #输入第四页数据
    for i in range(2,23):
        driver.find_element_by_xpath( '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[51]/div').click()
        time.sleep(1)
        driver.find_element_by_name("P").clear()
        time.sleep(0.3)
        driver.find_element_by_name("P").send_keys(str(listActivePower[i+72]))
        time.sleep(1)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[52]/div').click()
        time.sleep(0.4)
        driver.find_element_by_name("Q").clear()
        time.sleep(0.3)
        driver.find_element_by_name("Q").send_keys(str( listReactivePower[i+72]))
        time.sleep(0.4)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[63]/div').click()
        time.sleep(0.4)
        driver.find_element_by_name("V").clear()
        time.sleep(0.3)
        driver.find_element_by_name("V").send_keys(str( listDistributionVolt[i+72]))
        time.sleep(0.4)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
        SendKeys.SendKeys("TAB")
        time.sleep(0.3)
    driver.find_element_by_name("V").send_keys(Keys.TAB)
    time.sleep(2)
    #输入数据后进行保存
    driver.find_element_by_id("save_meas_button-btnEl").click()
    time.sleep(5)
    ConfirmMsg=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    if u'修改数据保存成功！'==ConfirmMsg:
        driver.find_element_by_id('button-1005-btnEl').click()
        print "配变数据保存成功，请进行理论线损计算"
    else:
        driver.find_element_by_id("button-1005-btnEl").click()
        print "没有要保存的数据！"
    time.sleep(5)
    #输入配变有功电量数据
    #点击量测类型分组框
    driver.find_element_by_id("customGroupBox-inputEl").click()
    time.sleep(2)
    #点击“日量测类型分组”
    driver.find_element_by_xpath("//div[10]/div/ul/li").click()
    time.sleep(2)
    lista=[222,221,226,227]
    for i in range(2,6):
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[3]/div').click()
        time.sleep(2)
        driver.find_element_by_name("PQ_P_A").clear()
        time.sleep(0.3)
        driver.find_element_by_name("PQ_P_A").send_keys(lista[i-2])
        time.sleep(3)
    driver.find_element_by_name("PQ_P_A").send_keys(Keys.TAB)
    time.sleep(1)
    #输入数据后进行保存
    driver.find_element_by_id("save_meas_button-btnEl").click()
    time.sleep(5)
    ConfirmMsg=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    if u'修改数据保存成功！'==ConfirmMsg:
        driver.find_element_by_id('button-1005-btnEl').click()
        print "配变数据保存成功，请进行理论线损计算"
    else:
        driver.find_element_by_id("button-1005-btnEl").click()
        print "没有要保存的数据！"
except Exception,e:
    print (e)
finally:
    time.sleep(10)
    driver.close()
