#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/10/17 10:20
当前项目名称  ：AutoTest
功能          ：
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import  SendKeys
try:
    #登陆
    driver=webdriver.Chrome()
    driver.get('http://192.168.0.141:3030/eagle2/coc/application/controller/frame/LoginUI.action')
    driver.maximize_window()
    time.sleep(1)
    driver.find_element_by_id('username').clear()
    #输入用户名
    driver.find_element_by_id('username').send_keys('wjtest')
    time.sleep(1)
    driver.find_element_by_id('password-mask').clear()
    #输入用户密码
    driver.find_element_by_id('password').send_keys('000000')
    time.sleep(1)
    #点击登陆按钮
    driver.find_element_by_class_name('loginbutton').click()
    time.sleep(1)
    driver.implicitly_wait(10)
    #完成登陆

    #点击数据中心
    driver.find_element_by_id('Dacmenu-panel_header_hd-textEl').click()
    time.sleep(2)
    #点击运行数据管理
    driver.find_element_by_xpath("//div/div/div[2]/div[3]/div/table/tbody/tr[4]/td/div/img").click()
    time.sleep(2)
    #点击输电网运行数据录入
    driver.find_element_by_xpath("//div/div[2]/div[3]/div/table/tbody/tr[5]/td/div/img[2]").click()
    time.sleep(2)
    #点击理论线损运行数据录入
    driver.find_element_by_xpath("//div/div/div/div/div[2]/div[3]/div/table/tbody/tr[7]/td/div").click()
    time.sleep(10)
    #已打开数据中心►运行数据管理►输电网运行数据录入►理论线损运行数据录入报表

    #点击页面上方的收缩按钮
    driver.find_element_by_id('header-panel-splitter-collapseEl').click()

    #默认量测分组,为理论线损小时数据，不用修改
    #选择日期,为2014-08-01
    driver.switch_to_frame('iframe_120001140302')
    driver.find_element_by_xpath("//div/div/div[2]/div/div/table[5]/tbody/tr/td[2]/table/tbody/tr/td[3]/div").click()
    time.sleep(1)
    driver.switch_to_frame('ext-gen1039')
    driver.find_element_by_xpath("//div/div/div/div/div[2]/div/div/table/tbody/tr/td[6]/a/em/span").click()
    time.sleep(1)
    driver.find_element_by_xpath("//body/div/div/div/div/div/div[2]/div/div/div/div/em/button/span").click()
    time.sleep(2)

    #切换到负荷tab页面
    driver.switch_to.parent_frame()
    driver.find_element_by_xpath("//body/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div[6]/em/button/span").click()
    time.sleep(3)

    #准备主网计算变电站1-负荷1的运行数据，并录入保存
    activepower=[2.115,2.116,2.117,2.118,2.119,2.12,2.121,2.122,2.123,2.124,2.125,2.126,2.127,2.128,2.129,2.13,2.131,2.132,2.133,2.134,2.135,2.136,2.137,2.138]
    reactivepower=[0.125,0.126,0.127,0.128,0.129,0.13,0.131,0.132,0.133,0.134,0.135,0.136,0.137,0.138,0.139,0.14,0.141,0.142,0.143,0.144,0.145,0.146,0.147,0.148]
    time.sleep(1)
    for i in range(2,26):
        target='//div/div/div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[49]/div'
        driver.find_element_by_xpath(target).click()
        time.sleep(0.2)
        driver.find_element_by_name("P").clear()
        time.sleep(0.2)
        driver.find_element_by_name("P").send_keys(str(activepower[i-2]))
        time.sleep(0.2)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[53]/div').click()
        time.sleep(0.2)
        driver.find_element_by_name("Q").clear()
        time.sleep(0.2)
        driver.find_element_by_name("Q").send_keys(str( reactivepower[i-2]))
        time.sleep(0.2)
    time.sleep(1)
    driver.find_element_by_name("Q").send_keys(Keys.TAB)
    time.sleep(1)
    driver.find_element_by_id('button-1015-btnInnerEl').click()
    time.sleep(1)
    ConfirmMsg1=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    time.sleep(2)
    if ConfirmMsg1==u'修改数据保存成功！':
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站1-负荷1已录入运行数据并成功保存'
    else:
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站1-负荷1数据已录入，不需要重新录入'
    #主网计算变电站1-负荷1的运行数据录入完成

    #准备主网计算变电站1-负荷2的运行数据，并录入保存
    activepower=[2.215,2.216,2.217,2.218,2.219,2.22,2.221,2.222,2.223,2.224,2.225,2.226,2.227,2.228,2.229,2.23,2.231,2.232,2.233,2.234,2.235,2.236,2.237,2.238]
    reactivepower=[0.225,0.226,0.227,0.228,0.229,0.23,0.231,0.232,0.233,0.234,0.235,0.236,0.237,0.238,0.239,0.24,0.241,0.242,0.243,0.244,0.245,0.246,0.247,0.248]
    driver.find_element_by_xpath("//div/div/div[2]/div[2]/div/table/tbody/tr[3]/td[5]/div").click()
    time.sleep(2)
    for i in range(2,26):
        target = '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[49]/div'
        driver.find_element_by_xpath(target).click()
        time.sleep(0.2)
        driver.find_element_by_name("P").clear()
        time.sleep(0.2)
        driver.find_element_by_name("P").send_keys(str(activepower[i-2]))
        time.sleep(0.2)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[53]/div').click()
        time.sleep(0.2)
        driver.find_element_by_name("Q").clear()
        time.sleep(0.2)
        driver.find_element_by_name("Q").send_keys(str( reactivepower[i-2]))
        time.sleep(0.2)
    time.sleep(1)
    driver.find_element_by_name("Q").send_keys(Keys.TAB)
    time.sleep(1)
    driver.find_element_by_id('button-1015-btnInnerEl').click()
    time.sleep(1)
    ConfirmMsg2=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    time.sleep(2)
    if ConfirmMsg2==u'修改数据保存成功！':
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站1-负荷2已录入运行数据并成功保存'
    else:
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站1-负荷2数据已录入，不需要重新录入'
    #主网计算变电站1-负荷2的运行数据录入完成

    #准备主网计算变电站2-负荷1的运行数据，并录入保存
    activepower=[1.515,1.516,1.517,1.518,1.519,1.52,1.521,1.522,1.523,1.524,1.525,1.526,1.527,1.528,1.529,1.53,1.531,1.532,1.533,1.534,1.535,1.536,1.537,1.538]
    reactivepower=[1.025,1.026,1.027,1.028,1.029,1.03,1.031,1.032,1.033,1.034,1.035,1.036,1.037,1.038,1.039,1.04,1.041,1.042,1.043,1.044,1.045,1.046,1.047,1.048]
    driver.find_element_by_xpath("//div/div/div[2]/div[2]/div/table/tbody/tr[4]/td[5]/div").click()
    time.sleep(2)
    for i in range(2,26):
        target = '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[49]/div'
        driver.find_element_by_xpath(target).click()
        time.sleep(0.2)
        driver.find_element_by_name("P").clear()
        time.sleep(0.2)
        driver.find_element_by_name("P").send_keys(str(activepower[i-2]))
        time.sleep(0.2)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[53]/div').click()
        time.sleep(0.2)
        driver.find_element_by_name("Q").clear()
        time.sleep(0.2)
        driver.find_element_by_name("Q").send_keys(str( reactivepower[i-2]))
        time.sleep(0.2)
    time.sleep(1)
    driver.find_element_by_name("Q").send_keys(Keys.TAB)
    time.sleep(1)
    driver.find_element_by_id('button-1015-btnInnerEl').click()
    time.sleep(1)
    ConfirmMsg2=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    time.sleep(2)
    if ConfirmMsg2==u'修改数据保存成功！':
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站2-负荷1已录入运行数据并成功保存'
    else:
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站2-负荷1数据已录入，不需要重新录入'
    #主网计算变电站2-负荷1的运行数据录入完成

    #准备主网计算变电站2-负荷2的运行数据，并录入保存
    activepower=[1.815,1.816,1.817,1.818,1.819,1.82,1.821,1.822,1.823,1.824,1.825,1.826,1.827,1.828,1.829,1.83,1.831,1.832,1.833,1.834,1.835,1.836,1.837,1.838]
    reactivepower=[0.425,0.426,0.427,0.428,0.429,0.43,0.431,0.432,0.433,0.434,0.435,0.436,0.437,0.438,0.439,0.44,0.441,0.442,0.443,0.444,0.445,0.446,0.447,0.448]
    driver.find_element_by_xpath("//div/div/div[2]/div[2]/div/table/tbody/tr[5]/td[5]/div").click()
    time.sleep(2)
    for i in range(2,26):
        target = '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[49]/div'
        driver.find_element_by_xpath(target).click()
        time.sleep(0.2)
        driver.find_element_by_name("P").clear()
        time.sleep(0.2)
        driver.find_element_by_name("P").send_keys(str(activepower[i-2]))
        time.sleep(0.2)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[53]/div').click()
        time.sleep(0.2)
        driver.find_element_by_name("Q").clear()
        time.sleep(0.2)
        driver.find_element_by_name("Q").send_keys(str( reactivepower[i-2]))
        time.sleep(0.2)
    time.sleep(1)
    driver.find_element_by_name("Q").send_keys(Keys.TAB)
    time.sleep(1)
    driver.find_element_by_id('button-1015-btnInnerEl').click()
    time.sleep(1)
    ConfirmMsg2=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    time.sleep(2)
    if ConfirmMsg2==u'修改数据保存成功！':
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站2-负荷2已录入运行数据并成功保存'
    else:
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站2-负荷2数据已录入，不需要重新录入'
    #主网计算变电站2-负荷2的运行数据录入完成

    #切换到PQ节点发电机tab页
    driver.find_element_by_xpath("//body/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div[8]/em/button/span").click()
    time.sleep(3)
    #准备主网计算变电站1-电源1的运行数据，并录入保存
    activepower=[4.81,4.82,4.83,4.84,4.85,4.86,4.87,4.88,4.89,4.9,4.91,4.9,4.89,4.88,4.87,4.86,4.85,4.84,4.83,4.82,4.81,4.8,4.79,4.78]
    reactivepower=[1.825,1.826,1.827,1.828,1.829,1.83,1.831,1.832,1.833,1.834,1.835,1.836,1.837,1.838,1.839,1.84,1.841,1.842,1.843,1.844,1.845,1.846,1.847,1.848]
    #driver.find_element_by_xpath("//div/div/div[2]/div[2]/div/table/tbody/tr[2]/td[49]/div").click()
    time.sleep(2)
    for i in range(2,26):
        target = '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[49]/div'
        driver.find_element_by_xpath(target).click()
        time.sleep(0.2)
        driver.find_element_by_name("P").clear()
        time.sleep(0.2)
        driver.find_element_by_name("P").send_keys(str(activepower[i-2]))
        time.sleep(0.2)
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[53]/div').click()
        time.sleep(0.2)
        driver.find_element_by_name("Q").clear()
        time.sleep(0.2)
        driver.find_element_by_name("Q").send_keys(str( reactivepower[i-2]))
        time.sleep(0.2)
    time.sleep(1)
    driver.find_element_by_name("Q").send_keys(Keys.TAB)
    time.sleep(1)
    driver.find_element_by_id('button-1015-btnInnerEl').click()
    time.sleep(1)
    ConfirmMsg2=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    time.sleep(2)
    if ConfirmMsg2==u'修改数据保存成功！':
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站1-电源1已录入运行数据并成功保存'
    else:
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站1-电源1数据已录入，不需要重新录入'
    #主网计算变电站2-电源1的运行数据录入完成

    #切换到平衡节点发电机tab页
    driver.find_element_by_xpath("//body/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div[10]/em/button/span").click()
    time.sleep(3)
    #准备主网计算变电站1-电源2的运行数据，并录入保存
    voltage=[330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2,330.2]
    #driver.find_element_by_xpath("//div/div/div[2]/div[2]/div/table/tbody/tr[2]/td[49]/div").click()
    time.sleep(2)
    for i in range(2,26):
        target = '//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[61]/div'
        driver.find_element_by_xpath(target).click()
        time.sleep(0.2)
        driver.find_element_by_name("V").clear()
        time.sleep(0.2)
        driver.find_element_by_name("V").send_keys(str(voltage[i-2]))
        time.sleep(0.2)
    time.sleep(1)
    driver.find_element_by_name("V").send_keys(Keys.TAB)
    time.sleep(1)
    driver.find_element_by_id('button-1015-btnInnerEl').click()
    time.sleep(1)
    ConfirmMsg2=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    time.sleep(2)
    if ConfirmMsg2==u'修改数据保存成功！':
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站1-电源2已录入运行数据并成功保存'
    else:
       driver.find_element_by_id('button-1005-btnEl').click()
       print '主网计算变电站1-电源2数据已录入，不需要重新录入'
    #主网计算变电站1-电源2的运行数据录入完成

except Exception,e:
    print (e)
finally:
    time.sleep(3)
    driver.close()