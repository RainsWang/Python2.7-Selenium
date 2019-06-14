#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/4 17:51
当前项目名称  ：6-20计算
功能          ：录入6-20kv计算所需要的运行数据
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

try:
    #开始登录操作
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
     #点击页面的收缩按钮
    driver.find_element_by_id('header-panel-splitter-collapseEl').click()
    time.sleep(3)
    #量测类型选择理论月运行数据
    driver.switch_to_frame('iframe_120001140402')
    time.sleep(3)
    driver.find_element_by_id("customGroupBox-inputEl").click()
    time.sleep(3)
    #driver.find_element_by_css_selector("#boundlist-1842-listEl > ul > li:nth-child(1)").click()
    #driver.find_element_by_xpath("//div[@id='uploadFormMonth']/div/ul/li").click()
    driver.find_element_by_xpath("//div[10]/div/ul/li").click()
    time.sleep(2)
    #选择变电站
    driver.find_element_by_id("substationId-inputEl").click()
    time.sleep(2)

    #录入配线运行数据信息
    driver.find_element_by_xpath("//div[2]/div[2]/div/table/tbody/tr[2]/td[3]/div").click()
    driver.find_element_by_name("PQ_P_A").clear()
    #正向有功电量录入
    driver.find_element_by_name("PQ_P_A").send_keys('6666')
    time.sleep(2)
    #功率因数录入
    driver.find_element_by_xpath("//div[2]/div[2]/div/table/tbody/tr[2]/td[59]/div").click()
    driver.find_element_by_name("F").clear()
    driver.find_element_by_name("F").send_keys('0.23')
    time.sleep(2)
    #负荷形状系数录入
    driver.find_element_by_xpath("//div[2]/div[2]/div/table/tbody/tr[2]/td[72]/div").click()
    driver.find_element_by_name("K").clear()
    driver.find_element_by_name("K").send_keys('1.23')
    time.sleep(2)
    #输入数据后进行保存
    driver.find_element_by_id("save_meas_button-btnInnerEl").click()
    ConfirmMsg=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    if u'修改数据保存成功！'==ConfirmMsg:
        driver.find_element_by_id('button-1005-btnEl').click()
        print "配线数据保存成功，请进行理论线损计算"
    else:
        driver.find_element_by_id("button-1005-btnEl").click()
        print "没有要保存的数据！"
    time.sleep(2)

    #输入配变数据
    driver.find_element_by_id("tab-1162-btnEl").click()
    time.sleep(2)
    lista=[222,221,226,227]
    for i in range(2,6):
        driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr['+str(i)+']/td[3]/div').click()
        time.sleep(2)
        driver.find_element_by_name("PQ_P_A").clear()
        driver.find_element_by_name("PQ_P_A").send_keys(lista[i-2])
        time.sleep(3)
    driver.find_element_by_name("PQ_P_A").send_keys(Keys.TAB)
    #输入数据后进行保存
    driver.find_element_by_id("save_meas_button-btnEl").click()
    time.sleep(2)
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




