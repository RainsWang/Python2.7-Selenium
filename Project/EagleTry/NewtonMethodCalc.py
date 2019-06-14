#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/10/22 16:33
当前项目名称  ：AutoTest
功能          ：
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
try:
    #登陆
    driver=webdriver.Chrome()
    driver.get('http://192.168.0.141:3030/eagle2/coc/application/controller/frame/LoginUI.action')
    driver.maximize_window()
    time.sleep(0.2)
    driver.find_element_by_id('username').clear()
    #输入用户名
    driver.find_element_by_id('username').send_keys('wjtest')
    time.sleep(0.2)
    driver.find_element_by_id('password-mask').clear()
    #输入用户密码
    driver.find_element_by_id('password').send_keys('000000')
    time.sleep(0.2)
    #点击登陆按钮
    driver.find_element_by_class_name('loginbutton').click()
    time.sleep(0.2)
    driver.implicitly_wait(10)
    #完成登陆

    # 点击理论线损
    driver.find_element_by_id('Lltmenu-panel_header_hd').click()
    time.sleep(1)
    # 点击理论线损手工计算
    driver.find_element_by_xpath("//div[@id ='Lltmenu-panel-body']/div/table/tbody/tr[2]/td/div").click()
    time.sleep(5)
    #切换iframe
    driver.switch_to_frame('iframe_13000110')
    #默认进入输电网理论线损计算页面，选择单位
    driver.find_element_by_xpath("//div[2]/div/div/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td/div/input").click()
    time.sleep(5) 
    #点击计算
    driver.find_element_by_css_selector('#calc_start_button').click()
    driver.implicitly_wait(60)
    alerttext = driver.find_element_by_id("messagebox-1001-displayfield-inputEl").text
    print u'弹出提示内容为：'
    print alerttext
    if u'是否要打开当前全网总损耗？' == alerttext:
        print u"计算成功"
    else:
        print u"计算失败"
except Exception,e:
    raise e
finally:
    time.sleep(10)
    driver.quit()