#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：汪洁
IDE           ：PyCharm Community Edition
时间          ：2018/10/23 13:57
当前项目名称  ：AutoTest
功能          ：
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
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

    #点击页面上方的收缩按钮
    driver.find_element_by_id('header-panel-splitter-collapseEl').click()

    #点击理论线损
    driver.find_element_by_id("Lltmenu-panel_header_hd").click()
    time.sleep(1)

    #点击输电网计算结果查询
    driver.find_element_by_xpath("//div/div/div[3]/div[3]/div/table/tbody/tr[4]/td/div/img").click()
    time.sleep(1)
    #点击整点报表
    driver.find_element_by_xpath("//div/div/div[3]/div[3]/div/table/tbody/tr[7]/td/div/img[2]").click()
    time.sleep(1)
    #点击输电网线损表
    driver.find_element_by_xpath("//div/div/div/div[3]/div[3]/div/table/tbody/tr[8]/td/div").click()
    time.sleep(5)

    #切换到报表的frame
    driver.switch_to_frame('iframe_4028801b3ec1fef3013ec63b58c61ac8')
    #选择开始时间
    driver.find_element_by_xpath("//div[2]/div/div/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div").click()
    time.sleep(0.2)
    driver.find_element_by_xpath("//div[8]/div/table/tbody/tr/td[6]/a/em/span").click()
    time.sleep(0.2)
    driver.find_element_by_xpath("//div[8]/div/div[2]/div[3]/em/button/span").click()
    time.sleep(0.2)
    #选择结束时间
    driver.find_element_by_xpath("//div[2]/div/div/table[2]/tbody/tr/td[2]/table/tbody/tr/td[2]/div").click()
    time.sleep(0.2)
    driver.find_element_by_xpath("//div[9]/div/table/tbody/tr/td[6]/a/em/span").click()
    time.sleep(0.2)
    driver.find_element_by_xpath("//div[9]/div/div[2]/div[3]/em/button/span").click()
    time.sleep(0.2)
    #点击查询
    driver.find_element_by_xpath("//div/div/div[2]/div/div[2]/div/div/div/div[3]/em/button/span").click()
    time.sleep(5)

except Exception,e:
    print (e)
finally:
    time.sleep(10)
    driver.close()