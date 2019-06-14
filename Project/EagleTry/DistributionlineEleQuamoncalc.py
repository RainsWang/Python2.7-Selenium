#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/6 19:59
当前项目名称  ：126login
功能          ：配线月电量法计算
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

try:
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

    #点击理论线损计算
    driver.find_element_by_id("Lltmenu-panel_header_hd").click()
    driver.find_element_by_xpath("//div[3]/div[3]/div/table/tbody/tr[2]/td/div/img[2]").click()
    time.sleep(2)
    #切换到配电线路理论计算页面
    driver.switch_to_frame('iframe_13000110')
    driver.find_element_by_id("tab-1154-btnInnerEl").click()
    time.sleep(2)
    driver.find_element_by_xpath("//div[2]/div/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td/div/img").click()
    time.sleep(2)
    driver.find_element_by_xpath("//div[2]/div[2]/div/div/div[2]/div/div[2]/div/table/tbody/tr[3]/td/div/img[2]").click()
    time.sleep(2)
    driver.find_element_by_xpath("//div[2]/div[2]/div/div/div[2]/div/div[2]/div/table/tbody/tr[4]/td/div/input").click()

    #点击计算
    driver.find_element_by_id("district620_calc_start_button-btnEl").click()
    driver.implicitly_wait(120)

    alerttext=driver.find_element_by_id("messagebox-1001-displayfield-inputEl").text
    if alerttext==u'是否要打开当前配线的计算结果？':
        driver.find_element_by_id("button-1007-btnInnerEl").click()
        print "计算成功，跳出页面"
    else:
        print "无法正常计算"
    time.sleep(20)
except Exception,e:
    print (e)
finally:
    time.sleep(10)
    driver.close()
