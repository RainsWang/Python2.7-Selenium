#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2018/12/27 17:32
当前项目名称  ：Auto
功能          ：配电线路线损率分段统计结果核对
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
try:
    #登录操作
    driver=webdriver.Chrome()
    driver.get("http://192.168.0.11:3030/eagle2")
    driver.maximize_window()
    #driver.find_element_by_id('username').clear()
    driver.find_element_by_css_selector("#username").clear()
    #driver.find_element_by_id('username').send_keys('lj')
    driver.find_element_by_css_selector("#username").send_keys('lj')
    #driver.find_element_by_id('username').send_keys(Keys.TAB)
    driver.find_element_by_css_selector("#username").send_keys(Keys.TAB)
    #driver.find_element_by_id('password').clear()
    driver.find_element_by_css_selector("#password").clear()
    #driver.find_element_by_id('password').send_keys('000000')
    driver.find_element_by_css_selector("#password").send_keys('000000')
    #driver.find_element_by_class_name('loginbutton').click()
    driver.find_element_by_css_selector("[class='loginbutton']").click()
    driver.implicitly_wait(10)
    #点击线损分析
    #driver.find_element_by_id('Llamenu-panel_header_hd-textEl').click()
    driver.find_element_by_css_selector("#Llamenu-panel_header_hd-textEl").click()
    time.sleep(2)
    #点击配电网计算结果分析
    driver.find_element_by_xpath('//div/div/div[6]/div[3]/div/table/tbody/tr[4]/td/div/img').click()
    time.sleep(2)
    #点击日分析
    driver.find_element_by_xpath('//div/div/div[6]/div[3]/div/table/tbody/tr[6]/td/div/img[2]').click()
    time.sleep(2)
    #点击配电线路线损率分段统计
    driver.find_element_by_xpath('//div/div/div[6]/div[3]/div/table/tbody/tr[9]/td/div').click()
    time.sleep(2)
    #切换到配电线路线损率分段统计的frame中
    driver.switch_to.frame('iframe_4028801b404d7a9401404d7dd2530154')
    time.sleep(5)
    #点击导出
    driver.find_element_by_id("4028801b3ff1c37d013ff48d9e350012-btnInnerEl").click()
    time.sleep(3)
except Exception,e:
    print(e)
finally:
    time.sleep(10)
    driver.close()