#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2018/12/19 10:37
当前项目名称  ：Auto
功能          ：配电线路潮流精确算法-日报表-配电线路线损表结果核对
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
try:
    #登录操作
    driver=webdriver.Chrome()
    driver.get("http://192.168.0.11:3030/eagle2")
    driver.maximize_window()
    driver.find_element_by_id('username').clear()
    driver.find_element_by_id('username').send_keys('lj')
    driver.find_element_by_id('username').send_keys(Keys.TAB)
    driver.find_element_by_id('password').clear()
    driver.find_element_by_id('password').send_keys('000000')
    driver.find_element_by_class_name('loginbutton').click()
    driver.implicitly_wait(10)
    #点击理论线损
    driver.find_element_by_id('Lltmenu-panel_header_hd-textEl').click()
    time.sleep(1)
    #点击配电网计算结果查询
    driver.find_element_by_xpath('//div/div/div[3]/div[3]/div/table/tbody/tr[5]/td/div/img').click()
    time.sleep(2)
    #点击日报表
    driver.find_element_by_xpath('//div/div/div[3]/div[3]/div/table/tbody/tr[7]/td/div/img[2]').click()
    time.sleep(2)
    #点击配电线路线损表
    driver.find_element_by_xpath('//div/div/div[3]/div[3]/div/table/tbody/tr[8]/td/div/img[4]').click()
    time.sleep(2)
    #切换到配电线路线损表的frame中
    driver.switch_to_frame('iframe_297efe173af22531013af33990d80000')
    driver.implicitly_wait(20)
    #点击导出
    driver.find_element_by_id('button-1083-btnIconEl').click()
    time.sleep(3)
except Exception,e:
    print(e)
finally:
    time.sleep(10)
    driver.close()
