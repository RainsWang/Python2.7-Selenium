#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2018/12/27 16:50
当前项目名称  ：Auto
功能          ：配电线路综合分析日报表-功率因数追溯（配电线路功率因数分析报表）核对
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
    #点击线损分析
    driver.find_element_by_id('Llamenu-panel_header_hd-textEl').click()
    time.sleep(2)
    #点击配电网计算结果分析
    driver.find_element_by_xpath("//div/div/div[6]/div[3]/div/table/tbody/tr[4]/td/div/img").click()
    time.sleep(2)
    #点击日分析
    driver.find_element_by_xpath("//div/div/div[6]/div[3]/div/table/tbody/tr[6]/td/div/img[2]").click()
    time.sleep(2)
    #点击配电线路综合分析
    driver.find_element_by_xpath("//div/div/div[6]/div[3]/div/table/tbody/tr[7]/td/div/img[4]").click()
    #切换到配电线路综合分析的frame中
    driver.switch_to.frame('iframe_4028801b3ec7233c013ec9eeabbe09a9')
    time.sleep(5)
    #点击查询条件中日期的下拉框
    driver.find_element_by_xpath("//div/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]/div").click()
    time.sleep(2)
    driver.find_element_by_xpath("//div/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]/div").click()
    time.sleep(2)
    #切换到日期选择框中的frame
    cf=driver.find_element_by_css_selector("div#ext-gen1041>iframe")
    driver.switch_to.frame(cf)
    driver.implicitly_wait(10)
    #点击日期下拉框年的下拉框
    driver.find_element_by_id("ext-gen1092").click()
    time.sleep(2)
    #选择日期下拉框年中的2015
    driver.find_element_by_xpath("//body/div[4]/div/ul/li[2]").click()
    time.sleep(3)
    #点击日期下拉框代表日的下拉框
    driver.find_element_by_id('ext-gen1095').click()
    time.sleep(2)
    #选择日期下拉框代表日中的2015-08-01
    driver.find_element_by_xpath("//body/div[7]/div/ul/li").click()
    time.sleep(3)
    #点击日期下拉框中的确定按钮
    driver.find_element_by_id('button-1023-btnInnerEl').click()
    time.sleep(3)
    driver.switch_to.parent_frame()   ##跳出日期下拉框的frame
    driver.implicitly_wait(10)
    #点击页面中的查询按钮
    driver.find_element_by_class_name('x-btn-inner').click()
    time.sleep(5)
    #点击功率因数追溯
    driver.find_element_by_xpath("//table/tbody/tr[2]/td[13]/div/a/div/span").click()
    time.sleep(5)
    #跳出配电线路线损分析的frame
    driver.switch_to.default_content()
    driver.implicitly_wait(10)
    #切换到配电线路功率因数分析报表（日）frame中
    driver.switch_to.frame('iframe_4028801b3e83acb7013e86ead6880055')
    time.sleep(5)
    #点击导出按钮
    driver.find_element_by_class_name('x-btn-inner').click()
    time.sleep(3)
except Exception,e:
    print(e)
finally:
    time.sleep(10)
    driver.close()