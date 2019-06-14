#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/12 17:59
当前项目名称  ：620Calc
功能          ：
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
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
    #点击理论线损计算
    driver.find_element_by_id("Lltmenu-panel_header_hd").click()
    time.sleep(2)
    #点击配电线路计算结果查询
    driver.find_element_by_xpath("//div[3]/div[3]/div/table/tbody/tr[5]/td/div/img").click()
    time.sleep(2)
    #点击月报表
    driver.find_element_by_xpath("//div[3]/div[3]/div/table/tbody/tr[6]/td/div/img[2]").click()
    #点击配电线路线损表
    driver.find_element_by_xpath("//div[3]/div[3]/div/table/tbody/tr[7]/td/div/img[4]").click()
    time.sleep(2)
    #切换到结果frame
    driver.switch_to_frame('iframe_297efe173ab9e771013aba8af15d000a')
    driver.implicitly_wait(30)
    for i in range(1089,1124):
        Targetid = 'gridcolumn-'+ str(i) +'-textEl'
        TargetClass = 'x-grid-cell-gridcolumn-'+ str(i)
        target = driver.find_element_by_id(Targetid)
        driver.execute_script("arguments[0].scrollIntoView();", target) #拖动到可见的元素去
        title = driver.find_element_by_id(Targetid).text
        PostmortemPara = driver.find_element_by_class_name(TargetClass).text
        print (title+':  '+PostmortemPara)
        time.sleep(1)
except Exception,e:
    print (e)
finally:
    time.sleep(10)
    driver.close()