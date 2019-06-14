#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2018/12/28 15:13
当前项目名称  ：Auto
功能          ：配电线路线损率分段-变电站追溯结果核对
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
    time.sleep(5)
    #点击线损分析
    driver.find_element_by_id('Llamenu-panel_header_hd-textEl').click()
    time.sleep(2)
    #点击配电网计算结果分析
    driver.find_element_by_xpath("//div/div/div[6]/div[3]/div/table/tbody/tr[4]/td/div/img").click()
    time.sleep(2)
    #点击日分析
    driver.find_element_by_xpath("//div/div/div[6]/div[3]/div/table/tbody/tr[6]/td/div/img[2]").click()
    time.sleep(2)
    #点击配电线路线损率分段统计
    driver.find_element_by_xpath("//div/div/div[6]/div[3]/div/table/tbody/tr[9]/td/div").click()
    time.sleep(2)
    #切换到配电线路线损率分段统计的frame中
    driver.switch_to.frame('iframe_4028801b404d7a9401404d7dd2530154')
    time.sleep(5)
    #切换到配电线路线损率分段统计查询结果的frame中
    driver.switch_to.frame('viewframe')
    time.sleep(5)
    #点击单位追溯
    driver.find_element_by_link_text(u"计算结果图形").click()
    time.sleep(3)
    driver.switch_to.parent_frame()   ##跳出结果中的frame
    time.sleep(5)
    driver.switch_to.default_content()   ##跳出刚进入页面的frame
    time.sleep(5)
    #切换到配电网线损率分段统计追溯报表（分单位日）frame中
    driver.switch_to.frame('iframe_4028801b4033ffd5014037f37eff007a')
    time.sleep(5)
    #切换到配电网线损率分段统计追溯报表（分单位日）查询结果的frame中
    driver.switch_to.frame('viewframe')
    time.sleep(5)
    #点击变电站名称追溯
    driver.find_element_by_link_text(u"东河110kV变电站").click()
    time.sleep(3)
    driver.switch_to.parent_frame()   ##跳出结果中的frame
    time.sleep(5)
    driver.switch_to.parent_frame()   ##跳出刚进入页面的frame
    time.sleep(5)
    #切换到配电网线损率分段统计追溯报表（分变电站日）的frame中
    driver.switch_to.frame('iframe_4028801b4033ffd5014037f3cbd00084')
    time.sleep(5)
    #点击导出
    driver.find_element_by_id('4028801b4033ffd5014037f3cbd20088-btnInnerEl').click()
    time.sleep(5)
except Exception,e:
    print(e)
finally:
    time.sleep(7)
    driver.close()


