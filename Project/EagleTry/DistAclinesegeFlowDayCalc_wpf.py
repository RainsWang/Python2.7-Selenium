#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：吴鹏飞
IDE           ：PyCharm Community Edition
时间          ：2019/1/10 11:22
当前项目名称  ：Project
功能          ：
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys


try:
    driver=webdriver.Chrome()
    driver.get('http://192.168.0.11:3030/eagle2/coc/application/controller/frame/LoginUI.action')
    time.sleep(1)
    driver.maximize_window()
    time.sleep(1)
    driver.find_element_by_css_selector('#username').clear()
    time.sleep(1)
    driver.find_element_by_css_selector('#username').send_keys('lj')
    time.sleep(1)
    driver.find_element_by_css_selector('#username').send_keys(Keys.TAB)
    time.sleep(1)
    driver.find_element_by_css_selector('#password').clear()
    time.sleep(1)
    driver.find_element_by_css_selector('#password').send_keys('000000')
    time.sleep(1)
    driver.find_element_by_css_selector('.loginbutton').click()
    time.sleep(3)
    driver.find_element_by_css_selector('#Dacmenu-panel_header_hd-textEl').click()
    time.sleep(1)
    driver.find_element_by_css_selector('div#treeview-1017>table>tbody>tr:nth-child[3]>td>div>img').click()
    time.sleep(2)



except Exception,e:
    print (e)
finally:
    time.sleep(10)
    driver.close()