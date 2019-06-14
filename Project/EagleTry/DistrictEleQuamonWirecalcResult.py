#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：王敏
IDE           ：PyCharm
时间          ：2018-12-27 18:51
当前项目名称  ：Project
功能          ：导线损耗报表
-------------------------------------漂亮的分割线---------------------------------------------'''
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

try:
    driver = webdriver.Chrome()
    driver.get("http://192.168.0.141:3030/eagle2/coc/application/controller/frame/LoginUI.action")
    driver.maximize_window()
    driver.find_element_by_id('username').clear()
    driver.find_element_by_id('username').send_keys(u'wm县')
    driver.find_element_by_id('username').send_keys(Keys.TAB)
    driver.find_element_by_id('password').clear()
    driver.find_element_by_id('password').send_keys('000000')
    driver.find_element_by_class_name('loginbutton').click()
    time.sleep(8)
    # 理论手工计算
   # 理论线损ID
    driver.find_element_by_id('Lltmenu-panel_header_hd').click()
    time.sleep(2)

    driver.find_element_by_xpath('//div[3]/div[3]/div/table/tbody/tr[5]/td/div/img').click()
    time.sleep(3)
    driver.find_element_by_xpath('//div[3]/div[3]/div/table/tbody/tr[6]/td/div/img[2]').click()
    time.sleep(2)

    driver.find_element_by_xpath('//div[3]/div[3]/div/table/tbody/tr[8]/td/div').click()
    time.sleep(2)
    driver.switch_to.frame('iframe_4028801b3ef3c580013ef4715d2a22b3')
    time.sleep(2)
    driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td[9]/div/a/div').click()
    time.sleep(2)
    driver.switch_to.parent_frame()
    time.sleep(2)
    driver.switch_to.frame('iframe_4028801b3ea5f3cd013ea63d89bf03f8')
    time.sleep(5)
    driver.find_element_by_id('button-1055-btnIconEl').click()

    time.sleep(2)

except Exception, e:
    print e
finally:
    time.sleep(5)
    driver.quit()