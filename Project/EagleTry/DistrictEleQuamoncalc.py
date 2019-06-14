#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：王敏
IDE           ：PyCharm
时间          ：2018-12-25 10:16
当前项目名称  ：Project
功能          ：理论手工计算
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
    driver.find_element_by_xpath('//div[3]/div[2]/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[2]/td/div/img[2]').click()
    time.sleep(7)
  #切换到配电网400V理论线损所在frame并点击
    driver.switch_to.frame('iframe_13000110')
    time.sleep(2)
    driver.find_element_by_id('tab-1163-btnInnerEl').click()
    time.sleep(2)
  #直接选择计算的台区
  #勾选wm县电量法配线月结果验证，总共四层，单位、变电站、配电线路、台区
  #单位xpath
    driver.find_element_by_xpath('//div/div[2]/div[2]/div/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td/div/img').click()
    time.sleep(2)
  #变电站xpath
    driver.find_element_by_xpath('//div/div[2]/div[2]/div/div/div[2]/div/div[2]/div/table/tbody/tr[3]/td/div/img[2]').click()
    time.sleep(2)
  #配电线路
    driver.find_element_by_xpath('//div/div[2]/div[2]/div/div/div[2]/div/div[2]/div/table/tbody/tr[4]/td/div/img[3]').click()
    time.sleep(2)
  # 台区
    driver.find_element_by_xpath('//div/div[2]/div[2]/div/div/div[2]/div/div[2]/div/table/tbody/tr[5]/td/div/input').click()
    time.sleep(2)
    # 点击月数据
    driver.find_element_by_id('district400monCalcData-inputEl').click()
    time.sleep(2)
  # 点击计算按钮
    driver.find_element_by_id('district400_calc_start_button-btnInnerEl').click()
    time.sleep(20)
  # 计算成功后弹出框信息ID以及取消按钮ID
    driver.find_element_by_id('messagebox-1001-displayfield-inputEl').click()
    time.sleep(5)
    driver.find_element_by_id('button-1006-btnInnerEl') .click()
    time.sleep(5)

except Exception, e:
    print e
finally:
    time.sleep(5)
    driver.quit()