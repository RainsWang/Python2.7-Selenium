#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：刘娟
IDE           ：PyCharm Community Edition
时间          ：2018/12/18 11:18
当前项目名称  ：Auto
功能          ：配电线路潮流精确算法日计算
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait

try:
    #登陆
    driver=webdriver.Chrome()
    driver.get('http://192.168.0.11:3030/eagle2')
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
    #点击理论线损手工计算
    driver.find_element_by_xpath('//div[3]/div[3]/div/table/tbody/tr[2]/td/div').click()
    time.sleep(2)
    #进入手工计算页面后，切换到配电网6-20kV理论线损计算tab
    driver.switch_to_frame('iframe_13000110')
    driver.find_element_by_id('tab-1162-btnInnerEl').click()
    time.sleep(2)
    #点击单位、变电站、配线
    driver.find_element_by_xpath('//div/div/div[3]/div/div[2]/div/table/tbody/tr[2]/td/div/img').click()
    driver.find_element_by_xpath('//div/div/div[3]/div/div[2]/div/table/tbody/tr[3]/td/div/img[2]').click()
    driver.find_element_by_xpath('//div/div/div[3]/div/div[2]/div/table/tbody/tr[4]/td/div/input').click()
    time.sleep(3)
    #计算方法中选择潮流精确算法
    driver.find_element_by_id('district620calcMethodTM-inputEl').click()
    time.sleep(1)
    #点击计算
    driver.find_element_by_id('district620_calc_start_button-btnInnerEl').click()
    driver.implicitly_wait(120)

    alerttext=driver.find_element_by_id('messagebox-1001-displayfield-inputEl').text
    if alerttext==u'是否要打开当前配线的计算结果？':
        driver.find_element_by_id('button-1006-btnInnerEl').click()
        print"计算成功，跳转到相应报表页面"
    else:
        print"无法正常计算"
        time.sleep(20)

except Exception,e:
    print(e)
finally:
    time.sleep(10)
    driver.close()










