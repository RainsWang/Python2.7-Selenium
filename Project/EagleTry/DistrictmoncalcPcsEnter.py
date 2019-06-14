#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：王敏
IDE           ：PyCharm
时间          ：2018-12-27 15:08
当前项目名称  ：Project
功能          ：运行数据录入
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
    # 运行数据录入enter
    # 数据中心id
    driver.find_element_by_id('Dacmenu-panel_header_hd').click()
    time.sleep(2)
    # 运营数据管理路径
    driver.find_element_by_xpath('//div[@id="Dacmenu-panel-body"]/div/table/tbody/tr[3]/td/div/img[1]').click()
    time.sleep(2)
    # 台区线路运行数据录入
    driver.find_element_by_xpath('//div[@id="Dacmenu-panel-body"]/div/table/tbody/tr[6]/td/div/img[2]').click()
    time.sleep(2)
    # 台区理论线损运行数据录入
    driver.find_element_by_xpath('//div[@id="Dacmenu-panel-body"]/div/table/tbody/tr[8]/td/div').click()
    time.sleep(2)
    # 台区理论线损运行数据录入查询页面frame
    driver.switch_to.frame('iframe_120001140502')
    # 查询条件之量测分组选择，先点击下拉框，再选择月量测类型分组
    driver.find_element_by_id('ext-gen1507').click()
    time.sleep(2)
    driver.find_element_by_xpath('//div[@id="boundlist-1829"]/div/ul/li[1]').click()
    time.sleep(5)

    # 月理论运行数据量测类型分组

    # 录入台区运行数据信息
    # 正向有功电量
    driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr[2]/td[3]/div').click()
    time.sleep(2)
    # 激活输入框
    # driver.find_element_by_name('PQ_P_A').click()
    driver.find_element_by_name('PQ_P_A').clear()
    driver.find_element_by_name('PQ_P_A').send_keys('100')
    time.sleep(2)
    # 功率因素录入
    driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr[2]/td[59]/div').click()
    # 激活输入框
    time.sleep(2)
    # driver.find_element_by_name('F')
    #  功率因素录入
    driver.find_element_by_name('F').clear()
    driver.find_element_by_name('F').send_keys('0.2')
    time.sleep(2)
    # 负荷形状系数录入
    driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr[2]/td[72]/div').click()
    time.sleep(2)
    # 激活输入框
    # driver.find_element_by_name('K').click()
    driver.find_element_by_name('K').clear()
    driver.find_element_by_name('K').send_keys('1')
    time.sleep(2)
    # 输入台区用户数据
    # 切换到台区用户tab页面
    driver.find_element_by_id('tab-1171-btnInnerEl').click()
    time.sleep(5)
    # 台区用户
    driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr[2]/td[3]/div').click()
    time.sleep(2)
    # 台区用户数据录入
    # driver.find_element_by_name('PQ_P_A').click()
    driver.find_element_by_name('PQ_P_A').clear()
    driver.find_element_by_name('PQ_P_A').send_keys('40')
    time.sleep(2)
    driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr[3]/td[3]/div').click()
    time.sleep(2)
    # driver.find_element_by_name('PQ_P_A').click()
    driver.find_element_by_name('PQ_P_A').clear()
    driver.find_element_by_name('PQ_P_A').send_keys('10')
    time.sleep(2)
    driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr[4]/td[3]/div').click()
    time.sleep(2)
    # driver.find_element_by_name('PQ_P_A').click()
    driver.find_element_by_name('PQ_P_A').clear()
    driver.find_element_by_name('PQ_P_A').send_keys('30')
    time.sleep(2)
    driver.find_element_by_xpath('//div[2]/div[2]/div/table/tbody/tr[5]/td[3]/div').click()
    time.sleep(2)
    # driver.find_element_by_name('PQ_P_A').click()
    driver.find_element_by_name('PQ_P_A').clear()
    driver.find_element_by_name('PQ_P_A').send_keys('20')
    time.sleep(2)
    # 输入数据后进行保存,以及弹出框内容校验
    driver.find_element_by_id('save_meas_button-btnInnerEl').click()
    time.sleep(2)
    # 保存按钮
    driver.find_element_by_id('messagebox-1001-displayfield-inputEl').click()
    time.sleep(2)
    # 弹出框信息
    driver.find_element_by_id('button-1005-btnInnerEl').click()
    time.sleep(8)
    # 确定按钮
except Exception, e:
    print e
finally:
    time.sleep(5)
    driver.quit()