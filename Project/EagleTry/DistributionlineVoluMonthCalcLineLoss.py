# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/9/28 0028 下午 3:40
当前项目名称  ：Auto
功能          ：月容量法计算线路损耗结果核对
-------------------------------------漂亮的分割线----------------------------------------------
'''
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
try:
    driver = webdriver.Chrome()
    driver.get('http://192.168.0.141:3030/eagle2/coc/application/controller/frame/LoginUI.action')
    driver.maximize_window()
    driver.find_element_by_id('username').clear()
    driver.find_element_by_id('username').send_keys('Autotest')
    driver.find_element_by_id('username').send_keys(Keys.TAB)
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("password").send_keys('000000')
    driver.find_element_by_class_name('loginbutton').click()
    driver.implicitly_wait(10)

    # 点击理论线损树
    driver.find_element_by_id('Lltmenu-panel_header_hd').click()
    #点击配电网计算结果查询
    driver.find_element_by_xpath("//*[@id = 'Lltmenu-panel-body']/div/table/tbody/tr[5]/td/div/img[1]").click()
    time.sleep(1)
    #点击月报表
    driver.find_element_by_xpath("//*[@id = 'Lltmenu-panel-body']/div/table/tbody/tr[6]/td/div/img[2]").click()
    time.sleep(1)
    #点击配电线路线损表
    driver.find_element_by_xpath("//*[@id = 'Lltmenu-panel-body']/div/table/tbody/tr[7]/td/div").click()
    time.sleep(1)
    #切换frame到配电线路线损表frame
    driver.switch_to.frame("iframe_297efe173ab9e771013aba8af15d000a")
    #点击线路损耗链接
    driver.find_element_by_xpath("//table/tbody/tr[3]/td/table/tbody/tr[2]/td[9]/div/a/div").click()
    time.sleep(2)
    driver.switch_to.parent_frame()
    driver.switch_to.frame("iframe_297efe173ab9e7bb013aba7334df0008")
    driver.implicitly_wait(5)
    driver.find_element_by_id("button-1055-btnIconEl").click()


except Exception,e:
    print e
finally:
    time.sleep(10)
    driver.quit()