# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/10/19 0019 上午 10:06
当前项目名称  ：Auto
功能          ：配电线路日容量法计算，计算日期为2014-8-1
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
    # 点击理论线损计算
    driver.find_element_by_xpath("//div[@id ='Lltmenu-panel-body']/div/table/tbody/tr[2]/td/div").click()
    # 切换到理论线损手工计算页面620计算的tab
    driver.switch_to.frame('iframe_13000110')
    driver.implicitly_wait(5)
    driver.find_element_by_css_selector('#tab-1154-btnWrap').click()
    # 勾选地区
    time.sleep(2)
    driver.find_element_by_xpath("//div[@id = 'district620_grid_tree-body']/div/table/tbody/tr[2]/td/div/input").click()

    # 选择计算方法为容量法
    driver.find_element_by_xpath("//tr[@id = 'district620calcMethodGroup-inputRow']/td[2]/table/tbody/tr/td/table[2]/tbody/tr/td[2]/input").click()
    #driver.find_element_by_css_selector("tr# district620calcMethodGroup-inputRow>td[1]>table>tbody>tr>td>table[1]>tbody>tr>td[1]>input").click()
    #选择计算类型为日计算
    time.sleep(1)
    driver.find_element_by_css_selector("#district620dayCalcData-inputEl").click()
    #选择日期为2014-8-1
    time.sleep(1)
    driver.find_element_by_css_selector("#dataTime620-inputEl").clear()
    driver.find_element_by_css_selector("#dataTime620-inputEl").send_keys("2014-08-01")  #日期输入框中直接输入2014-8-1
    time.sleep(1)

    # 点击计算
    driver.find_element_by_css_selector('#district620_calc_start_button').click()
    driver.implicitly_wait(60)
    alerttext = driver.find_element_by_id("messagebox-1001-displayfield-inputEl").text
    print alerttext
    if u'是否要打开当前配线的计算结果？' == alerttext:
        print "Pass"
    else:
        print "False"
except Exception,e:
    raise e
finally:
    time.sleep(10)
    driver.quit()