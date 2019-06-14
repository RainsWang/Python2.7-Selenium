#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：吴鹏飞
IDE           ：PyCharm Community Edition
时间          ：2019/1/10 14:23
当前项目名称  ：Project
功能          ：
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
try:
    #登录
    driver=webdriver.Chrome()
    time.sleep(1)
    driver.get('http://192.168.0.11:3030/eagle2/coc/application/controller/frame/LoginUI.action')
    time.sleep(1)
    driver.maximize_window()
    time.sleep(1)
    driver.find_element_by_css_selector('#username').clear()
    driver.find_element_by_css_selector('#username').send_keys('lj')
    driver.find_element_by_css_selector('#username').send_keys(Keys.TAB)
    driver.find_element_by_css_selector('#password').clear()
    driver.find_element_by_css_selector('#password').send_keys('000000')
    driver.find_element_by_css_selector('.loginbutton').click()
    time.sleep(3)
    print "登录成功！"
    #进入运行数据录入界面录入运行数据
    driver.find_element_by_css_selector('#Dacmenu-panel_header_hd-textEl').click()
    time.sleep(1)
    driver.find_element_by_css_selector('div#treeview-1017>table>tbody>tr:nth-child(3)>td>div>img').click()
    time.sleep(1)
    driver.find_element_by_css_selector('div#treeview-1017>table>tbody>tr:nth-child(4)>td>div>img:nth-child(2)').click()
    time.sleep(1)
    #进入理论运行数据录入界面，
    driver.find_element_by_css_selector('div#treeview-1017>table>tbody>tr:nth-child(6)>td>div').click()
    time.sleep(3)
    driver.switch_to_frame('iframe_120001140302')
    time.sleep(3)
    driver.find_element_by_css_selector('div#ext-gen1797').click()
    time.sleep(1)
    driver.find_element_by_css_selector('div#boundlist-2033-listEl>ul>li').click()
    time.sleep(1)
    driver.find_element_by_css_selector('div#ext-gen1815').click()
    time.sleep(1)
    driver.switch_to_frame('ext-gen1039')
    time.sleep(2)
    driver.find_element_by_css_selector('td#ext-gen1055>a').click()
    time.sleep(1)
    driver.find_element_by_css_selector('#button-1023-btnInnerEl').click()
    #点击查询
    #driver.find_element_by_css_selector('#button-1014-btnInnerEl').click()
    time.sleep(3)
    driver.switch_to.parent_frame()
    time.sleep(2)
    driver.find_element_by_css_selector('#tab-1315-btnInnerEl').click()
    time.sleep(3)
    # driver.find_element_by_css_selector('td.x-grid-cell.x-grid-cell-MeasurementColumn-1158>div').click()
    # time.sleep(1)
    # driver.find_element_by_css_selector("[name='P']").clear()
    # time.sleep(1)
    # driver.find_element_by_css_selector("[name='P']").send_keys('0.0694')
    ActivePower=[0.0694,0.0621,0.0621,0.0658,0.0585,0.084,0.106,0.676,0.7454,1.0634,0.6322,0.1315,0.095,0.6395,0.5372,0.5372,0.5956,0.5152,0.1206,0.2521,0.2631,0.1754,0.0731,0.0658]
    ReactivePower=[0.1023,0.106,0.106,0.1096,0.095,0.1206,0.1206,0.983,1.1145,1.1035,0.9757,0.19,0.1571,1.0487,0.9208,1.0305,0.9976,0.9135,0.1937,0.4421,0.4348,0.3143,0.0987,0.1023]
    Voltage=[111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5,111.5]
    TransportationCapacity=[0.43,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,0.65,0.36,0.58,100]
    # 录入负荷运行数据
    a='//div[2]/div[2]/div/table/tbody/tr['
    b=']/td[49]/div'
    for i in range(2,26):
        driver.find_element_by_xpath(a+str(i)+b).click()
        time.sleep(2)
        driver.find_element_by_css_selector("[name='P']").clear()
        time.sleep(2)
        driver.find_element_by_css_selector("[name='P']").send_keys(str(ActivePower[i-2]))
        time.sleep(2)
        driver.find_element_by_css_selector("[name='P']").send_keys(Keys.TAB)
        time.sleep(2)
        driver.find_element_by_css_selector("[name='Q']").clear()
        time.sleep(2)
        driver.find_element_by_css_selector("[name='Q']").send_keys(str(ReactivePower[i-2]))
        time.sleep(3)
        driver.find_element_by_css_selector("[name='Q']").send_keys(Keys.TAB)
        time.sleep(3)
    driver.find_element_by_css_selector('#button-1015-btnInnerEl').click()
    time.sleep(2)
    driver.find_element_by_css_selector('#button-1005-btnInnerEl').click()
    time.sleep(2)
    driver.find_element_by_css_selector('#tab-1319-btnInnerEl').click()
    time.sleep(2)
    c='//div[2]/div[2]/div/table/tbody/tr['
    d=']/td[61]/div'
    for i in range(2,26):
        driver.find_element_by_xpath(c+str(i)+d).click()
        time.sleep(2)
        driver.find_element_by_css_selector("[name='V']").clear()
        time.sleep(2)
        driver.find_element_by_css_selector("[name='V']").send_keys(str(Voltage[i-2]))
        time.sleep(2)
        driver.find_element_by_css_selector("[name='V']").send_keys(Keys.TAB)
        time.sleep(2)
    driver.find_element_by_css_selector('#button-1015-btnInnerEl').click()
    time.sleep(1)
    driver.find_element_by_css_selector('#button-1005-btnInnerEl').click()
    time.sleep(2)
    driver.find_element_by_css_selector('#tab-1321-btnInnerEl').click()
    time.sleep(2)
    e='//div[2]/div[2]/div/table/tbody/tr['
    f=']/td[74]/div'
    for i in range(2,26):
        driver.find_element_by_xpath(e+str(i)+f).click()
        time.sleep(2)
        driver.find_element_by_css_selector("[name='KVA']").clear()
        time.sleep(2)
        driver.find_element_by_css_selector("[name='KVA']").send_keys(str(TransportationCapacity[i-2]))
        time.sleep(2)
        driver.find_element_by_css_selector("[name='KVA']").send_keys(Keys.TAB)
        time.sleep(2)
    driver.find_element_by_css_selector('#button-1015-btnInnerEl').click()
    time.sleep(1)
    driver.find_element_by_css_selector('#button-1005-btnInnerEl').click()
    time.sleep(3)
    driver.switch_to_default_content()
    time.sleep(2)
    print "运行数据录入成功！"
    #进入理论线损计算界面进行计算
    driver.find_element_by_css_selector('#Lltmenu-panel_header_hd-textEl').click()
    time.sleep(2)
    #driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(2)>td>div>img:nth-child(2)')
    driver.find_element_by_css_selector('td.x-grid-cell-treecolumn.x-grid-cell.x-grid-cell-treecolumn-1019.x-grid-cell-first>div.x-grid-cell-inner>img:nth-child(2)').click()
    time.sleep(2)
    driver.switch_to_frame('iframe_13000110')
    time.sleep(2)
    driver.find_element_by_css_selector('.x-tree-checkbox').click()
    time.sleep(2)
    driver.find_element_by_css_selector('#calc_start_button-btnInnerEl').click()
    driver.implicitly_wait(60)
    driver.find_element_by_css_selector('#button-1007-btnInnerEl').click()
    time.sleep(2)
    driver.switch_to_default_content()
    time.sleep(2)
    print "计算成功！"
    #查看输电网整点其他损耗表！
    # driver.find_element_by_css_selector('#Lltmenu-panel_header_hd-textEl').click()
    # time.sleep(2)
    driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(4)>td>div>img').click()
    time.sleep(2)
    driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(7)>td>div>img:nth-child(2)').click()
    time.sleep(2)
    driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(16)>td>div').click()
    driver.implicitly_wait(60)
    driver.switch_to_frame('iframe_4028801b3ebf9003013ec0b37ab40f15')
    time.sleep(2)
    driver.find_element_by_css_selector('div.x-trigger-index-0.x-form-trigger.x-form-date-trigger.x-form-trigger-last.x-unselectable').click()
    time.sleep(2)
    driver.find_element_by_css_selector('table#datetimepicker-1095-eventEl>tbody>tr>td:nth-child(6)>a>em>span').click()
    time.sleep(1)
    driver.find_element_by_css_selector('#numberfield-1096-inputEl').clear()
    time.sleep(1)
    driver.find_element_by_css_selector('#numberfield-1096-inputEl').send_keys('0')
    time.sleep(1)
    driver.find_element_by_css_selector('#button-1100-btnInnerEl').click()
    time.sleep(3)
    driver.find_element_by_id('4028801b3ea5f3cd013ea6c8a15f05b4-btnInnerEl').click()
    time.sleep(3)
    driver.find_element_by_id('4028801b3ea5f3cd013ea6c8a15d05a9-btnInnerEl').click()
    time.sleep(5)
    driver.switch_to_default_content()
    time.sleep(2)
    print "查询并导出其他损耗表整点报表!"
    #查看输电网其他损耗表日报表！
    # driver.find_element_by_css_selector('#Lltmenu-panel_header_hd-textEl').click()
    # time.sleep(2)
    # driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(4)>td>div>img').click()
    # time.sleep(2)
    driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(6)>td>div>img:nth-child(2)').click()
    time.sleep(2)
    driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(12)>td>div').click()
    time.sleep(3)
    driver.switch_to_frame('iframe_4028801b3e3fabd9013e3fbb26d100e2')
    time.sleep(2)
    driver.find_element_by_id('4028801b3e2f6b33013e30ad20020161-btnInnerEl').click()
    time.sleep(3)
    driver.find_element_by_id('4028801b3e2f6b33013e30ad1ffe0155-btnInnerEl').click()
    time.sleep(5)
    driver.switch_to_default_content()
    time.sleep(2)
    print "查询并导出其他损耗表日报表！"
    #查看输电网其他损耗表月报表！
    # driver.find_element_by_css_selector('#Lltmenu-panel_header_hd-textEl').click()
    # time.sleep(2)
    # driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(4)>td>div>img').click()
    # time.sleep(2)
    driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(5)>td>div>img:nth-child(2)').click()
    time.sleep(2)
    driver.find_element_by_css_selector('div#Lltmenu-panel-body>div>table>tbody>tr:nth-child(11)>td>div').click()
    time.sleep(3)
    driver.switch_to_frame('iframe_4028801b3ebf9003013ec0b04f650f00')
    time.sleep(2)
    driver.find_element_by_id('4028801b3e2f6b33013e30ad4fad0172-btnInnerEl').click()
    time.sleep(3)
    driver.find_element_by_id('4028801b3e2f6b33013e30ad4fa90166-btnInnerEl').click()
    time.sleep(5)
    print "查询并导出其他损耗表月报表！"
except Exception,e:
    print e
finally:
    time.sleep(3)
    driver.quit()

