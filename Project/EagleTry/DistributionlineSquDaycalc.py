#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018-9-27 21:01
当前项目名称  ：Auto
功能          ：6-20kv均方根电流法计算
-------------------------------------漂亮的分割线----------------------------------------------'''
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import re
try:
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

    #点击理论线损
    driver.find_element_by_xpath("//div[3]/div/div/div/div/div[2]/img").click()
    time.sleep(1)
    '''新增配线日均方根电流法任务插件，仅为了练习'''
    #点击计算任务设置
    driver.find_element_by_xpath("//div[3]/div[3]/div/table/tbody/tr[3]/td/div/img").click()
    time.sleep(1)
    #点击自动计算任务设置
    '''driver.find_element_by_xpath("//div[3]/div[3]/div/table/tbody/tr[4]/td/div/img[3]").click()
    time.sleep(1)
    #切换到计算任务frame
    driver.switch_to_frame('iframe_13000102')
    #点击新增按钮
    driver.find_element_by_id("add_task_id-btnWrap").click()
    time.sleep(1)
    #点击配电线路任务
    driver.find_element_by_xpath("//div[6]/div/div[2]/div[2]/div[2]/a/span").click()
    time.sleep(1)
    #输入日均方根电流法计算任务名称
    driver.find_element_by_id("textfield-1063-inputEl").send_keys(u"配线日均方根电量法计算")
    time.sleep(2)
    #点击计算方法选择框
    driver.find_element_by_xpath("//div/div/div/div/fieldset/div/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div").click()
    time.sleep(2)
    #点击均方根电流法
    driver.find_element_by_xpath("//div[10]/div/ul/li[3]").click()
    time.sleep(2)
    #取消是否启动勾选
    driver.find_element_by_id("checkboxfield-1050-inputEl").click()
    time.sleep(2)
    #勾选配线全选框
    driver.find_element_by_xpath("//div[2]/div/div/fieldset/div/div/div[2]/div/table/tbody/tr[2]/td/div/input").click()
    time.sleep(2)
    #点击确定按钮
    driver.find_element_by_id("button-1074-btnInnerEl").click()
    time.sleep(2)'''
    #点击任务执行情况查询
    driver.find_element_by_xpath("//div[3]/div[3]/div/table/tbody/tr[5]/td/div/img[3]").click()
    time.sleep(2)
    #切换到任务设置frame
    driver.switch_to_frame('iframe_13000104')
    #点击所有任务前面的加号
    driver.find_element_by_xpath("//div/div[3]/div/table/tbody/tr[2]/td/div/div").click()
    time.sleep(2)
    #选中配线日均方根电流法任务
    driver.find_element_by_xpath("//div/div[3]/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/div").click()
    time.sleep(2)
    #点击明细数据
    driver.find_element_by_id("button-1026-btnInnerEl").click()
    time.sleep(2)
    #跳出当面frame
    driver.switch_to.parent_frame()## 跳出当前iframe
    driver.implicitly_wait(10)
    #切换到日均方根电流法任务执行情况明细frame
    driver.switch_to_frame('iframe_5685F266AA6D40BEAADCB10D5AFA67AF')
    time.sleep(2)
    #点击日期选择框
    driver.find_element_by_xpath("//div[2]/div/div/div/table/tbody[2]/tr/td[2]/table/tbody/tr/td[2]/div").click()
    time.sleep(2)
    #选择2014-8-1日
    #点击日期选择框
    driver.find_element_by_id("splitbutton-1056-btnWrap").click()
    time.sleep(2)
    #选择2014年
    driver.find_element_by_xpath("//div[7]/div[2]/div/div[2]/div[2]/a").click()
    time.sleep(2)
    #选择8月
    driver.find_element_by_xpath("//div[7]/div[2]/div/div/div[4]/a").click()
    time.sleep(2)
    #点击确定
    driver.find_element_by_id("button-1059-btnInnerEl").click()
    time.sleep(2)
    #选择1号
    driver.find_element_by_xpath("//div[7]/div/table/tbody/tr/td[6]/a/em/span").click()
    time.sleep(2)
    #点击查询
    driver.find_element_by_id("button-1046-btnInnerEl").click()
    time.sleep(2)
    #选择任务
    driver.find_element_by_xpath("//div[2]/div/div/div[3]/div/table/tbody/tr[2]/td/div/div").click()
    time.sleep(2)
    #点击补算
    driver.find_element_by_id("button-1047-btnInnerEl").click()
    time.sleep(2)
    #切换到补算框中frame
    CurrentPageSource = driver.page_source
    FrameDynamicID = re.findall("iframe id=\"eagleWindow(.*?)Iframe",CurrentPageSource)
    Frame_id = "eagleWindow"+ FrameDynamicID+"Iframe"
    driver.switch_to_frame(Frame_id)
    #点击弹出框中补算
    driver.find_element_by_id("retry_button_id-btnEl").click()
    time.sleep(15)
    alerttext=driver.find_element_by_id("ext-gen1055").text
    if alerttext==u'计算完成':
        print "计算成功！"
    else:
        print "无法正常计算"
    time.sleep(20)
except Exception,e:
    print (e)
finally:
    time.sleep(10)
    driver.close()