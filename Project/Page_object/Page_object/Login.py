#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/8 16:43
当前项目名称  ：Login
功能          ：登录封装
-------------------------------------漂亮的分割线----------------------------------------------'''
from selenium import webdriver
import  time
import yaml
from Public_Theory.GetProjectFilePath import GetProjectFilePath
from selenium.webdriver.common.keys import Keys

class Login():
    '''登录'''
    def __init__(self,driver):
        self.ProjectFilePath=GetProjectFilePath()
        #self.driver=webdriver.Chrome()
        self.driver=driver
        #获取页面元素
        self.Page_object_data_file=open(self.ProjectFilePath+"\Page_object\Data\Login.yaml")
        self.Page_data=yaml.load(self.Page_object_data_file)
        self.Page_object_data_file.close()
        #将界面元素信息赋给相对应的变量
        self.Login_url=self.Page_data['Login']['url']
        self.username=self.Page_data['Login']['username_id']
        self.password=self.Page_data['Login']['password_id']
        self.loginbutton_class_name=self.Page_data['Login']['loginbutton_class_name']
    def Login_Function(self,username,password):
        '''登录'''
        self.driver.get(self.Login_url)
        self.driver.maximize_window()
        #输入数据
        self.driver.find_element_by_id(self.username).clear()
        self.driver.find_element_by_id(self.username).send_keys(username)
        self.driver.find_element_by_id(self.username).send_keys(Keys.TAB)
        self.driver.find_element_by_id(self.password).clear()
        self.driver.find_element_by_id(self.password).send_keys(password)
        self.driver.find_element_by_class_name(self.loginbutton_class_name).click()
        time.sleep(5)
        '''if ('http://192.168.0.141:3030/eagle2/coc/application/controller/frame/LoginUI.action'==self.driver.current_url):
            title = '未登录成功'
        else:
            title=self.driver.find_element_by_link_text(username).text'''
        title=self.driver.find_element_by_link_text(username).text
        return title
if __name__ == "__main__":
    driver = webdriver.Chrome()
    Try = Login(driver)
    Try.Login_Function('lyltest','000000')
    driver.close()


