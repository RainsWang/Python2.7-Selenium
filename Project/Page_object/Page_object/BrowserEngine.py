# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：Administrator
IDE           ：PyCharm
时间          ：2018/9/19 0019 上午 11:04
当前项目名称  ：Auto
功能          ：浏览器引擎的封装类
-------------------------------------漂亮的分割线----------------------------------------------
'''
import ConfigParser
from selenium import webdriver
from Public_Theory.GetProjectFilePath import GetProjectFilePath
class BrowserEngine(object):
    def __init__(self):
        self.config = ConfigParser.ConfigParser()
        self.config_path = GetProjectFilePath() + "\\Page_object\\Data\\Config.ini"
        self.config.read(self.config_path)
        self.browser = self.config.get("browserType", "browserName")
        #print self.browser
    def GetDriver(self):
        if self.browser == "Firefox":
            self.driver = webdriver.Firefox()
        elif self.browser == "Chrome":
            self.driver = webdriver.Chrome()
        elif self.browser == "IE":
            self.driver = webdriver.Ie()
        return self.driver

if __name__ == '__main__':
    browser = BrowserEngine()
    driver = browser.GetDriver()
    driver.get('http://www.baidu.com')
    driver.implicitly_wait(10)
    driver.quit()
