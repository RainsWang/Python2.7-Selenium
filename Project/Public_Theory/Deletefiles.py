#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018-8-1 14:34
当前项目名称  ：620Calc
功能          ：case运行之前先删除“u'C:\\Users\\Administrator\\Downloads'”中的文件，以便计算结果导出的excel名称的
正确性
-------------------------------------漂亮的分割线----------------------------------------------'''
import os
import getpass
#curr_path=u'C:\\Users\\Administrator\\Downloads'
#print path
#path=u'C:\\Users\\Administrator\\Downloads'
def del_file():
    path=os.path.join(u'C:\\Users\\',getpass.getuser(),'Downloads')
    ls=os.listdir(path)#获得指定目录中的内容
    for i in ls:
        c_path=os.path.join(path,i)#将多个路径组合后返回
        if os.path.isdir(c_path):#判断某一路径是否为目录
            del_file(c_path)
        else:
            os.remove(c_path)#删除指定路径的文件
        #print c_path
if __name__=='__main__':
    del_file()
