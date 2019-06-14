#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/8 16:49
当前项目名称  ：获取当前根目录
功能          ：获取到当前路径
-------------------------------------漂亮的分割线----------------------------------------------'''
import os
def GetProjectFilePath():
    FilePath = os.path.dirname(__file__)
    ProjectFilePath  = os.path.dirname(FilePath)
    return ProjectFilePath

if __name__ == '__main__':
    print(GetProjectFilePath())