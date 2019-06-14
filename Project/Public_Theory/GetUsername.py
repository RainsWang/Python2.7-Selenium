# -*- coding: utf-8 -*-
'''
-------------------------------------漂亮的分割线----------------------------------------------
作者          ：王禹
IDE           ：PyCharm
时间          ：2018/9/29 0029 下午 3:19
当前项目名称  ：Auto
功能          ：获取当前系统运行登录的用户
-------------------------------------漂亮的分割线----------------------------------------------
'''

import getpass

def GetUsername():
    return getpass.getuser()

if __name__ == "__main__":
    print(GetUsername())