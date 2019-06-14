#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
作者          ：李艳丽
时间          ：2019/3/5 15:11
项目和文件名称：PyCharm Community Edition Pandas.py
脚本功能说明  ：利用pandas读取excel中的数据
-------------------------------------漂亮的分割线----------------------------------------------'''
import xlrd
from Public_Theory.GetUsername import GetUsername
import pandas as pd
import operator
import numpy
from pandas import Series,DataFrame

class ExcelData:
    def __init__(self,data_path):
        self.data_path=data_path
    def readExcel(self):
        book = xlrd.open_workbook(self.data_path,encoding_override='gbk')
        L=[]
        for sheet in book.sheets():
            df = pd.read_excel(self.data_path,sheet.name,index_col = None,na_values= ['9999'])
            #df = pd.read_excel(self.data_path,sheet.name,index_col = 0)
            pd.set_option('display.width',None)
            #pd.set_option('expand_frame_repr',False)
            df1=df.head()
            df2 = numpy.array(df1).tolist()
            L.append(df2)
        return L

if __name__=='__main__':
    data_path=u'C:\\Users\\'+GetUsername()+u'\\Downloads\\低压台区综合分析报表(月).xls'
    get_data=ExcelData(data_path)
    print(get_data.readExcel())
    data_path1 = u'C:\\Users\\'+GetUsername()+u'\\Downloads\\低压台区综合分析报表(月) (1).xls'
    get_data1 = ExcelData(data_path1)
    print(get_data1.readExcel())
    result = operator.eq(get_data,get_data1)
    print(result)
