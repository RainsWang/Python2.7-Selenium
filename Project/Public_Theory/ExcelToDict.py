#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''-------------------------------------漂亮的分割线----------------------------------------------
脚本功能说明  ：
作者          ：李艳丽
IDE           ：PyCharm Community Edition
时间          ：2018/7/26 9:32
当前项目名称  ：620Calc
功能          ：读取excel中的数据
-------------------------------------漂亮的分割线----------------------------------------------'''
import xlrd
import json
from Public_Theory.GetUsername import GetUsername
class ExcelData:
    def __init__(self,data_path,sheetname):
        self.data_path=data_path
        #excellujing
        self.sheetname=sheetname
        #excel的sheet名
        self.data=xlrd.open_workbook(self.data_path)
        #打开excel
        #切换到相应的sheet
        self.table=self.data.sheet_by_name(self.sheetname)
        #第二行作为key值
        self.key=self.table.row_values(1)
        #获取表格的行数
        self.rowNum=self.table.nrows
        #获取表格的列数
        self.colNum=self.table.ncols
    def readExcel(self):
        if self.rowNum<=2:
            print("表数据为空！！")
        else:
            L=[]
            for i in range(2,self.rowNum-1):
                sheet_data={}
                for j in range(1,self.colNum):
                    sheet_data[self.key[j]]=self.table.row_values(i)[j]
                L.append(sheet_data)
            return L
            #return json.dumps(L, encoding="UTF-8", ensure_ascii=False)
        #encode('gbk').decode('gbk').encode('utf8')
if __name__=='__main__':
    data_path=u'C:\\Users\\'+GetUsername()+u'\\Downloads\\低压台区综合分析报表(月).xls'
    sheetname=u'低压台区综合分析报表(月)'
    get_data=ExcelData(data_path,sheetname)
    #print(get_data.readExcel())
    print(json.dumps(get_data.readExcel(),  ensure_ascii=False))

