#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
import xlrd
import xlwt

style = xlwt.XFStyle()
font = xlwt.Font()
font.name = 'Calibri'
font.height = 20*11
style.font = font
	
surname = []
name = []
patronymic = []
index = []
src = xlrd.open_workbook("кандидаты завод.xlsx")
sheet = src.sheet_by_index(0)
#col = sheet.col_values(1)
wb = xlwt.Workbook()
ws = wb.add_sheet('1')
for row in range(1,sheet.nrows):
    index.append(sheet.cell(row,3).value.encode("cp1251"))
for i in range(0,len(index)):
    index[i] = index[i].decode("cp1251").lstrip()
    index[i] = index[i].replace('-',' ')
    ws.write(i,0,index[i])
    surname.append(index[i].split(' ')[0])
    name.append(index[i].split(' ')[1])
    patronymic.append(index[i].split(' ')[2])
print 'Surname', len(surname)
print 'Name', len(name)
print 'Patromymic', len(patronymic)
print 'Index', len(index)
for i in range(len(surname)):
    ws.write(i,1,surname[i])
for i in range(len(name)):
    ws.write(i,2,name[i])
for i in range(len(patronymic)):
    ws.write(i,3,patronymic[i])

'''for i in range(0,len(patronymic)):
    try:
        print patronymic[i]
    except IndexError:
        print i
        print surname[i]
print index[i].split()[1]
    if index[i].split(' ')[1].startswith(""): index[i].split(' ')[1].lstrip()
    ws.write(i+1,2,index[i].split(' ')[1])
    try:
        ws.write(i+1,3,index[i].split(" ")[2])
        #print (index[i].split(" ")[2].decode("cp1251")).split("-")[0]
    except IndexError:
        ws.write(i+1,3,index[i].split("-")[1])
    #print index[i].decode("cp1251")
#print index[6].split(" ")[1].decode("cp1251")'''
wb.save('dest.xls')