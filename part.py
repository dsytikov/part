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
    if index[i].decode("cp1251") == '': print "Empty cell ", i
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
wb.save('dest.xls')