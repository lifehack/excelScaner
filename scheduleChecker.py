__author__ = 'anjing'
#coding=utf-8

import xlrd

file = raw_input("输入文件路径：")

xls = xlrd.open_workbook(file)

while True:
    target = raw_input(u"输入课程名：").decode('utf-8')

    for x in range(xls.nsheets):
        sh = xls.sheets()[x]

        for r in range(sh.nrows):
            for c in range(sh.ncols):
                value = sh.cell_value(r, c)
                if target in unicode(value):
                    print sh.name.encode('utf-8'), target, sh.cell_value(r,12), sh.cell_value(r,11), sh.cell_value(r,4), sh.cell_value(r, 8), sh.cell_value(r,13)