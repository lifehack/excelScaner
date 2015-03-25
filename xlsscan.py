from xlrd import *

import os
import sys
from tkinter import *

def scan(valuelist, path, key, log):
    if not path:
        return

    if not key:
        return

    result = list()

    def findrc(sh, i):
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                v = str(sh.cell_value(r, c)).replace(" ","")
                if i==v:
                    return (r,c)

    for f in os.listdir(path):

        ufile = f

        p = path + ufile

        if os.path.isdir(p):
            result.extend(scan(valuelist, p+"/", key, log))
            continue

        if not (".xls" in f or ".xlsx" in f):
            continue

        try:
            xls = open_workbook(p)
        except:
            continue

        sh = xls.sheets()[0]

        rtemp = findrc(sh, key)
        if not rtemp:
            continue

        row = rtemp[0]

        tmp = dict()

        for vl in valuelist:
            tmp[vl] = findrc(sh, vl)

        for r in tmp.keys():
            if not tmp[r]:
                continue

            tmp[r] = str( sh.cell_value(row, tmp[r][1]) )

            log.config(state=NORMAL)
            log.insert(END, tmp[r])
            log.insert(END, ",")
            log.config(state=DISABLED)

        result.append(tmp)

        log.config(state=NORMAL)
        log.insert(END, "\n")
        log.config(state=DISABLED)

    return result

def check(path, log):
    if not path:
        return

    def findrc(sh, i):
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                v = unicode(sh.cell_value(r, c)).replace(" ","")
                if cmp(i,v)==0:
                    return (r,c)

    for f in os.listdir(path):

        if not isinstance(f,unicode):
            ufile = f.decode('gbk').encode('utf-8')
        else:
            ufile = f

        p = unicode(path + ufile)

        if os.path.isdir(p):
            check(p+"/", log)
            continue

        if not (".xls" in f or ".xlsx" in f):
            continue

        try:
            xls = open_workbook(p)
        except:
            continue

        if xls.nsheets>1:
            for i in range(1,xls.nsheets):
                log.config(state=NORMAL)
                log.insert(END, "文件%s: 发现额外表%s\n" % (p,xls.sheets()[i].name))
                log.config(state=DISABLED)

        sh = xls.sheets()[0]

        colindex1 = findrc(sh, "课程名称")
        colindex2 = findrc(sh, "考试成绩")
        colindex3 = findrc(sh, "学号")
        colindex4 = findrc(sh, "姓名")
        colindex5 = findrc(sh, "学年学期")

        if (not colindex1) or (not colindex2) or (not colindex3) or (not colindex4) or (not colindex5):
            log.config(state=NORMAL)
            log.insert(END, "文件%s: 未找到相关列\n" % p)
            log.config(state=DISABLED)
            continue

        for r in range(3, sh.nrows):
            v1 = unicode( sh.cell_value(r, colindex1[1]) )
            v2 = unicode( sh.cell_value(r, colindex2[1]) )
            v3 = unicode( sh.cell_value(r, colindex3[1]) )
            v4 = unicode( sh.cell_value(r, colindex4[1]) )
            v5 = unicode( sh.cell_value(r, colindex5[1]) )

            if (not v3) or (len(v3.strip())==0) or (not v4) or (len(v4.strip())==0):
                continue

            if (not v1) or (len(v1.strip())==0) or (not v2) or (len(v2.strip())==0) or (not v5) or (len(v5.strip())==0):
                log.config(state=NORMAL)
                log.insert(END, "文件%s, 表%s: 第%d行无课程名称或考试成绩或学年学期\n" % (p,sh.name,r+1))
                log.config(state=DISABLED)

    return