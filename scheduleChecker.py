__author__ = 'anjing'
#coding=utf-8

from xlrd import *
from Tkinter import *
from tkFileDialog import *
from tkMessageBox import *
import os
from ConfigParser import *

def readConfig():
    if not os.path.exists("config.ini"):
        cfgfile = open("config.ini", "w")

        cf = ConfigParser()
        cf.add_section('ExcelFilePath')
        cf.set('ExcelFilePath', 'path', "")
        cf.add_section('Field')

        cf.set('Field', 'Field_0', "")

        cf.add_section("Query")
        cf.set("Query", "key", "")

        cf.write(cfgfile)
        cfgfile.close()

        return

    cf = ConfigParser()
    cf.read("config.ini")

    filepath.set(cf.get("ExcelFilePath", "path"))

    for k, v in cf.items("Field"):
        if v.strip():
            field.insert(END, v)

    query.set(cf.get("Query", "key"))

def writeConfig():
    cfgfile = open("config.ini", "w")

    cf = ConfigParser()
    cf.add_section('ExcelFilePath')
    cf.set('ExcelFilePath', 'path', filepath.get())
    cf.add_section('Field')

    ft = field.get(0, END)

    for i in range(len(ft)):
        cf.set('Field', 'Field_'+str(i), ft[i].encode('utf-8'))

    cf.add_section("Query")
    cf.set("Query", "key", query.get())

    cf.write(cfgfile)
    cfgfile.close()

    sys.exit()

root = Tk()

root.title(u"课表核对")
root.resizable(width=False, height=False)

pathframe = LabelFrame(root, text=u"路径设置")

pathtip = Label(pathframe,text=u"文件路径: ")
pathtip.pack(side=LEFT)

filepath = StringVar()
pathinput = Entry(pathframe, width=50, textvariable=filepath)
pathinput.pack(side=LEFT, fill=X, expand=1, padx=2)

def filebrowse(target):
    target.set('') #清空entry里面的内容
    #调用filedialog模块的askdirectory()函数去打开文件夹
    fp = askdirectory()
    if fp:
        target.set(fp) #将选择好的路径加入到entry里面

pathbtn = Button(pathframe, text=u"浏览", command=lambda: filebrowse(filepath))
pathbtn.pack(side=LEFT, padx=2)

operateframe = Frame(root)

filedframe = LabelFrame(operateframe, text=u"内容选择")

rowframe = Frame(filedframe)
rowframe.pack(side=TOP)
rowlabel = Label(rowframe, text=u"核对信息: ")
rowlabel.pack(side=LEFT)
title = StringVar()
rowinput = Entry(rowframe, textvariable=title)
rowinput.pack(side=LEFT, fill=X, expand=1)

def insertfield(event):

    if not title.get().strip():
        return

    if not title.get().strip() in fieldname.get():
        field.insert(END, title.get())
        return

rowinput.bind("<Return>", insertfield)

fieldname = StringVar()
field = Listbox(filedframe, listvariable=fieldname, selectmode=MULTIPLE)
field.pack(side=TOP, fill=BOTH, expand=1)

rowbtn = Button(rowframe, text=u"添加", command=lambda: insertfield(""))
rowbtn.pack(side=LEFT)

queryframe = LabelFrame(operateframe, text=u"查询结果")

inputframe = Frame(queryframe)
inputframe.pack(side=TOP, fill=X, expand=1)
querylabel = Label(inputframe, text=u"查询对象: ")
querylabel.pack(side=LEFT)
query = StringVar()
queryinput = Entry(inputframe, textvariable=query)
queryinput.pack(side=LEFT, fill=X, expand=1)

queryresult = Text(queryframe)
queryresult.pack(side=TOP, fill=X, expand=1)

def queryprocess():
    queryresult.delete(0.0, END)

    queryresult.insert(END, query.get()+"\n==========================================\n")

    if not filepath:
        showwarning(u"错误", u"请输入路径！")
        return

    if not fieldname.get():
        showwarning(u"错误", u"请输入查询信息！")
        return

    result = dict()

    for item in field.get(0, END):
        result[item] = ""

    def findrc(sh, i):
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                v = unicode(sh.cell_value(r, c)).replace(' ','')
                if i in v:
                    return (r,c)

    for f in os.listdir(filepath.get()):
        if not (".xls" in f or ".xlsx" in f):
            continue

        p = filepath.get() + f

        xls = open_workbook(p)
        for x in range(xls.nsheets):
            sh = xls.sheets()[x]

            q = findrc(sh, query.get())
            if not q:
                continue

            row = q[0]

            for info in result.keys():
                result[info] = findrc(sh, info)

            for r in result.keys():
                if not result[r]:
                    continue

                print result
                result[r] = unicode(sh.cell_value(row, result[r][1]))

                queryresult.insert(END, result[r])
                queryresult.insert(END, "\t")

            queryresult.insert(END, "\n")

querybtn = Button(inputframe, text=u"查询", command=queryprocess)
querybtn.pack(side=LEFT)

filedframe.pack(side=LEFT, fill=BOTH, expand=1, anchor=W, padx=5)
queryframe.pack(side=RIGHT, anchor=E, padx=5)

pathframe.pack(side=TOP, fill=X, expand=1, padx=5, pady=2)
operateframe.pack(side=TOP, fill=X, expand=1, padx=5, pady=2)

root.wm_protocol("WM_DELETE_WINDOW", writeConfig)

readConfig()

root.mainloop()
